defmodule Exoffice.Parser.Excel2003.OLE do
  use Bitwise, only_operators: true

  defstruct binary: nil,
            num_big_block_depot_blocks: 0,
            root_start_block: nil,
            sbd_start_block: nil,
            extension_block: nil,
            num_extension_blocks: 0,
            big_block_chain: nil,
            small_block_chain: nil,
            entry: nil,
            document_summary_information: nil,
            summary_information: nil,
            props: [],
            workbook: nil,
            root_entry: nil

  @identifier_ole <<0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1>>

  # Size of a sector = 512 bytes
  @big_block_size                 0x200

  # Size of a short sector = 64 bytes
  @small_block_size               0x40

  # Size of a directory entry always = 128 bytes
  @property_storage_block_size    0x80

  # Minimum size of a standard stream = 4096 bytes, streams smaller than this are stored as short streams
  @small_block_threshold          0x1000

  # header offsets
  @num_big_block_depot_blocks_pos 0x2c
  @root_start_block_pos           0x30
  @small_block_depot_block_pos    0x3c
  @extension_block_pos            0x44
  @num_extension_block_pos        0x48
  @big_block_depot_blocks_pos     0x4c

  # property storage offsets (directory offsets)
  @size_of_name_pos               0x40
  @type_pos                       0x42
  @start_block_pos                0x74
  @size_pos                       0x78

  def identifier_ole, do: @identifier_ole

  def get_stream(_, _, -2, stream) do
    stream
  end

  def get_stream(ole, prop, block, stream) do
    case prop.size < @small_block_threshold do
      true ->
        root_data = read_data(
          ole.binary,
          "",
          Enum.at(ole.props, ole.root_entry).start_block,
          ole.big_block_chain
        )
        pos = block * @small_block_size
        new_stream = stream <> binary_part(root_data, pos, @small_block_size)
        get_stream(ole, prop, get_int_4d(ole.small_block_chain, block * 4), new_stream)
      false ->
        num_blocks = case rem(prop.size, @big_block_size) == 0 do
          true -> prop.size / @big_block_size
          false -> prop.size / @big_block_size + 1
        end
        case num_blocks == 0 do
          true -> ""
          false ->
            pos = (block + 1) * @big_block_size
            new_stream = stream <> binary_part(ole.binary, pos, @big_block_size)
            get_stream(ole, prop, get_int_4d(ole.big_block_chain, block * 4), new_stream)
        end
    end
  end

  @doc """
  Initial parse of binary data. Returns OLE struct with parsed parts
  """
  def parse_blocks(binary) do
    # TODO: Rewrite this to use pipe operator and return {:error, reason} in case of any failure

    # Total number of sectors used for the SAT
    num_big_block_depot_blocks = get_int_4d(binary, @num_big_block_depot_blocks_pos)

    # SecID of the first sector of the directory stream
    root_start_block = get_int_4d(binary, @root_start_block_pos)

    # SecID of the first sector of the SSAT (or -2 if not extant)
    sbd_start_block = get_int_4d(binary, @small_block_depot_block_pos)

    # SecID of the first sector of the MSAT (or -2 if no additional sectors are used)
    extension_block = get_int_4d(binary, @extension_block_pos)

    # Total number of sectors used by MSAT
    num_extension_blocks = get_int_4d(binary, @num_extension_block_pos)

    bbd_blocks = case num_extension_blocks != 0 do
      true -> (@big_block_size - @big_block_depot_blocks_pos) / 4 |> round
      false -> num_big_block_depot_blocks
    end

    {big_block_depot_blocks, _} = case bbd_blocks do
      0 -> {[], @big_block_depot_blocks_pos}
      _ -> 0..bbd_blocks - 1
        |> Enum.reduce({[], @big_block_depot_blocks_pos}, fn _, {acc, pos} ->
          {acc ++ [get_int_4d(binary, pos)], pos + 4}
        end)
    end

    big_block_depot_blocks = case num_extension_blocks do
      0 -> big_block_depot_blocks
      _ -> [_, _, big_block_depot_blocks, _] = 0..num_extension_blocks - 1
        |> Enum.reduce([extension_block, bbd_blocks, big_block_depot_blocks, (extension_block + 1) * @big_block_size], fn _, [extension_block, bbd_blocks, big_block_depot_blocks, pos] ->
          blocks_to_read = min(num_big_block_depot_blocks - bbd_blocks, @big_block_size / 4 - 1) |> round

          {big_block_depot_blocks, pos} = case bbd_blocks + blocks_to_read do
            0 -> {big_block_depot_blocks, pos}
            _ ->
              Enum.reduce(bbd_blocks..bbd_blocks + blocks_to_read, {big_block_depot_blocks, pos}, fn _, {acc, pos} ->
                {acc ++ [get_int_4d(binary, pos)], pos + 4}
              end)
          end

          big_block_depot_blocks = Enum.slice(big_block_depot_blocks, 0, Enum.count(big_block_depot_blocks) - 1)

          bbd_blocks = bbd_blocks + blocks_to_read
          new_extension_block = if bbd_blocks < num_big_block_depot_blocks, do: get_int_4d(binary, pos), else: extension_block

          [new_extension_block, bbd_blocks, big_block_depot_blocks, pos]
        end)
        big_block_depot_blocks
    end

    bbs = @big_block_size / 4
    big_block_chain = case num_big_block_depot_blocks do
      0 -> ""
      _ ->
        {_, big_block_chain} = 0..num_big_block_depot_blocks - 1
        |> Enum.reduce({0, ""}, fn i, {pos, big_block_chain} ->
          new_pos = (Enum.at(big_block_depot_blocks, i) + 1) * @big_block_size
          {new_pos + @big_block_size, big_block_chain <> binary_part(binary, new_pos, @big_block_size)}
        end)
        big_block_chain
    end

    small_block_chain = get_sbd_block(binary, big_block_chain, "", sbd_start_block, 0)
    entry = read_data(binary, "", root_start_block, big_block_chain)

    loader = %__MODULE__{
      binary: binary,
      num_big_block_depot_blocks: num_big_block_depot_blocks,
      root_start_block: root_start_block,
      sbd_start_block: sbd_start_block,
      extension_block: extension_block,
      num_extension_blocks: num_extension_blocks,
      big_block_chain: big_block_chain,
      small_block_chain: small_block_chain,
      entry: entry
    }

    {:ok, read_property_sets(0, loader)}
  end

  def read_property_sets(nil, loader) do
    loader
  end

  def read_property_sets(offset, loader) do
    #  loop through entires, each entry is 128 bytes
    entry_len = byte_size(loader.entry)
    case offset < entry_len do
      true ->
        # entry data (128 bytes)
        d = binary_part(loader.entry, offset, @property_storage_block_size)

        # size in bytes of name
        name_size = decoded_binary_at(d, @size_of_name_pos) ||| (decoded_binary_at(d, @size_of_name_pos + 1) <<< 8)

        # type of entry
        type = decoded_binary_at(d, @type_pos)

        # sectorID of first sector or short sector, if this entry refers to a stream (the case with workbook)
        # sectorID of first sector of the short-stream container stream, if this entry is root entry
        start_block = get_int_4d(d, @start_block_pos)

        size = get_int_4d(d, @size_pos)
        name = String.replace(binary_part(d, 0, name_size), "\x00", "")

        props = %{
          name: name,
          type: type,
          start_block: start_block,
          size: size
        }
        loader = %{loader | props: loader.props ++ [props]}

        # tmp helper to simplify checks
        up_name = String.upcase(name)

        # Workbook directory entry (BIFF5 uses Book, BIFF8 uses Workbook)
        {workbook, root_entry} = cond do
          up_name == "WORKBOOK" || up_name == "BOOK" ->
            {Enum.count(loader.props) - 1, loader.root_entry}
          up_name == "ROOT ENTRY" || up_name == "R" ->
            {loader.workbook, Enum.count(loader.props) - 1}
          true -> {loader.workbook, loader.root_entry}
        end

        # Summary information
        summary_information = case name == <<5>> <> "SummaryInformation" do
          true -> Enum.count(loader.props) - 1
          false -> loader.summary_information
        end

        # Additional Document Summary information
        document_summary_information = case name == <<5>> <> "DocumentSummaryInformation" do
          true -> Enum.count(loader.props) - 1
          false -> loader.document_summary_information
        end
        loader = %{loader |
          summary_information: summary_information,
          document_summary_information: document_summary_information,
          workbook: workbook,
          root_entry: root_entry,
        }
        read_property_sets(offset + @property_storage_block_size, loader)
      false -> read_property_sets(nil, loader)
    end
  end

  defp read_data(_, data, -2, _) do
    data
  end

  defp read_data(binary, data, block, big_block_chain) do
    pos = (block + 1) * @big_block_size
    new_data = data <> binary_part(binary, pos, @big_block_size)
    new_block = get_int_4d(big_block_chain, block * 4)
    read_data(binary, new_data, new_block, big_block_chain)
  end

  defp get_sbd_block(_, _, small_block_chain, -2, _) do
    small_block_chain
  end

  defp get_sbd_block(binary, big_block_chain, small_block_chain, sbd_block, pos) do
      pos = (sbd_block + 1) * @big_block_size
      small_block_chain = small_block_chain <> binary_part(binary, pos, @big_block_size)
      pos = pos + @big_block_size
      sbd_block = get_int_4d(big_block_chain, sbd_block * 4)
      get_sbd_block(binary, big_block_chain, small_block_chain, sbd_block, pos)
  end

  def get_int_4d(binary, pos) do
    or_24 = decoded_binary_at(binary, pos + 3)
    ord_24 = case or_24 >= 128 do
      #  negative number
      true ->
        -abs((256 - or_24) <<< 24)
      false ->
        (or_24 &&& 127) <<< 24
    end

    decoded_binary_at(binary, pos) ||| (decoded_binary_at(binary, pos + 1) <<< 8) ||| (decoded_binary_at(binary, pos + 2) <<< 16) ||| ord_24
  end

  def get_int_2d(binary, pos) do
    decoded_binary_at(binary, pos) ||| (decoded_binary_at(binary, pos + 1) <<< 8)
  end

  def decoded_binary_at(binary, pos) do
    binary
    |> binary_part(pos, 1)
    |> :binary.decode_unsigned
  end

end
