defmodule Exoffice.Parser.Excel2003.Loader do
  alias Exoffice.Parser.Excel2003
  alias Exoffice.Parser.Excel2003.OLE
  alias Exoffice.Parser.Excel2003.Cell
  alias Exoffice.Parser.Excel2003.String, as: ExofficeString
  alias Exoffice.Parser.Excel2003.Index
  alias Xlsxir.StateManager
  use Bitwise, only_operators: true

  # ParseXL definitions
  @xls_biff8 0x0600
  @xls_biff7 0x0500
  @xls_workbook_globals 0x0005
  @xls_worksheet 0x0010

  # Calendar
  # Base date of 1st Jan 1900 = 1.0
  @calendar_windows_1900 1900
  # Base date of 2nd Jan 1904 = 1.0
  @calendar_mac_1904 1904

  # record identifiers
  @xls_type_sheet 0x0085
  @xls_type_bof 0x0809
  @xls_type_codepage 0x0042
  @xls_type_datemode 0x0022
  @xls_type_sst 0x00FC
  @xls_type_continue 0x003C
  @xls_type_labelsst 0x00FD
  @xls_type_number 0x0203
  @xls_type_blank 0x0201
  @xls_type_eof 0x000A

  # sheet state
  @sheetstate_visible "visible"
  @sheetstate_hidden "hidden"
  @sheetstate_veryhidden "veryHidden"

  defstruct data: nil,
            summary_information: nil,
            document_summary_information: nil,
            sst_tid: nil

  def load(path, sheet \\ nil) do
    with {:ok, file} <- File.open(path, [:read, :binary]),
         {:ok, ole} <- :file.read(file, 8),
         true <- ole == OLE.identifier_ole(),
         {:ok, binary} <- File.read(path),
         {:ok, ole} <- OLE.parse_blocks(binary),
         loader <- get_stream(ole),
         {stream, _pos, excel} <- parse(loader, 0, %Excel2003{data_size: byte_size(loader.data)}),
         pids = parse_sheets(loader, excel, sheet) do
      Enum.map(pids, fn {status, pid, _} -> {status, pid} end)
    else
      {:error, reason} -> {:error, reason}
    end
  end

  defp get_stream(ole) do
    [data, summary_information, document_summary_information] =
      Enum.map(
        [
          ole.workbook,
          ole.summary_information,
          ole.document_summary_information
        ],
        fn
          nil ->
            nil

          prop_name ->
            prop = Enum.at(ole.props, prop_name)
            OLE.get_stream(ole, prop, prop.start_block, "")
        end
      )

    Index.new()

    %__MODULE__{
      data: data,
      summary_information: summary_information,
      document_summary_information: document_summary_information,
      sst_tid: GenServer.call(StateManager, :new_table)
    }
  end

  defp parse_sheets(%__MODULE__{sst_tid: sst_tid} = loader, excel, sheet) do
    sheets = if is_nil(sheet), do: excel.sheets, else: [Enum.at(excel.sheets, sheet)]

    pids =
      sheets
      |> Enum.filter(fn %{sheet_type: type} ->
        type == <<0>>
      end)
      |> Enum.map(fn sheet ->
        table_id = GenServer.call(StateManager, :new_table)
        parse_sheet_part(loader, sheet.offset, excel, table_id)
        {:ok, table_id, excel}
      end)

    Index.del()
    :ets.delete(sst_tid)
    pids
  end

  def parse_sheet_part(loader, pos, excel), do: {loader, pos, excel}

  def parse_sheet_part(%__MODULE__{data: stream} = loader, pos, excel, table_id) when byte_size(stream) > pos do
    code = OLE.get_int_2d(stream, pos)

    case code do
      @xls_type_bof -> read_bof(loader, pos, excel, :parse_sheet_part, table_id)
      @xls_type_labelsst -> read_label_sst(loader, pos, excel, table_id)
      @xls_type_number -> read_number(loader, pos, excel, table_id)
      @xls_type_blank -> read_blank(loader, pos, excel, table_id)
      @xls_type_eof -> read_default(loader, pos, excel, :parse_sheet_part, nil)
      _ -> read_default(loader, pos, excel, :parse_sheet_part, table_id)
    end
  end

  def parse_sheet_part(%__MODULE__{} = loader, pos, excel, _table_id) do
    {loader, pos, excel}
  end

  def parse(%__MODULE__{data: stream} = loader, pos, excel) when byte_size(stream) - 4 > pos do
    code = OLE.get_int_2d(stream, pos)

    case code do
      @xls_type_bof -> read_bof(loader, pos, excel)
      @xls_type_sheet -> read_sheet(loader, pos, excel)
      @xls_type_codepage -> read_codepage(loader, pos, excel)
      @xls_type_datemode -> read_datemode(loader, pos, excel)
      @xls_type_sst -> read_sst(loader, pos, excel)
      _ -> read_default(loader, pos, excel)
    end
  end

  def parse(%__MODULE__{} = loader, pos, excel) do
    {loader, pos, excel}
  end

  defp read_bof(%__MODULE__{data: stream} = loader, pos, excel, fun \\ :parse, pid \\ nil) do
    length = OLE.get_int_2d(stream, pos + 2)
    record_data = binary_part(stream, pos + 4, length)

    new_pos = pos + length + 4

    # offset: 2; size: 2; type of the following data
    substream_type = OLE.get_int_2d(record_data, 2)

    case substream_type do
      @xls_workbook_globals ->
        version = OLE.get_int_2d(record_data, 0)

        if version != @xls_biff8 && version != @xls_biff7 do
          {:error, "Cannot read this Excel file. Version is too old."}
        else
          parse(loader, new_pos, %{excel | version: version})
        end

      @xls_worksheet ->
        # do not use this version information for anything
        # it is unreliable (OpenOffice doc, 5.8), use only version information from the global stream
        apply(__MODULE__, fun, if(pid, do: [loader, new_pos, excel, pid], else: [loader, new_pos, excel]))

      _ ->
        # substream, e.g. chart
        # just skip the entire substream
        read_bof_default(loader, new_pos, excel)
    end
  end

  defp read_label_sst(%__MODULE__{data: stream, sst_tid: sst_tid} = loader, pos, excel, pid) do
    length = OLE.get_int_2d(stream, pos + 2)
    record_data = read_record_data(stream, pos + 4, length)

    # offset: 0; size: 2; index to row
    row = OLE.get_int_2d(record_data, 0) + 1

    # offset: 2; size: 2; index to column
    column = OLE.get_int_2d(record_data, 2)
    column_string = Cell.string_from_column_index(column)

    # offset: 6; size: 4; index to SST record
    index = OLE.get_int_4d(record_data, 6)

    value =
      :ets.lookup(sst_tid, index)
      |> List.first()
      |> Tuple.to_list()
      |> Enum.at(1)

    # add cell
    case :ets.match(pid, {row, :"$1"}) do
      [[cells]] ->
        :ets.insert(pid, {row, cells ++ [[column_string <> to_string(row), value]]})

      _ ->
        :ets.insert(pid, {row, [[column_string <> to_string(row), value]]})
    end

    parse_sheet_part(loader, pos + 4 + length, excel, pid)
  end

  defp read_number(%__MODULE__{data: stream} = loader, pos, excel, pid) do
    length = OLE.get_int_2d(stream, pos + 2)
    record_data = read_record_data(stream, pos + 4, length)

    # offset: 0; size: 2; index to row
    row = OLE.get_int_2d(record_data, 0) + 1

    # offset: 2; size 2; index to column
    column = OLE.get_int_2d(record_data, 2)
    column_string = Cell.string_from_column_index(column)

    value = extract_number(binary_part(record_data, 6, 8))

    # add cell
    case :ets.match(pid, {row, :"$1"}) do
      [[cells]] ->
        :ets.insert(pid, {row, cells ++ [[column_string <> to_string(row), value]]})

      _ ->
        :ets.insert(pid, {row, [[column_string <> to_string(row), value]]})
    end

    parse_sheet_part(loader, pos + 4 + length, excel, pid)
  end

  defp read_blank(%__MODULE__{data: stream} = loader, pos, excel, pid) do
    length = OLE.get_int_2d(stream, pos + 2)
    record_data = read_record_data(stream, pos + 4, length)

    # offset: 0; size: 2; row index
    row = OLE.get_int_2d(record_data, 0) + 1

    # offset: 2; size: 2; col index
    column = OLE.get_int_2d(record_data, 2)
    column_string = Cell.string_from_column_index(column)

    # add cell
    case :ets.match(pid, {row, :"$1"}) do
      [[cells]] ->
        :ets.insert(pid, {row, cells ++ [[column_string <> to_string(row), nil]]})

      _ ->
        :ets.insert(pid, {row, [[column_string <> to_string(row), nil]]})
    end

    parse_sheet_part(loader, pos + 4 + length, excel, pid)
  end

  defp extract_number(data) do
    rknumhigh = OLE.get_int_4d(data, 4)
    rknumlow = OLE.get_int_4d(data, 0)
    sign = (rknumhigh &&& 0x80000000) >>> 31
    exp = ((rknumhigh &&& 0x7FF00000) >>> 20) - 1023
    mantissa = 0x100000 ||| (rknumhigh &&& 0x000FFFFF)
    mantissa_low1 = (rknumlow &&& 0x80000000) >>> 31
    mantissa_low2 = rknumlow &&& 0x7FFFFFFF

    value =
      if 20 - exp > 1023 do
        0
      else
        mantissa / :math.pow(2, 20 - exp)
      end

    value
    |> (fn v ->
          if mantissa_low1 != 0 && 21 - exp <= 1023 do
            v + 1 / :math.pow(2, 21 - exp)
          else
            v
          end
        end).()
    |> (fn v ->
          if 52 - exp > 1023 do
            v
          else
            v + mantissa_low2 / :math.pow(2, 52 - exp)
          end
        end).()
    |> (fn v -> if sign != 0, do: v * -1, else: v end).()
  end

  defp read_codepage(%__MODULE__{data: stream} = loader, pos, excel) do
    length = OLE.get_int_2d(stream, pos + 2)
    record_data = read_record_data(stream, pos + 4, length)

    # offset: 0; size: 2; code page identifier
    case OLE.get_int_2d(record_data, 0) |> codepage_to_name do
      {:ok, codepage} -> parse(loader, pos + length + 4, %{excel | codepage: codepage})
      {:error, reason} -> {:error, reason}
    end
  end

  defp codepage_to_name(codepage) do
    case codepage do
      # ASCII
      367 ->
        {:ok, "ASCII"}

      # OEM US
      437 ->
        {:ok, "CP437"}

      # OEM Arabic
      720 ->
        {:error, "Code page 720 not supported."}

      # OEM Greek
      737 ->
        {:ok, "CP737"}

      # OEM Baltic
      775 ->
        {:ok, "CP775"}

      # OEM Latin I
      850 ->
        {:ok, "CP850"}

      # OEM Latin II (Central European)
      852 ->
        {:ok, "CP852"}

      # OEM Cyrillic
      855 ->
        {:ok, "CP855"}

      # OEM Turkish
      857 ->
        {:ok, "CP857"}

      # OEM Multilingual Latin I with Euro
      858 ->
        {:ok, "CP858"}

      # OEM Portugese
      860 ->
        {:ok, "CP860"}

      # OEM Icelandic
      861 ->
        {:ok, "CP861"}

      # OEM Hebrew
      862 ->
        {:ok, "CP862"}

      # OEM Canadian (French)
      863 ->
        {:ok, "CP863"}

      # OEM Arabic
      864 ->
        {:ok, "CP864"}

      # OEM Nordic
      865 ->
        {:ok, "CP865"}

      # OEM Cyrillic (Russian)
      866 ->
        {:ok, "CP866"}

      # OEM Greek (Modern)
      869 ->
        {:ok, "CP869"}

      # ANSI Thai
      874 ->
        {:ok, "CP874"}

      # ANSI Japanese Shift-JIS
      932 ->
        {:ok, "CP932"}

      # ANSI Chinese Simplified GBK
      936 ->
        {:ok, "CP936"}

      # ANSI Korean (Wansung)
      949 ->
        {:ok, "CP949"}

      # ANSI Chinese Traditional BIG5
      950 ->
        {:ok, "CP950"}

      # UTF-16 (BIFF8)
      1200 ->
        {:ok, "UTF-16LE"}

      # ANSI Latin II (Central European)
      1250 ->
        {:ok, "CP1250"}

      # ANSI Cyrillic
      1251 ->
        {:ok, "CP1251"}

      # CodePage is not always correctly set when the xls file was saved by Apple's Numbers program
      0 ->
        {:ok, "CP1252"}

      # ANSI Latin I (BIFF4-BIFF7)
      1252 ->
        {:ok, "CP1252"}

      # ANSI Greek
      1253 ->
        {:ok, "CP1253"}

      # ANSI Turkish
      1254 ->
        {:ok, "CP1254"}

      # ANSI Hebrew
      1255 ->
        {:ok, "CP1255"}

      # ANSI Arabic
      1256 ->
        {:ok, "CP1256"}

      # ANSI Baltic
      1257 ->
        {:ok, "CP1257"}

      # ANSI Vietnamese
      1258 ->
        {:ok, "CP1258"}

      # ANSI Korean (Johab)
      1361 ->
        {:ok, "CP1361"}

      # Apple Roman
      10000 ->
        {:ok, "MAC"}

      # Macintosh Japanese
      10001 ->
        {:ok, "CP932"}

      # Macintosh Chinese Traditional
      10002 ->
        {:ok, "CP950"}

      # Macintosh Korean
      10003 ->
        {:ok, "CP1361"}

      # Macintosh Greek
      10006 ->
        {:ok, "MACGREEK"}

      #  Macintosh Cyrillic
      10007 ->
        {:ok, "MACCYRILLIC"}

      # Macintosh - Simplified Chinese (GB 2312)
      10008 ->
        {:ok, "CP936"}

      # Macintosh Central Europe
      10029 ->
        {:ok, "MACCENTRALEUROPE"}

      # Macintosh Icelandic
      10079 ->
        {:ok, "MACICELAND"}

      # Macintosh Turkish
      10081 ->
        {:ok, "MACTURKISH"}

      # UTF-16 (BIFF8) This isn't correct, but some Excel writer libraries erroneously use Codepage 21010 for UTF-16LE
      21010 ->
        {:ok, "UTF-16LE"}

      # Apple Roman
      32768 ->
        {:ok, "MAC"}

      # ANSI Latin I (BIFF2-BIFF3)
      32769 ->
        {:error, "Code page 32769 not supported."}

      # Unicode (UTF-7)
      65000 ->
        {:ok, "UTF-7"}

      # Unicode (UTF-8)
      65001 ->
        {:ok, "UTF-8"}

      _ ->
        {:error, "Unknown codepage: " <> codepage}
    end
  end

  defp read_bof_default(%__MODULE__{data: stream} = loader, pos, excel) do
    code = OLE.get_int_2d(stream, pos)
    length = OLE.get_int_2d(stream, pos + 2)

    case code != @xls_type_eof && pos < excel.data_size do
      true -> read_bof_default(loader, pos + length + 4, excel)
      false -> excel
    end
  end

  defp read_datemode(%__MODULE__{data: stream} = loader, pos, excel) do
    length = OLE.get_int_2d(stream, pos + 2)
    record_data = read_record_data(stream, pos + 4, length)

    # offset: 0; size: 2; 0 = base 1900, 1 = base 1904
    excel = if binary_part(record_data, 0, 1) == <<1>>, do: %{excel | base_date: @calendar_mac_1904}, else: excel
    parse(loader, pos + length + 4, excel)
  end

  defp read_sheet(%__MODULE__{data: stream} = loader, pos, excel) do
    length = OLE.get_int_2d(stream, pos + 2)
    record_data = read_record_data(stream, pos + 4, length)

    # offset: 0; size: 4; absolute stream position of the BOF record of the sheet
    # NOTE: not encrypted
    rec_offset = OLE.get_int_4d(stream, pos + 4)

    # offset: 4; size: 1; sheet state
    sheet_state =
      case binary_part(record_data, 4, 0) do
        <<1>> -> @sheetstate_hidden
        <<2>> -> @sheetstate_veryhidden
        _ -> @sheetstate_visible
      end

    # offset: 5; size: 1; sheet type
    sheet_type = binary_part(record_data, 5, 1)

    # offset: 6; size: var; sheet name
    rec_name =
      case excel.version do
        @xls_biff8 ->
          binary_part(record_data, 6, byte_size(record_data) - 6)
          |> ExofficeString.read_unicode_string_short()

        @xls_biff7 ->
          binary_part(record_data, 6, byte_size(record_data) - 6)
          |> ExofficeString.read_byte_string_short(excel.codepage)
      end

    sheet = %{name: rec_name.value, offset: rec_offset, sheet_state: sheet_state, sheet_type: sheet_type}
    parse(loader, pos + length + 4, %{excel | sheets: excel.sheets ++ [sheet]})
  end

  defp get_spliced_record_data(stream, pos, splice_offsets, data, i, @xls_type_continue) do
    # offset: 2; size: 2; length
    length = OLE.get_int_2d(stream, pos + 2)
    data = data <> read_record_data(stream, pos + 4, length)
    splice_offsets = splice_offsets ++ [Enum.at(splice_offsets, i - 1) + length]
    new_pos = pos + length + 4

    get_spliced_record_data(stream, new_pos, splice_offsets, data, i + 1, OLE.get_int_2d(stream, new_pos))
  end

  defp get_spliced_record_data(_, pos, splice_offsets, data, _, _) do
    {data, splice_offsets, pos}
  end

  defp read_sst(%__MODULE__{data: stream, sst_tid: sst_tid} = loader, pos, excel) do
    Index.new()

    # get spliced record data
    {record_data, splice_offsets, pos} = get_spliced_record_data(stream, pos, [0], <<>>, 1, @xls_type_continue)

    nm = OLE.get_int_4d(record_data, 4)

    0..(nm - 1)
    |> Stream.scan(8, fn _, pos ->
      {num_chars, pos} = {OLE.get_int_2d(record_data, pos), pos + 2}
      {option_flags, pos} = {OLE.decoded_binary_at(record_data, pos), pos + 1}

      # bit: 0; mask: 0x01; 0 = compressed; 1 = uncompressed
      is_compressed = (option_flags &&& 0x01) == 0

      # bit: 2; mask: 0x02; 0 = ordinary; 1 = Asian phonetic
      has_asian = (option_flags &&& 0x04) != 0

      # bit: 3; mask: 0x03; 0 = ordinary; 1 = Rich-Text
      has_rich_text = (option_flags &&& 0x08) != 0

      # number of Rich-Text formatting runs
      {formatting_runs, pos} =
        case has_rich_text do
          true -> {OLE.get_int_2d(record_data, pos), pos + 2}
          false -> {nil, pos}
        end

      # size of Asian phonetic setting
      {extended_run_length, pos} =
        case has_asian do
          true -> {OLE.get_int_2d(record_data, pos), pos + 4}
          false -> {nil, pos}
        end

      len = if is_compressed, do: num_chars, else: num_chars * 2

      limit_pos = Enum.drop_while(splice_offsets, &(pos > &1)) |> List.first()

      {ret_str, is_compressed, pos} =
        case pos + len <= limit_pos do
          true ->
            {binary_part(record_data, pos, len), is_compressed, pos + len}

          false ->
            # character array is split between records

            # first part of character array
            ret_str = binary_part(record_data, pos, limit_pos - pos)

            bytes_read = limit_pos - pos

            # remaining characters in Unicode string
            chars_left = num_chars - if is_compressed, do: bytes_read, else: bytes_read / 2

            pos = limit_pos

            get_ret_str(record_data, splice_offsets, ret_str, chars_left, pos, is_compressed)
        end

      # convert to UTF-8
      ret_str = ExofficeString.encode_utf_16(ret_str, is_compressed)

      # read additional Rich-Text information, if any
      {_fmt_runs, pos} =
        case has_rich_text do
          true ->
            # list of formatting runs
            fmt_runs =
              Enum.reduce(0..(formatting_runs - 1), fn i, acc ->
                # first formatted character; zero-based
                char_pos = OLE.get_int_2d(record_data, pos + i * 4)

                # index to font record
                font_index = OLE.get_int_2d(record_data, pos + 2 + i * 4)

                acc ++ [[char_pos, font_index]]
              end)

            {fmt_runs, pos + 4 * formatting_runs}

          false ->
            {[], pos}
        end

      pos = if has_asian, do: pos + extended_run_length, else: pos

      :ets.insert(sst_tid, {Index.get(), ret_str})
      Index.inc()
      pos
    end)
    |> Enum.into([])

    parse(loader, pos, excel)
  end

  defp get_ret_str(record_data, splice_offsets, ret_str, chars_left, pos, is_compressed) when chars_left > 0 do
    # look up next limit position, in case the string span more than one continue record
    limit_pos = Enum.drop_while(splice_offsets, &(pos >= &1)) |> List.first()

    # repeated option flags
    # OpenOffice.org documentation 5.21
    {option, pos} = {OLE.decoded_binary_at(record_data, pos), pos + 1}

    {ret_str, chars_left, is_compressed, len} =
      cond do
        is_compressed && option == 0 ->
          # 1st fragment compressed
          # this fragment compressed
          len = min(chars_left, limit_pos - pos) |> round
          {ret_str <> binary_part(record_data, pos, len), chars_left - len, true, len}

        !is_compressed && option != 0 ->
          # 1st fragment uncompressed
          # this fragment uncompressed
          len = min(chars_left * 2, limit_pos - pos) |> round
          {ret_str <> binary_part(record_data, pos, len), round(chars_left - len / 2), false, len}

        !is_compressed && option == 0 ->
          # 1st fragment uncompressed
          # this fragment compressed
          len = min(chars_left, limit_pos - pos) |> round

          ret_str =
            Enum.reduce(0..(len - 1), ret_str, fn i, acc ->
              acc <> binary_part(record_data, pos + i, 1) <> <<0>>
            end)

          {ret_str, chars_left - len, false, len}

        true ->
          # 1st fragment compressed
          # this fragment uncompressed
          ret_str =
            Enum.reduce(0..(byte_size(ret_str) - 1), "", fn i, acc ->
              acc <> binary_part(ret_str, i, 1) <> <<0>>
            end)

          len = min(chars_left * 2, limit_pos - pos) |> round
          ret_str = ret_str <> binary_part(record_data, pos, len)
          {ret_str, round(chars_left - len / 2), false, len}
      end

    get_ret_str(record_data, splice_offsets, ret_str, chars_left, pos + len, is_compressed)
  end

  defp get_ret_str(_record_data, _splice_offsets, ret_str, _, pos, is_compressed) do
    {ret_str, is_compressed, pos}
  end

  defp read_record_data(binary, pos, length) do
    # Encryption is not supported
    binary_part(binary, pos, length)
  end

  defp read_default(%__MODULE__{data: stream} = loader, pos, excel, fun \\ :parse, pid \\ nil) do
    length = OLE.get_int_2d(stream, pos + 2)

    # move stream pointer to next record
    new_pos = pos + length + 4
    apply(__MODULE__, fun, if(pid, do: [loader, new_pos, excel, pid], else: [loader, new_pos, excel]))
  end
end
