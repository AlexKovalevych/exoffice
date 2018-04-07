defmodule Exoffice.Parser.Excel2003.String do
  alias Exoffice.Parser.Excel2003.OLE
  use Bitwise, only_operators: true

  defstruct [:value, :size]

  def read_unicode_string_short(data) do
    # offset: 0: size: 1; length of the string (character count)
    character_count = OLE.decoded_binary_at(data, 0)

    string = binary_part(data, 1, byte_size(data) - 1) |> read_unicode_string(character_count)

    # add 1 for the string length
    %{string | size: string.size + 1}
  end

  def read_unicode_string(data, character_count) do
    # offset: 0: size: 1; option flags
    # bit: 0; mask: 0x01; character compression (0 = compressed 8-bit, 1 = uncompressed 16-bit)
    is_compressed = !((0x01 &&& OLE.decoded_binary_at(data, 0)) >>> 0)

    # bit: 2; mask: 0x04; Asian phonetic settings
    has_asian = 0x04 &&& OLE.decoded_binary_at(data, 0) >>> 2

    # bit: 3; mask: 0x08; Rich-Text settings
    has_rich_text = 0x08 &&& OLE.decoded_binary_at(data, 0) >>> 3

    # offset: 1: size: var; character array
    # this offset assumes richtext and Asian phonetic settings are off which is generally wrong
    # needs to be fixed
    length = if is_compressed, do: character_count * 2, else: character_count
    value = encode_utf_16(binary_part(data, 1, length), is_compressed)
    size = if is_compressed, do: 1 + character_count, else: 1 + 2 * character_count

    %__MODULE__{
      value: value,
      size: size
    }
  end

  def encode_utf_16(string, false) do
    convert_encoding(string, "UTF-16LE", "UTF-8")
  end

  def encode_utf_16(string, true) do
    string
    |> uncompress_byte_string
    |> encode_utf_16(false)
  end

  defp uncompress_byte_string(string) do
    str_len = byte_size(string)

    Enum.reduce(0..(str_len - 1), <<>>, fn i, acc ->
      acc <> binary_part(string, i, 1) <> "\0"
    end)
  end

  def read_byte_string_short(string, codepage) do
    # offset: 0; size: 1; length of the string (character count)
    ln = OLE.decoded_binary_at(string, 0)

    # offset: 1: size: var; character array (8-bit characters)
    value = decode_codepage(binary_part(string, 1, ln), codepage)

    %__MODULE__{
      value: value,
      # size in bytes of data structure
      size: 1 + ln
    }
  end

  defp decode_codepage(string, codepage) do
    convert_encoding(string, codepage, "UTF-8")
  end

  def convert_encoding(value, from, to) do
    :iconv.convert(from, to, value)
  end

  defp decode_utf_16(str, bom_be \\ true) do
    case byte_size(str) < 2 do
      true ->
        str

      false ->
        c0 = OLE.decoded_binary_at(str, 0)
        c1 = OLE.decoded_binary_at(str, 1)

        {str, bom_be} =
          cond do
            c0 == 0xFE && c1 == 0xFF -> {binary_part(str, 2, byte_size(str) - 2), bom_be}
            c0 == 0xFF && c1 == 0xFE -> {binary_part(str, 2, byte_size(str) - 2), false}
            true -> {str, bom_be}
          end

        len = byte_size(str)

        Enum.reduce(0..(len - 1), "", fn i, acc ->
          val =
            case bom_be do
              true -> (OLE.decoded_binary_at(str, i) <<< 4) <> OLE.decoded_binary_at(str, i + 1)
              false -> (OLE.decoded_binary_at(str, i + 1) <<< 4) <> OLE.decoded_binary_at(str, i)
            end

          acc <> if val == 0x228, do: "\n", else: val
        end)
    end
  end
end
