defmodule Exoffice.Parser.Excel2003.Cell do
  def string_from_column_index(p_column_index \\ 0) do
    # Determine column string
    cond do
      p_column_index < 26 -> <<65 + p_column_index>>
      p_column_index < 702 -> <<64 + Float.floor(p_column_index / 26)>> <> <<65 + rem(p_column_index, 26)>>
      true -> <<64 + Float.floor((p_column_index - 26) / 676)>> <> <<65 + Float.floor(rem(p_column_index - 26, 676) / 26)>> <> <<65 + rem(p_column_index, 26)>>
    end
  end
end
