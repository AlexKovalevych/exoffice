defmodule Exoffice.Parser.Excel2007 do

  @behaviour Exoffice.Parser

  def extensions, do: [".xlsx"]

  def parse(path, options \\ []) do
    Xlsxir.multi_extract(path)
  end

  def parse_sheet(path, index, options \\ []) do
    Xlsxir.extract(path, index, options)
  end

  def count_rows(pid) do
    Xlsxir.get_multi_info(pid, :rows)
  end

  def get_rows(pid) do
    Stream.map(1..count_rows(pid), fn row ->
      Xlsxir.get_row(pid, row)
    end)
  end

  def close(pid) do
    Xlsxir.close(pid)
  end

end
