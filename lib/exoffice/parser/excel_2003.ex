defmodule Exoffice.Parser.Excel2003 do
  alias Exoffice.Parser.Excel2003.Loader

  defstruct sheets: [],
            encryption: 0,
            version: nil,
            data_size: 0,
            codepage: nil,
            base_date: 1900,
            pids: []

  @behaviour Exoffice.Parser

  def extensions, do: [".xls"]

  def parse(path, options \\ []) do
    Loader.load(path)
  end

  def parse_sheet(path, index, options \\ []) do
    Loader.load(path, index)
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
