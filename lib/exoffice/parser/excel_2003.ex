defmodule Exoffice.Parser.Excel2003 do
  alias Exoffice.Parser.Excel2003.Loader

  defstruct sheets: [],
            pid: nil,
            encryption: 0,
            version: nil,
            data_size: 0,
            codepage: nil,
            base_date: 1900,
            pids: []

  @behaviour Exoffice.Parser

  def extensions, do: [".xls"]

  @doc """
  ## Example
  Parse file `test.xls` in `./test/test_data`:

  iex> [{:ok, {pid, table_id1}}, {:ok, {pid, table_id2}}] = Exoffice.Parser.Excel2003.parse("./test/test_data/test.xls")
  iex> Enum.member?(:ets.all, table_id1) && Enum.member?(:ets.all, table_id2)
  true

  """
  def parse(path, _options \\ []) do
    Loader.load(path)
  end

  @doc """

  ## Example

  iex> {:ok, {_pid, table_id}} = Exoffice.Parser.Excel2003.parse_sheet("./test/test_data/test.xls", 1)
  iex> Enum.member?(:ets.all, table_id)
  true

  """
  def parse_sheet(path, index, _options \\ []) do
    Loader.load(path, index) |> List.first
  end

  @doc """

  ## Example

  iex> {:ok, {_pid, table_id}} = Exoffice.Parser.Excel2003.parse_sheet("./test/test_data/test.xls", 1)
  iex> Exoffice.Parser.Excel2003.count_rows(table_id)
  10

  """
  def count_rows(pid) do
    Xlsxir.get_multi_info(pid, :rows)
  end

  @doc """

  ## Example

  iex> {:ok, {_pid, table_id}} = Exoffice.Parser.Excel2003.parse_sheet("./test/test_data/test.xls", 1)
  iex> Exoffice.Parser.Excel2003.get_rows(table_id) |> Enum.to_list
  [[23.0, 3.0, 12.0, 1.0, nil], [2.0, 12.0, 41.0, nil, nil],
  [nil, nil, nil, nil, nil], [nil, nil, nil, nil, nil],
  [nil, nil, nil, nil, nil], [nil, nil, nil, nil, nil],
  [nil, nil, nil, nil, nil], [nil, nil, nil, nil, nil],
  [nil, nil, nil, nil, nil], [nil, nil, nil, nil, nil]]

  """
  def get_rows(pid) do
    Stream.map(1..count_rows(pid), fn row ->
      Xlsxir.get_row(pid, row)
    end)
  end

  @doc """

  ## Example

  iex> [{:ok, {pid, table_id1}}, {:ok, {pid, table_id2}}] = Exoffice.Parser.Excel2003.parse("./test/test_data/test.xls")
  iex> Enum.member?(:ets.all, table_id1) && Enum.member?(:ets.all, table_id2)
  true

  iex> [{:ok, {pid, table_id1}}, {:ok, {pid, table_id2}}] = Exoffice.Parser.Excel2003.parse("./test/test_data/test.xls")
  iex> Exoffice.Parser.Excel2003.close(pid)
  iex> Enum.member?(:ets.all, table_id1) || Enum.member?(:ets.all, table_id2)
  false

  """
  def close(pid) do
    Xlsxir.close(pid)
  end

end
