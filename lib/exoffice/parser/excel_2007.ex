defmodule Exoffice.Parser.Excel2007 do

  @behaviour Exoffice.Parser

  def extensions, do: [".xlsx"]

  @doc """

  ## Example

  iex> [{:ok, {pid, table_id1}}, {:ok, {pid, table_id2}}] = Exoffice.Parser.Excel2007.parse("./test/test_data/test.xlsx")
  iex> Enum.member?(:ets.all, table_id1) && Enum.member?(:ets.all, table_id2)
  true

  """
  def parse(path, _options \\ []) do
    Xlsxir.multi_extract(path)
  end

  @doc """

  ## Example

  iex> {:ok, {_pid, table_id}} = Exoffice.Parser.Excel2007.parse_sheet("./test/test_data/test.xlsx", 1)
  iex> Enum.member?(:ets.all, table_id)
  true

  """
  def parse_sheet(path, index, _options \\ []) do
    Xlsxir.multi_extract(path, index, false)
  end

  @doc """

  ## Example

  iex> {:ok, {_pid, table_id}} = Exoffice.Parser.Excel2007.parse_sheet("./test/test_data/test.xlsx", 1)
  iex> Exoffice.Parser.Excel2007.count_rows(table_id)
  10

  """
  def count_rows(pid) do
    Xlsxir.get_multi_info(pid, :rows)
  end

  @doc """

  ## Example

  iex> {:ok, {_pid, table_id}} = Exoffice.Parser.Excel2007.parse_sheet("./test/test_data/test.xlsx", 1)
  iex> Exoffice.Parser.Excel2007.get_rows(table_id) |> Enum.to_list
  [[23, 3, 12, 1, nil], [2, 12, 41, nil, nil],
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

  iex> [{:ok, {pid, table_id1}}, {:ok, {pid, table_id2}}] = Exoffice.Parser.Excel2007.parse("./test/test_data/test.xlsx")
  iex> Enum.member?(:ets.all, table_id1) && Enum.member?(:ets.all, table_id2)
  true

  iex> [{:ok, {pid, table_id1}}, {:ok, {pid, table_id2}}] = Exoffice.Parser.Excel2007.parse("./test/test_data/test.xlsx")
  iex> Exoffice.Parser.Excel2007.close(pid)
  iex> Enum.member?(:ets.all, table_id1) || Enum.member?(:ets.all, table_id2)
  false

  """
  def close(pid) do
    Xlsxir.close(pid)
  end

end
