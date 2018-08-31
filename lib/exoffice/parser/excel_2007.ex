defmodule Exoffice.Parser.Excel2007 do
  @behaviour Exoffice.Parser

  def extensions, do: [".xlsx"]

  @doc """

  ## Example

  iex> [{:ok, pid1}, {:ok, pid2}] = Exoffice.Parser.Excel2007.parse("./test/test_data/test.xlsx")
  iex> Enum.member?(:ets.all, pid1) && Enum.member?(:ets.all, pid2)
  true

  """
  def parse(path, options \\ []) do
    multi_extract(path, nil, options)
  end

  @doc """

  ## Example

  iex> {:ok, pid} = Exoffice.Parser.Excel2007.parse_sheet("./test/test_data/test.xlsx", 1)
  iex> Enum.member?(:ets.all, pid)
  true

  """
  def parse_sheet(path, index, options \\ []) do
    multi_extract(path, index, options)
  end

  @doc """

  ## Example

  iex> {:ok, pid} = Exoffice.Parser.Excel2007.parse_sheet("./test/test_data/test.xlsx", 1)
  iex> Exoffice.Parser.Excel2007.count_rows(pid)
  10

  """
  def count_rows(pid) do
    Xlsxir.get_multi_info(pid, :rows)
  end

  @doc """

  ## Example

  iex> {:ok, pid} = Exoffice.Parser.Excel2007.parse_sheet("./test/test_data/test.xlsx", 1)
  iex> Exoffice.Parser.Excel2007.get_rows(pid) |> Enum.to_list
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

  iex> [{:ok, pid1}, {:ok, pid2}] = Exoffice.Parser.Excel2007.parse("./test/test_data/test.xlsx")
  iex> Enum.member?(:ets.all, pid1) && Enum.member?(:ets.all, pid2)
  true

  iex> [{:ok, pid1}, {:ok, pid2}] = Exoffice.Parser.Excel2007.parse("./test/test_data/test.xlsx")
  iex> Exoffice.Parser.Excel2007.close(pid1)
  iex> Exoffice.Parser.Excel2007.close(pid2)
  iex> Enum.member?(:ets.all, pid1) || Enum.member?(:ets.all, pid2)
  false

  """
  def close(pid) do
    Xlsxir.close(pid)
  end

  defp multi_extract(path, index, options) do
    # Use default values of Xlsxir for timer and excel params
    Xlsxir.multi_extract(path, index, false, nil, options)
  end
end
