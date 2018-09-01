defmodule Exoffice.Parser.CSV do
  @behaviour Exoffice.Parser

  def extensions, do: [".csv"]

  @doc """

  iex> [{:ok, pid}] = Exoffice.Parser.CSV.parse("./test/test_data/test.csv")
  iex> is_pid(pid)
  true

  """
  def parse(path, options \\ []) do
    stream = File.stream!(path) |> CSV.decode!(options)

    case Agent.start_link(fn -> stream end, name: String.to_atom(path)) do
      {:ok, pid} ->
        [{:ok, pid}]

      # Reoped previously opened file
      {:error, {:already_started, pid}} ->
        close(pid)
        parse(path, options)

      {:error, reason} ->
        [{:error, reason}]
    end
  end

  @doc """

  ## Example

  iex> {:ok, pid} = Exoffice.Parser.CSV.parse_sheet("./test/test_data/test.csv", 1)
  iex> is_pid(pid)
  true

  """
  def parse_sheet(path, _, options \\ []) do
    parse(path, options) |> List.first()
  end

  @doc """

  ## Example

  iex> {:ok, pid} = Exoffice.Parser.CSV.parse_sheet("./test/test_data/test.csv", 1)
  iex> Exoffice.Parser.CSV.count_rows(pid)
  22

  """
  def count_rows(pid) do
    Agent.get(pid, &Enum.count/1)
  end

  @doc """

  ## Example

  iex> {:ok, pid} = Exoffice.Parser.CSV.parse_sheet("./test/test_data/test.csv", 1)
  iex> Exoffice.Parser.CSV.get_rows(pid) |> Enum.to_list
  [
    ["2", "23", "23", "2", "asg", "2", "sadg"],
    ["sd", "123", "2", "3", "12", "", "23"],
    ["g", "", "", "1", "", "1", ""],
    ["2016-01-01", "", "", "", "3", "", ""],
    ["", "", "", "", "", "", ""],
    ["", "", "", "", "", "", ""],
    ["", "", "", "", "", "", ""],
    ["", "", "", "", "", "", ""],
    ["", "", "", "", "", "", ""],
    ["", "", "", "", "", "", ""],
    ["", "", "", "", "", "", ""],
    ["", "", "", "", "", "", ""],
    ["", "", "", "", "", "", ""],
    ["", "", "", "", "", "", ""],
    ["", "", "", "", "", "", ""],
    ["", "", "", "", "", "", ""],
    ["", "", "", "", "", "", ""],
    ["", "", "", "", "", "", ""],
    ["", "", "", "", "", "", ""],
    ["", "", "", "", "", "", ""],
    ["", "", "", "", "", "", ""],
    ["", "", "", "", "", "", ""]
  ]
  """
  def get_rows(pid) do
    Agent.get(pid, & &1)
  end

  @doc """

  ## Example

  iex> [{:ok, pid}] = Exoffice.Parser.CSV.parse("./test/test_data/test.csv")
  iex> Process.alive? pid
  true

  iex> [{:ok, pid}] = Exoffice.Parser.CSV.parse("./test/test_data/test.csv")
  iex> Exoffice.Parser.CSV.close(pid)
  iex> Process.alive? pid
  false

  """
  def close(pid) do
    if Process.alive?(pid), do: Agent.stop(pid), else: :ok
  end
end
