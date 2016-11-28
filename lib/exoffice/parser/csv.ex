defmodule Exoffice.Parser.CSV do

  @behaviour Exoffice.Parser

  def extensions, do: [".csv"]

  @doc """

  ## Example
  Parse file `test.csv` in `./test/test_data`:

  iex> [{:ok, pid}] = Exoffice.parse("./test/test_data/test.csv")
  iex> is_pid(pid)
  true

  iex> [{:ok, pid}] = Exoffice.parse("./test/test_data/test.csv")
  iex> rows = Exoffice.count_rows(pid)
  5

  """
  def parse(path, options \\ []) do
    stream = File.stream!(path) |> CSV.decode(options)
    case Agent.start_link(fn -> stream end, name: String.to_atom(path)) do
        {:ok, pid} -> [{:ok, pid}]

        # Reoped previously opened file
        {:error, {:already_started, pid}} ->
          close(pid)
          parse(path, options)

        {:error, reason} -> [{:error, reason}]
    end
  end

  def parse_sheet(path, _, options \\ []) do
    parse(path, options) |> List.first
  end

  def count_rows(pid) do
    Agent.get(pid, &Enum.count/1)
  end

  def get_rows(pid) do
    Agent.get(pid, &(&1))
  end

  def close(pid) do
    if Process.alive?(pid), do: Agent.stop(pid), else: :ok
  end

end
