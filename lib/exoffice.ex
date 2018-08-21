defmodule Exoffice do
  alias Exoffice.Parser.{Excel2007, Excel2003, CSV}

  @moduledoc """
  Parse data from ".xls", ".xlsx", ".csv" files and save it to ETS. Provides common interface
  to count, get rows as a stream and close opened Agent process
  """

  @default_parsers [Excel2007, Excel2003, CSV]

  @doc """
  Parses and saves rows to ETS. Returns list of tuples, where each tuple is either
  {:ok, pid, parser} or {:error, reason}
  ## Parameters
  - `path` - file path of a file (".xls", ".xlsx" or ".csv" file) as a string
  - `options` - keyword list of options:
    - `sheet` - index of worksheet to parse (optional). By default all sheets will be parsed
    - `parsers` - extra parsers to use (optional). It will be added to the configuration parsers and the default parsers.
    - `parser_options` - options for the parser (optional)

  ## Example
  Parse all worksheets in different files:

      iex> [{:ok, pid1, _}, {:ok, pid2, _}] = Exoffice.parse("./test/test_data/test.xls")
      iex> Enum.member?(:ets.all, pid1) && Enum.member?(:ets.all, pid2)
      true

      iex> [{:ok, pid1, _}, {:ok, pid2, _}] = Exoffice.parse("./test/test_data/test.xlsx")
      iex> Enum.member?(:ets.all, pid1) && Enum.member?(:ets.all, pid2)
      true

      iex> [{:ok, pid, _}] = Exoffice.parse("./test/test_data/test.csv")
      iex> is_pid(pid)
      true
  """
  def parse(path, options \\ []) do
    sheet = options[:sheet]
    parser_options = options[:parser_options] || []

    config = Application.get_env(:exoffice, __MODULE__, [])

    options_parsers = options[:parsers] || []
    config_parsers = config[:parsers] || []

    parsers = options_parsers ++ config_parsers ++ @default_parsers
    extension = Path.extname(path)

    parser = find_parser(parsers, extension)

    if is_nil(parser) do
      {:error, "No parser for this file"}
    else
      case is_nil(sheet) do
        true ->
          pids = parser.parse(path, parser_options)

          Enum.map(pids, fn
            {:ok, pid} -> {:ok, pid, parser}
            {:error, reason} -> {:error, reason}
          end)

        false ->
          result = parser.parse_sheet(path, sheet, parser_options)

          case result do
            {:ok, pid} -> {:ok, pid, parser}
            _ -> result
          end
      end
    end
  end

  @doc """
  Returns stream of parsed rows

  ## Parameters
  - `pid` - is a pid, returned after parsing file
  - `parser` - is a module, used for parsing a file, returned with pid after parsing

  ## Example

      iex> [{:ok, pid1, parser1}, {:ok, _pid2, _parser2}] = Exoffice.parse("./test/test_data/test.xls")
      iex> Exoffice.get_rows(pid1, parser1) |> Enum.count
      23

      iex> [{:ok, pid1, parser1}, {:ok, _pid2, _parser2}] = Exoffice.parse("./test/test_data/test.xlsx")
      iex> Exoffice.get_rows(pid1, parser1) |> Enum.count
      23

      iex> [{:ok, pid, parser}] = Exoffice.parse("./test/test_data/test.csv")
      iex> Exoffice.get_rows(pid, parser) |> Enum.count
      22

  """
  def get_rows(pid, parser) do
    parser.get_rows(pid)
  end

  @doc """
  Closes Agent process of the parsed file. Should be run when you finished working with parsed data.

  ## Parameteres
  - `pid` - is a pid, returned after parsing file
  - `parser` - is a module, used for parsing a file, returned with pid after parsing

  ## Example
      iex> [{:ok, pid1, parser1}, {:ok, pid2, parser2}] = Exoffice.parse("./test/test_data/test.xls")
      iex> Exoffice.close(pid1, parser1)
      iex> Exoffice.close(pid2, parser2)
      iex> Enum.member?(:ets.all, pid1) || Enum.member?(:ets.all, pid2)
      false

      iex> [{:ok, pid1, parser1}, {:ok, pid2, parser2}] = Exoffice.parse("./test/test_data/test.xlsx")
      iex> Exoffice.close(pid1, parser1)
      iex> Exoffice.close(pid2, parser2)
      iex> Enum.member?(:ets.all, pid1) || Enum.member?(:ets.all, pid2)
      false

      iex> [{:ok, pid, parser}] = Exoffice.parse("./test/test_data/test.csv")
      iex> Exoffice.close(pid, parser)
      iex> Process.alive?(pid)
      false

  """
  def close(pid, parser) do
    parser.close(pid)
  end

  @doc """
  Count rows in parsed file.

  ## Parameteres
  - `pid` - is a pid, returned after parsing file
  - `parser` - is a module, used for parsing a file, returned with pid after parsing

  ## Example
      iex> [{:ok, pid1, parser1}, {:ok, pid2, parser2}] = Exoffice.parse("./test/test_data/test.xls")
      iex> [Exoffice.count_rows(pid1, parser1), Exoffice.count_rows(pid2, parser2)]
      [23,10]

      iex> [{:ok, pid1, parser1}, {:ok, pid2, parser2}] = Exoffice.parse("./test/test_data/test.xlsx")
      iex> [Exoffice.count_rows(pid1, parser1), Exoffice.count_rows(pid2, parser2)]
      [23,10]

      iex> [{:ok, pid, parser}] = Exoffice.parse("./test/test_data/test.csv")
      iex> Exoffice.count_rows(pid, parser)
      22

  """
  def count_rows(pid, parser) do
    parser.count_rows(pid)
  end

  defp find_parser(parsers, extension) do
    Enum.reduce_while(parsers, nil, fn parser, acc ->
      if Enum.member?(parser.extensions, extension), do: {:halt, parser}, else: {:cont, acc}
    end)
  end
end
