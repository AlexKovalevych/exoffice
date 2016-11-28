defmodule Exoffice do
  alias Exoffice.Parser.{Excel2007, Excel2003, CSV}

  @default_parsers [Excel2007, Excel2003, CSV]

  def parse(path, sheet \\ nil) do
    config = Application.get_env(:exoffice, __MODULE__, [])
    parsers = case config[:parsers] do
      nil -> []
      parsers -> parsers
    end
    parsers = parsers ++ @default_parsers
    extension = Path.extname(path)
    parser = Enum.reduce_while(parsers, nil, fn parser, acc ->
      if Enum.member?(parser.extensions, extension), do: {:halt, parser}, else: {:cont, acc}
    end)

    if is_nil(parser) do
      {:error, "No parser for this file"}
    else
      case is_nil(sheet) do
        true ->
          pids = parser.parse(path)
          Enum.map(pids, fn
            {:ok, pid} -> {:ok, pid, parser}
            {:error, reason} -> {:error, reason}
          end)
        false ->
          result = parser.parse_sheet(path, sheet)
          case result do
            {:ok, pid} -> {:ok, pid, parser}
            _ -> result
          end
      end
    end
  end

  def get_rows(pid, parser) do
    parser.get_rows(pid)
  end

  def close(pid, parser) do
    parser.close(pid)
  end

  def count_rows(pid, parser) do
    parser.count_rows(pid)
  end

end
