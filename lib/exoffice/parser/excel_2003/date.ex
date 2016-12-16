defmodule Exoffice.Parser.Excel2003.Date do

  epoch = {{1970, 1, 1}, {0, 0, 0}}
  @epoch :calendar.datetime_to_gregorian_seconds(epoch)

  @doc """
  Converts date from excel to erlang datetime

  ## Parameters
  - `value` - date from the excel file (usually a number, e.g. 40908.0)
  - `base_date` - `base_date` field from the excel struct. Defines Windows or Mac date (1900 or 1904)

  ## Example
  iex> Exoffice.Parser.Excel2003.Date.to_date(40908.0, 1904)
  {{2016, 1, 1}, {0, 0, 0}}

  """
  def to_date(value, base_date) do
    excel_base_date = case base_date do
      # Windows
      1900 -> if value < 60, do: 25568, else: 25569
      # Mac
      _ -> 24107
    end

    (value - excel_base_date) * 86400 |> round |> from_timestamp
  end

  defp from_timestamp(timestamp) do
    timestamp
    |> (fn v -> v + @epoch end).()
    |> :calendar.gregorian_seconds_to_datetime
  end

end
