defmodule DocTest do
  use ExUnit.Case
  doctest Exoffice
  doctest Exoffice.Parser.CSV
  doctest Exoffice.Parser.Excel2007
  doctest Exoffice.Parser.Excel2003
  doctest Exoffice.Parser.Excel2003.Date
end
