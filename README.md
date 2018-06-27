# Exoffice

[![Build Status](https://travis-ci.org/AlexKovalevych/exoffice.svg?branch=master)](https://travis-ci.org/AlexKovalevych/exoffice)

Exoffice is an Elixir library that parses common Excel formats: .xls (Excel 2003), .xlsx, (Excel 2007), .csv and saves parsed data into `ets` (excep .csv, which uses stream)

## Installation

If [available in Hex](https://hex.pm/docs/publish), the package can be installed as:

  1. Add `exoffice` to your list of dependencies in `mix.exs`:

    ```elixir
    def deps do
      [{:exoffice, "~> 0.2.0"}]
    end
    ```

## Usage

Each parser has functions to parse, count rows, get rows and close file.

By default all sheets will be parsed:

```elixir
  [{:ok, pid1, parser1}, {:ok, pid2, parser1}] = Exoffice.parse("./test/test_data/test.xls")
  [{:ok, pid3, parser2}, {:ok, pid4, parser2}] = Exoffice.parse("./test/test_data/test.xlsx")
  [{:ok, pid5, parser3}] = Exoffice.parse("./test/test_data/test.csv")
```

To parse a single sheet:

```elixir
  {:ok, pid1, parser1} = Exoffice.parse("./test/test_data/test.xls", 1)
  {:ok, pid2, parser2} = Exoffice.parse("./test/test_data/test.xlsx", 1)
  {:ok, pid3, parser3} = Exoffice.parse("./test/test_data/test.csv", 1)
```

To count rows:

```
  Exoffice.count_rows(pid, parser)
```

To get rows:

```
  stream = Exoffice.get_rows(pid, parser)
```

Don't forget to close pid when you don't need data anymore:

```
  Exoffice.close(pid, parser)
```
