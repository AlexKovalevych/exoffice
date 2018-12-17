defmodule ExofficeTest do
  use ExUnit.Case

  describe "parse options" do
    test "use CSV options" do
      [{:ok, pid, parser}] = Exoffice.parse("./test/test_data/test_semicolon.csv", parser_options: [separator: ?;])

      expected = [
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

      assert is_pid(pid)
      assert Exoffice.get_rows(pid, parser) |> Enum.to_list() == expected
    end

    test "decode CSV in a safe way" do
      [{:ok, pid, parser}] =
        Exoffice.parse("./test/test_data/test_invalid.csv", parser_options: [safe: true, separator: ?;])

      expected = [
        {:ok, ["2", "23", "23", "2", "asg", "2", "sadg"]},
        {:ok, ["sd", "123", "2", "3", "12", "", "23"]},
        {:ok, ["g", "", "", "1", "", "1", ""]},
        {:ok, ["2016-01-01", "", "", "", "3", "", ""]},
        {:ok, ["", "", "", "", "", "", ""]},
        {:ok, ["", "", "", "", "", "", ""]},
        {:ok, ["", "", "", "", "", "", ""]},
        {:ok, ["", "", "", "", "", "", ""]},
        {:ok, ["", "", "", "", "", "", ""]},
        {:ok, ["", "", "", "", "", "", ""]},
        {:ok, ["", "", "", "", "", "", ""]},
        {:ok, ["", "", "", "", "", "", ""]},
        {:ok, ["", "", "", "", "", "", ""]},
        {:ok, ["", "", "", "", "", "", ""]},
        {:ok, ["", "", "", "", "", "", ""]},
        {:ok, ["", "", "", "", "", "", ""]},
        {:ok, ["", "", "", "", "", "", ""]},
        {:error, "Row has length 2 - expected length 7 on line 18"},
        {:error, "Row has length 5 - expected length 7 on line 19"},
        {:error, "Row has length 3 - expected length 7 on line 20"},
        {:error, "Row has length 6 - expected length 7 on line 21"},
        {:error, "Row has length 4 - expected length 7 on line 22"}
      ]

      assert is_pid(pid)
      assert Exoffice.get_rows(pid, parser) |> Enum.to_list() == expected
    end
  end

  describe "handle rich text" do
    test "parse .xls with rich text" do
      [{:ok, pid, parser}] = Exoffice.parse("./test/test_data/test_rich_text.xls")

      expected = [
        ["X條"]
      ]

      assert Exoffice.get_rows(pid, parser) |> Enum.to_list() == expected
    end
  end
end
