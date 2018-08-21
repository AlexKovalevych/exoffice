defmodule ExofficeTest do
  use ExUnit.Case

  describe "parse options" do
    test "use CSV options" do
      [{:ok, pid, parser}] = Exoffice.parse("./test/test_data/test_semicolon.csv", parser_options: [separator: ?;])

      expected = [
        ok: ["2", "23", "23", "2", "asg", "2", "sadg"],
        ok: ["sd", "123", "2", "3", "12", "", "23"],
        ok: ["g", "", "", "1", "", "1", ""],
        ok: ["2016-01-01", "", "", "", "3", "", ""],
        ok: ["", "", "", "", "", "", ""],
        ok: ["", "", "", "", "", "", ""],
        ok: ["", "", "", "", "", "", ""],
        ok: ["", "", "", "", "", "", ""],
        ok: ["", "", "", "", "", "", ""],
        ok: ["", "", "", "", "", "", ""],
        ok: ["", "", "", "", "", "", ""],
        ok: ["", "", "", "", "", "", ""],
        ok: ["", "", "", "", "", "", ""],
        ok: ["", "", "", "", "", "", ""],
        ok: ["", "", "", "", "", "", ""],
        ok: ["", "", "", "", "", "", ""],
        ok: ["", "", "", "", "", "", ""],
        ok: ["", "", "", "", "", "", ""],
        ok: ["", "", "", "", "", "", ""],
        ok: ["", "", "", "", "", "", ""],
        ok: ["", "", "", "", "", "", ""],
        ok: ["", "", "", "", "", "", ""]
      ]

      assert is_pid(pid)
      assert Exoffice.get_rows(pid, parser) |> Enum.to_list() == expected
    end
  end
end
