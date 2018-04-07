defmodule Exoffice.Parser do
  @moduledoc """
  Module which provides common interface for
  parsing office files.
  """

  @doc """
  Function to parse file by path.
  May return tuple with result or list of tuples if excel file has multiple
  worksheets
  """
  @callback parse(String.t(), Keyword) :: [{:ok, pid} | {:error, String.t()}]

  @doc """
  Function to parse file by path and given sheet index
  Returns result with pid or error with reason
  """
  @callback parse_sheet(String.t(), Integer, Keyword) :: {:ok, pid} | {:error, String.t()}

  @doc """
  Returns a list of supported extensions by parser
  """
  @callback extensions() :: [String.t()]

  @doc """
  Count rows by pid
  """
  @callback count_rows(pid) :: integer

  @doc """
  Returns a stream with rows
  """
  @callback get_rows(pid) :: Stream

  @doc """
  Close opened pid (file or worksheet)
  """
  @callback close(pid) :: :ok
end
