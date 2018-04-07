defmodule Exoffice.Parser.Excel2003.Index do
  @moduledoc """
  An Agent process named `Index` which holds state of an index. Provides functions to create the process, increment the index by 1, retrieve the current index
  and ultimately kill the process.
  """

  @doc """
  Initiates a new `Index` Agent process with a value of `0`.
  """
  def new do
    Agent.start_link(fn -> 0 end, name: __MODULE__)
  end

  @doc """
  Increments active `Index` Agent process by `1`.
  """
  def inc do
    Agent.update(__MODULE__, &(&1 + 1))
  end

  @doc """
  Returns current value of `Index` Agent process
  """
  def get do
    Agent.get(__MODULE__, & &1)
  end

  @doc """
  Deletes `Index` Agent process
  """
  def del do
    Agent.stop(__MODULE__)
  end
end
