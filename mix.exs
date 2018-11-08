defmodule Exoffice.Mixfile do
  use Mix.Project

  def project do
    [
      app: :exoffice,
      version: "0.3.0",
      name: "Exoffice",
      elixir: "~> 1.6",
      build_embedded: Mix.env() == :prod,
      start_permanent: Mix.env() == :prod,
      description: description(),
      package: package(),
      deps: deps()
    ]
  end

  # Configuration for the OTP application
  #
  # Type "mix help compile.app" for more information
  def application do
    [applications: [:logger, :iconv, :xlsxir]]
  end

  defp description do
    """
      File parser for popular excel formats: xls (Excel 2003), csv, xlsx (Excel 2007).
      Stores data in ets (except for csv, which uses stream).
    """
  end

  # Dependencies can be Hex packages:
  #
  #   {:mydep, "~> 0.3.0"}
  #
  # Or git/path repositories:
  #
  #   {:mydep, git: "https://github.com/elixir-lang/mydep.git", tag: "0.1.0"}
  #
  # Type "mix help deps" for more examples and options
  defp deps do
    [
      {:xlsxir, "~> 1.6"},
      {:csv, "~> 2.1"},
      {:ex_doc, "~> 0.19.1"},
      {:earmark, "~> 1.0"},
      {:iconv, "~> 1.0"}
    ]
  end

  defp package do
    [
      maintainers: ["Alex Kovalevych", "Rock Neurotiko (Miguel G)"],
      licenses: ["MIT License"],
      links: %{
        "Github" => "https://github.com/alexkovalevych/exoffice",
        "Change Log" => "https://hexdocs.pm/exoffice/changelog.html"
      }
    ]
  end
end
