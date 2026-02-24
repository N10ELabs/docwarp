class Docwarp < Formula
  desc "Bidirectional Markdown <-> DOCX converter"
  homepage "https://github.com/N10ELabs/docwarp"
  license "Apache-2.0"
  version "0.1.1"

  on_macos do
    if Hardware::CPU.arm?
      url "https://github.com/N10ELabs/docwarp/releases/download/v#{version}/docwarp-macos-aarch64"
      sha256 "8c3fdebc94a4733ccbbdf0488d8822432bbb942c0c6fec69010a7eccbfd1dba5"
    else
      url "https://github.com/N10ELabs/docwarp/releases/download/v#{version}/docwarp-macos-x86_64"
      sha256 "bbf6973dcf0d3574eab2859e4dc6ccd9b71e5256641e704d50b29152d5b8e229"
    end
  end

  on_linux do
    url "https://github.com/N10ELabs/docwarp/releases/download/v#{version}/docwarp-linux-x86_64"
    sha256 "d2fc7908f1a7de3a4594f2328c354f9d3e3cf8e13b70691348dc3e6488eade8c"
  end

  def install
    artifact = Dir["*"].find { |f| File.file?(f) }
    raise "expected a single release artifact" if artifact.nil?

    bin.install artifact => "docwarp"
  end

  test do
    assert_match "Convert documentation", shell_output("#{bin}/docwarp --help")
  end
end
