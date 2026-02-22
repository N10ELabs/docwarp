class Docwarp < Formula
  desc "Bidirectional Markdown <-> DOCX converter"
  homepage "https://github.com/N10ELabs/docwarp"
  license "Apache-2.0"
  version "0.1.0"
  # NOTE: Release workflow generates a checksummed formula per tag and
  # publishes it as release asset `docwarp.rb`.
  # This copy remains a template until a tagged release is cut.

  on_macos do
    if Hardware::CPU.arm?
      url "https://github.com/N10ELabs/docwarp/releases/download/v#{version}/docwarp-macos-aarch64"
      sha256 "8ac800a23b035d1313ceec3c92837d7c170f75b44a30c98a40d9d93d4eaf426d"
    else
      url "https://github.com/N10ELabs/docwarp/releases/download/v#{version}/docwarp-macos-x86_64"
      sha256 "661d05c708f70bb57530d5f38cdbd4aa2774a5cba38ee1aac2bf890d25169795"
    end
  end

  on_linux do
    url "https://github.com/N10ELabs/docwarp/releases/download/v#{version}/docwarp-linux-x86_64"
    sha256 "87cac0e49ca080b4f588c84c380de1b02d3594bb29e0adbdf71afae7cfc845ab"
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
