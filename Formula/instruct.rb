class Instruct < Formula
  desc "Bidirectional Markdown <-> DOCX converter"
  homepage "https://github.com/N10ELabs/instruct"
  version "0.1.0"
  # NOTE: Release workflow generates a checksummed formula per tag and
  # publishes it as release asset `instruct.rb`.
  # This copy remains a template until a tagged release is cut.

  on_macos do
    if Hardware::CPU.arm?
      url "https://github.com/N10ELabs/instruct/releases/download/v#{version}/instruct-macos-aarch64"
      sha256 "REPLACE_WITH_SHA256"
    else
      url "https://github.com/N10ELabs/instruct/releases/download/v#{version}/instruct-macos-x86_64"
      sha256 "REPLACE_WITH_SHA256"
    end
  end

  on_linux do
    url "https://github.com/N10ELabs/instruct/releases/download/v#{version}/instruct-linux-x86_64"
    sha256 "REPLACE_WITH_SHA256"
  end

  def install
    artifact = Dir["*"].find { |f| File.file?(f) }
    raise "expected a single release artifact" if artifact.nil?

    bin.install artifact => "instruct"
  end

  test do
    assert_match "Convert documentation", shell_output("#{bin}/instruct --help")
  end
end
