class Instruct < Formula
  desc "Bidirectional Markdown <-> DOCX converter"
  homepage "https://github.com/OWNER/instruct"
  version "0.1.0"

  on_macos do
    if Hardware::CPU.arm?
      url "https://github.com/OWNER/instruct/releases/download/v#{version}/instruct-macos-aarch64"
      sha256 "REPLACE_WITH_SHA256"
    else
      url "https://github.com/OWNER/instruct/releases/download/v#{version}/instruct-macos-x86_64"
      sha256 "REPLACE_WITH_SHA256"
    end
  end

  on_linux do
    url "https://github.com/OWNER/instruct/releases/download/v#{version}/instruct-linux-x86_64"
    sha256 "REPLACE_WITH_SHA256"
  end

  def install
    bin.install Dir["*"] => "instruct"
  end

  test do
    assert_match "Convert documentation", shell_output("#{bin}/instruct --help")
  end
end
