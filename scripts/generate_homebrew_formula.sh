#!/usr/bin/env bash
set -euo pipefail

if [[ $# -ne 7 ]]; then
  echo "usage: $0 <owner> <repo> <version> <macos_arm_sha256> <macos_x86_sha256> <linux_x86_sha256> <output_path>" >&2
  exit 1
fi

owner="$1"
repo="$2"
version="$3"
macos_arm_sha="$4"
macos_x86_sha="$5"
linux_x86_sha="$6"
output_path="$7"

cat >"$output_path" <<EOF
class Docwarp < Formula
  desc "Bidirectional Markdown <-> DOCX converter"
  homepage "https://github.com/${owner}/${repo}"
  version "${version}"

  on_macos do
    if Hardware::CPU.arm?
      url "https://github.com/${owner}/${repo}/releases/download/v#{version}/docwarp-macos-aarch64"
      sha256 "${macos_arm_sha}"
    else
      url "https://github.com/${owner}/${repo}/releases/download/v#{version}/docwarp-macos-x86_64"
      sha256 "${macos_x86_sha}"
    end
  end

  on_linux do
    url "https://github.com/${owner}/${repo}/releases/download/v#{version}/docwarp-linux-x86_64"
    sha256 "${linux_x86_sha}"
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
EOF

echo "wrote formula to ${output_path}"
