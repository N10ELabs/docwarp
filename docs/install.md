# Installation

This page covers installation from published release artifacts.

## Binary Download

Pick the artifact for your platform from the GitHub release:

- macOS Apple Silicon: `instruct-macos-aarch64`
- macOS Intel: `instruct-macos-x86_64`
- Linux x86_64: `instruct-linux-x86_64`
- Windows x86_64: `instruct-windows-x86_64.exe`

You can verify downloaded files with `checksums.txt` from the same release.

### macOS / Linux

```bash
# Example: Linux x86_64
curl -fL -o instruct \
  https://github.com/N10ELabs/instruct/releases/download/v0.1.0/instruct-linux-x86_64

chmod +x ./instruct
./instruct --help
```

### Windows (PowerShell)

```powershell
Invoke-WebRequest `
  -Uri "https://github.com/N10ELabs/instruct/releases/download/v0.1.0/instruct-windows-x86_64.exe" `
  -OutFile ".\instruct.exe"

.\instruct.exe --help
```

## Homebrew Tap

`instruct` publishes a Homebrew formula as release asset `instruct.rb` with release-specific checksums.

Install via tap:

```bash
brew tap N10ELabs/instruct https://github.com/N10ELabs/instruct
brew install instruct
```

If you need to install from a specific release formula file:

```bash
curl -fL -o instruct.rb \
  https://github.com/N10ELabs/instruct/releases/download/v0.1.0/instruct.rb

brew install --formula ./instruct.rb
```
