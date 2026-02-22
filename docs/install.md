# Installation

This page covers installation from published release artifacts.

## Binary Download

Pick the artifact for your platform from the GitHub release:

- macOS Apple Silicon: `docwarp-macos-aarch64`
- macOS Intel: `docwarp-macos-x86_64`
- Linux x86_64: `docwarp-linux-x86_64`
- Windows x86_64: `docwarp-windows-x86_64.exe`

You can verify downloaded files with `checksums.txt` from the same release.

### macOS / Linux

```bash
# Example: Linux x86_64
curl -fL -o docwarp \
  https://github.com/N10ELabs/docwarp/releases/download/v0.1.0/docwarp-linux-x86_64

chmod +x ./docwarp
./docwarp --help
```

### Windows (PowerShell)

```powershell
Invoke-WebRequest `
  -Uri "https://github.com/N10ELabs/docwarp/releases/download/v0.1.0/docwarp-windows-x86_64.exe" `
  -OutFile ".\docwarp.exe"

.\docwarp.exe --help
```

## Homebrew Tap

`docwarp` publishes a Homebrew formula as release asset `docwarp.rb` with release-specific checksums.

Install via tap:

```bash
brew tap N10ELabs/docwarp https://github.com/N10ELabs/docwarp
brew install docwarp
```

If you need to install from a specific release formula file:

```bash
curl -fL -o docwarp.rb \
  https://github.com/N10ELabs/docwarp/releases/download/v0.1.0/docwarp.rb

brew install --formula ./docwarp.rb
```
