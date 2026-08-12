# Clippy Possession

This project keeps the original Office XP Assistant as Clippy's interface and
loads a thin native Word add-in to intercept its built-in question box. The
modern intelligence and tools will remain on the macOS host.

## Repository layout

```text
src/
  addin/          Word COM add-in and COM export definition
  tools/          Native XP diagnostics and Microsoft Agent probe
scripts/
  demos/          Visible Clippy demonstrations
  diagnostics/    Microsoft Agent validation and discovery scripts
docs/
  captures/       Curated WindowWatch reverse-engineering evidence
  screenshots/    Visible acceptance-test evidence
  ARCHITECTURE.md
  DEVLOG.md
build/            Generated binaries; ignored by Git
build.bat         One-command Visual C++ 2010 x86 build
```

## Build on Windows XP

Copy the repository tree to `C:\clippy`, then run:

```bat
cd /d C:\clippy
build.bat
```

The build produces:

```text
build\ClippyShim.dll
build\ClippyProbe.exe
build\WindowWatch.exe
```

Register the Word add-in for the current XP user with:

```bat
regsvr32 /s C:\clippy\build\ClippyShim.dll
```

See [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) for the system design and
[docs/DEVLOG.md](docs/DEVLOG.md) for verified implementation details.
