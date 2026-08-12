# Clippy Possession

This project keeps the original Office XP Assistant as Clippy's interface and
loads a thin native Word add-in to intercept its built-in question box. The
modern intelligence and tools will remain on the macOS host.

## Repository layout

```text
src/
  addin/          Word COM add-in and COM export definition
  tools/          Native XP diagnostics and Microsoft Agent probe
server/            TypeScript HTTP bridge running on the macOS host
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

## Run the OpenAI-backed host server

The host requires Node.js 22 or newer and an OpenAI API key. Keep the key in
the host environment rather than in the repository. The model defaults to
`gpt-5.6-luna`; set `CLIPPY_OPENAI_MODEL` to override it.

```sh
cd server
npm install
npm test
npm run build
export OPENAI_API_KEY="your_api_key_here"
npm start
```

The server listens on `0.0.0.0:3210` and accepts `POST /message` with a JSON
body such as `{"text":"hello"}`. It sends each question to the OpenAI Responses
API as an independent turn and returns a compact Markdown-lite reply in the
same `{"text":"..."}` response shape. The XP shim safely maps emphasis, inline
code, links, headings, and one trailing list to the native formatting supported
by Office Assistant balloons. The XP shim reaches the server through QEMU/UTM's
host gateway at `10.0.2.2:3210`.

See the [OpenAI API quickstart](https://developers.openai.com/api/docs/quickstart)
for API-key setup.
