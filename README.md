# Axon Outlook add-in

A self-contained Microsoft Outlook add-in that adds **Move** and **Download** buttons to the ribbon
(and the right-click menu). When you Move an email it suggests the best subfolder(s), and offers to
create a new one if nothing fits.

Suggestions come from an **OpenAI-compatible chat API** that **you point it at** — so the same
add-in works for:
- **Fully on-site / private** — a local model server on your network (Ollama, vLLM, LM Studio,
  LocalAI, …). Email content never leaves your network, no cloud, no per-user key.
- **Cloud** — OpenAI's API.

If the API is unreachable, Move/Download still work; only the AI ranking is skipped. This add-in is
independent of the Axon desktop app (the floating dot).

---

## How it decides where to send the request

Each PC reads `%APPDATA%\AxonOutlook\config.json`:

```json
{ "api_base": "http://YOUR-SERVER:11434/v1", "api_key": "", "model": "qwen2.5:3b" }
```

- **`api_base`** — any OpenAI-compatible endpoint. Examples:
  - Ollama: `http://host:11434/v1`
  - vLLM: `http://host:8000/v1`
  - LM Studio: `http://host:1234/v1`
  - OpenAI: `https://api.openai.com/v1`
- **`api_key`** — blank for most local servers; your key for OpenAI.
- **`model`** — e.g. `qwen2.5:3b` (recommended local), `llama3.2:3b`, or `gpt-4o-mini` for OpenAI.

The installer writes this for you (it asks for the three values). See `config.example.json`.

> If there's no `config.json` and the add-in is sitting next to an Axon desktop app, it falls back
> to that app's baked-in OpenAI key — that's the bundled-with-the-dot case.

## 1. Set up a local model server (recommended, by IT)

On a machine reachable on your LAN, run any OpenAI-compatible server. Easiest is
[Ollama](https://ollama.com):

```bash
ollama pull qwen2.5:3b            # small, fast, strong multilingual — great for folder filing
# expose on the network: set OLLAMA_HOST=0.0.0.0, then run `ollama serve`
```

Its OpenAI-compatible endpoint is `http://THAT-MACHINE:11434/v1`. Open TCP 11434 to the office.
Test from a client: `curl http://THAT-MACHINE:11434/v1/models`.

## 2. Install (end users)

Run **`AxonOutlook-Setup.exe`**, enter the API base URL + model when asked, then restart Outlook.
The **Move** and **Download** buttons appear automatically. No admin rights needed.

## 3. Build & test (developers)

```powershell
powershell -ExecutionPolicy Bypass -File build.ps1       # compile AxonAddin.dll (closes Outlook)
powershell -ExecutionPolicy Bypass -File register.ps1    # register for the current user
# create %APPDATA%\AxonOutlook\config.json (see config.example.json), then start Outlook
```

`unregister.ps1` removes it. Build the installer by compiling `installer\AxonOutlook.iss` with
[Inno Setup](https://jrsoftware.org/isdl.php) (`ISCC.exe`).

## Layout

```
src/AxonAddin.cs        the add-in (ribbon, dialogs, OpenAI-compatible suggestion call)
icons/                  ribbon icons
build.ps1               compile the DLL
register.ps1 / unregister.ps1   per-user COM registration (dev)
installer/AxonOutlook.iss       Inno Setup installer (registers + writes config)
config.example.json     client config template
```
