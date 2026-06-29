# Axon Outlook add-in

A self-contained Microsoft Outlook add-in that adds **Move** and **Download** buttons to the ribbon
(and the right-click menu). When you Move an email, it suggests the best subfolder(s) — and offers to
create a new one if nothing fits.

**Privacy:** folder suggestions are produced by a **local LLM on your own server (via [Ollama](https://ollama.com))** —
email content never leaves your network. There is **no cloud API and no API key**. If the server is
unreachable, Move/Download still work; only the AI suggestions are skipped.

This add-in is independent of the Axon desktop app (the floating dot).

---

## 1. Set up the model server (once, by IT)

On a Windows/Linux/Mac server reachable on your LAN:

```bash
# install Ollama from https://ollama.com, then:
ollama pull llama3.2:3b          # ~2 GB; a small, fast model that's plenty for folder filing
# expose it on the network (default port 11434):
#   Windows: set the OLLAMA_HOST environment variable to 0.0.0.0 and restart Ollama
#   Linux/Mac: OLLAMA_HOST=0.0.0.0 ollama serve
```

Make sure TCP **11434** is open to the office network. Test from a client PC:

```
curl http://YOUR-SERVER:11434/api/tags
```

Other small models work too (e.g. `qwen2.5:3b`, `phi3:mini`) — set the name in the client config.

## 2. Client config

Each PC reads `%APPDATA%\AxonOutlook\config.json`:

```json
{ "ollama_url": "http://YOUR-SERVER:11434", "model": "llama3.2:3b" }
```

The installer writes this for you (it asks for the server URL). See `config.example.json`.

## 3. Install (end users)

Run **`AxonOutlook-Setup.exe`**, enter your Ollama server URL when asked, then restart Outlook.
The **Move** and **Download** buttons appear automatically. No admin rights needed.

## 4. Build & test (developers)

```powershell
powershell -ExecutionPolicy Bypass -File build.ps1       # compile AxonAddin.dll (closes Outlook)
powershell -ExecutionPolicy Bypass -File register.ps1    # register for the current user
# create %APPDATA%\AxonOutlook\config.json (copy config.example.json), then start Outlook
```

`unregister.ps1` removes it. To build the installer, compile `installer\AxonOutlook.iss` with
[Inno Setup](https://jrsoftware.org/isdl.php) (`ISCC.exe`).

## Layout

```
src/AxonAddin.cs        the add-in (ribbon, dialogs, Ollama suggestion call)
icons/                  ribbon icons
build.ps1               compile the DLL
register.ps1 / unregister.ps1   per-user COM registration (dev)
installer/AxonOutlook.iss       Inno Setup installer (registers + writes config)
config.example.json     client config template
```
