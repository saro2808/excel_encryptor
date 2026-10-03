# Excel Anonymiser & Restorer

A small local web app that replaces sensitive values in Excel files with
opaque tokens before you share the file with an AI chatbot, then restores the
original values — and all original formatting — in the result you get back.

---

## How it works

```
Original Excel  ──[Anonymise]──►  anonymised.xlsx  ──►  AI chatbot
                                  mapping_key.json
                                                         ▼
Restored Excel  ◄──[Restore]──   processed.xlsx   ◄──  AI output
```

1. **Anonymise tab** – upload your Excel, pick which columns contain private
   data, download a single ZIP that contains:
   - `anonymised.xlsx` – safe to share; contains tokens like
     `CustomerName__001`, `City__003`, …
   - `mapping_key.json` – the secret key; **keep this file safe**.

2. **Feed `anonymised.xlsx` to your AI chatbot** – ask it to process the data
   however you need, then download the processed result.

3. **Restore tab** – upload the AI-processed Excel and `mapping_key.json`;
   download the final Excel with all original values **and all original
   formatting** reinstated.

> Numeric / continuous columns are left untouched.  
> The app works directly on the Excel workbook, so fonts, colours, borders,
> merged cells, conditional formatting, charts, etc. are fully preserved.

---

## Starting the app (no terminal needed)

> **Prerequisite:** Python 3.9 or later must be installed.  
> Download it from <https://www.python.org> if needed.

The first time you launch, the app sets up a private virtual environment and
installs its dependencies automatically (takes ~1 minute). After that,
start-up is instant.

### Windows
Double-click **`launch.bat`**.

A command-prompt window will open — this is the app's server.  
A browser tab opens automatically at `http://localhost:8501`.  
**Keep the window open while you use the app.** Close it to stop.

### macOS
Double-click **`launch.command`** in Finder.

> If macOS shows a security warning the first time, right-click the file,
> choose **Open**, then click **Open** again.

A Terminal window will open — this is the app's server.  
A browser tab opens automatically at `http://localhost:8501`.  
**Keep the window open while you use the app.** Close it to stop.

### Linux / Ubuntu
Double-click **`launch.sh`** and choose **"Run as a Program"** (or similar,
depending on your file manager).

A terminal window will open and a browser tab opens automatically.  
**Keep the window open while you use the app.** Close it to stop.

> If your file manager asks whether to run or display the file, choose **Run**.  
> If nothing happens, open a terminal once and run: `chmod +x launch.sh`

---

## Mapping key format

The mapping key is plain JSON and can be inspected in any text editor:

```json
{
  "CustomerName": {
    "CustomerName__001": "Alice Smith",
    "CustomerName__002": "Bob Jones"
  },
  "City": {
    "City__001": "Berlin",
    "City__002": "Paris"
  }
}
```

---

## Project structure

```
excel_encryptor/
├── app.py              # Streamlit frontend (two-tab UI)
├── encryptor.py        # Core anonymise / restore logic (no UI dependency)
├── requirements.txt    # Python dependencies
├── smoke_test.py       # Round-trip correctness test
├── launch.bat          # Windows launcher (double-click)
├── launch.command      # macOS launcher  (double-click)
├── launch.sh           # Linux launcher  (double-click / Run as Program)
├── .streamlit/
│   └── config.toml     # Forces light theme
└── README.md
```

---

## Developer: running manually

```bash
# one-time
pip install -r requirements.txt

# start
streamlit run app.py

# test
python smoke_test.py
```

---

## Credits

Built by **Claude Sonnet**, used from within
**[Antigravity](https://deepmind.google/antigravity)** —
Google DeepMind's AI coding assistant.
