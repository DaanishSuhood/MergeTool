# MergeTool

A Windows mail-merge application written in AutoHotkey v2. It drives Microsoft Word, Excel, and Outlook through COM automation to turn one Word template plus one Excel data source into personalised emails, individual files, or printed output.

Everything runs from a single script, `MergeTool.ahk`, which is also compiled to a standalone `MergeTool.exe` (available from [Releases](https://github.com/DaanishSuhood/MergeTool/releases/latest)).

---

## Requirements

| Requirement | Notes |
|---|---|
| Windows | Tested on Windows 10/11 |
| [AutoHotkey v2.0+](https://www.autohotkey.com/) | Only needed to run the `.ahk` script; `MergeTool.exe` runs without it. The script declares `#Requires AutoHotkey v2.0` and will not run under v1 |
| Microsoft Word | Required for every merge type (template opening, merge execution, PDF export) |
| Microsoft Excel | Required only for the **Generate Contact Sheet** button |
| Microsoft Outlook (desktop) | Required for Email Merge; must be configured with at least one account |
| Microsoft Access Database Engine (ACE OLEDB) | Usually installed with Office. Used to read the Excel data source |

## Getting started

- **Exe:** download `MergeTool.exe` from the [latest release](https://github.com/DaanishSuhood/MergeTool/releases/latest) and double-click it. No AutoHotkey install needed.
- **Script:** install AutoHotkey v2, then double-click `MergeTool.ahk` (or run `AutoHotkey.exe MergeTool.ahk`).

The MergeTool window opens; settings from the previous session are restored automatically.

### Rebuilding the exe

After editing the script, recompile with [Ahk2Exe](https://github.com/AutoHotkey/Ahk2Exe/releases) (right-click `MergeTool.ahk` → *Compile Script* once Ahk2Exe is installed, or from the command line):

```
Ahk2Exe.exe /in MergeTool.ahk /out MergeTool.exe /base "C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe"
```

The icon (`MergeTool.ico`) and version info are set by the `;@Ahk2Exe-…` directives at the top of the script, so keep the `.ico` next to the script when compiling.

---

## How it works

### 1. Choose your files

- **Template** — a Word `.docx` or `.doc` containing merge fields (`«Name»`, `«Email»`, …) inserted via Word's *Mailings → Insert Merge Field*.
- **Data Source** — an Excel `.xlsx`, `.xlsm`, `.xlsb`, or `.xls` workbook. Data is read from the workbook's **first worksheet**, whatever it's named, with the first row treated as column headers.

If you don't have a contact list yet, **Generate Contact Sheet…** creates a new workbook pre-populated with the columns `Name`, `Email`, `Subject`, `CC`, `BCC`, `Attachment`, opens it in Excel, and selects it as the data source.

### 2. Load fields

Click **Load Fields**. The tool opens the template read-only, connects it to the workbook, and reads the column names into every field dropdown. The status bar reports how many fields and records were found.

Field loading must succeed before a merge can start.

### 3. Pick a merge type

Selecting a merge type swaps the options panel and resizes the window to fit; the choice is remembered between sessions.

---

## Merge types

### Email Merge

Renders the template for each record as HTML and sends it through Outlook. Emails go to the Outlook **Outbox** via `MailItem.Send()`.

| Option | Purpose |
|---|---|
| **Send From** | Which configured Outlook account to send as. **Refresh** re-reads the account list from Outlook. |
| **Delay** | Seconds to wait between sends (0–600, default 15). Helps avoid mail-server rate limits and spam blocks. |
| **Email Field** | Column holding the recipient address. Required. |
| **Subject** | Column holding the subject line. Required. |
| **CC Fields / BCC Fields** | Up to three columns each; non-empty values are joined with `;`. |
| **Attach Fields** | Up to three columns holding per-recipient file paths. Paths that don't exist are skipped. |
| **Common Attach** | One file attached to every message. |

The email body is the template merged for that record, exported by Word as *Filtered HTML*. Images from Word's `_files` folder are inlined as base64 data URIs so they survive after the temporary files are deleted.

### Document Merge

Creates one PDF file per record.

| Option | Purpose |
|---|---|
| **Filename Field** | Column used to name each file |
| **Prefix** | Optional string prepended to every filename |
| **Output Folder** | Destination; created if it doesn't exist |

Filenames are sanitised (invalid Windows characters replaced with `_`, trimmed to 200 characters, empty values become `Document`), and collisions are resolved by appending `_1`, `_2`, … PDF export is limited to the template's page count to avoid blank trailing pages.

When the merge finishes you're offered a shortcut to open the output folder.

### Printer Merge

| Option | Purpose |
|---|---|
| **Printer** | Chosen from the installed printers; the selection is remembered |
| **Copies** | 1–99 copies of each record |
| **Collate copies** | On: each copy prints as a complete set. Off: copies of each page print together |

All records merge into a single document, which is then sent to the printer with `PrintOut` so copy count and collation apply to the job as a whole. A confirmation dialog showing printer, record count, and total documents appears before anything prints.

---

## Progress, cancelling, and errors

- The progress bar, status line, and detail line update per record.
- **Cancel** stops document and email merges at the next record boundary. (A print job is a single operation and can only be cancelled before it's sent.)
- Per-record failures are counted rather than aborting the run. A summary dialog at the end lists successes, failures, and up to the first 10 error messages.

## Saved preferences

Settings are written to `MergeTool_config.ini` next to the script or exe, under a `[Settings]` section:

| Key | Saved value |
|---|---|
| `LastTemplate` | Last selected Word template |
| `LastDataSource` | Last selected Excel workbook |
| `LastOutputFolder` | Last document-merge output folder |
| `LastMergeType` | `email`, `documents`, or `print` |
| `LastPrinter` | Last selected printer |
| `LastSendFromAccount` | Last selected Outlook account |
| `LastEmailDelaySeconds` | Delay between email sends |

Paths are only restored if the file still exists. Delete the INI file to reset everything.

---

## Troubleshooting

**"Failed to open data source"**
The tool tries four connection methods in order: Word's own auto-detection, `Microsoft.ACE.OLEDB.16.0`, `Microsoft.ACE.OLEDB.12.0`, and (for `.xls`) `Microsoft.Jet.OLEDB.4.0`. The error lists what each attempt reported. Common causes:
- The workbook is open in Excel — close it and retry.
- Your data is on a later sheet — only the first worksheet is read. Move it to the front of the workbook.
- No ACE OLEDB provider installed; install the Microsoft Access Database Engine matching your Office bitness.

**No fields listed after Load Fields**
The first row of the first worksheet must contain header names, and the sheet must contain at least one data row.

**Outlook account dropdown only shows "(Default Account)"**
Outlook isn't running or isn't configured. Start Outlook, then click **Refresh**.

**Emails sit in the Outbox**
That's by design — messages are handed to Outlook, which sends them on its own schedule. Press Send/Receive in Outlook to flush the Outbox immediately.

**Merge fields print literally instead of being replaced**
The template's fields must be real Word merge fields, not typed text. Insert them from Word's *Mailings* ribbon.

**Word or Excel processes left running**
Every merge closes documents and quits its application in a `finally` block, but a hard kill of the script can orphan a process. End `WINWORD.EXE` / `EXCEL.EXE` in Task Manager if that happens.

---

## Notes on the code

`MergeTool.ahk` is organised in labelled sections: COM constants, GUI construction (one creation function per options panel plus show/hide/layout helpers), data-source access (`OpenDataSource`, `EnumerateFields`, `GetRecordCount`, `GetFieldValue`, `SetActiveRecord`), a Word session wrapper (`InitializeWordSession` / `CleanupWordSession`), one merge routine per type (`MergeToDocuments`, `MergeToEmailDirect`, `MergeToPrinter`), and preference persistence.

`MergeToEmail` and `CreateOutlookDraft` implement a draft-creating variant of the email merge (messages saved to Drafts rather than sent). They are complete but not wired to the GUI, which calls `MergeToEmailDirect` instead.

---

## License

[MIT](LICENSE)
