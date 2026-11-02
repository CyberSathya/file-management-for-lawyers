# File Management System for Lawyers

A Google Sheets + Apps Script web app for finding and filing physical case files kept in racks.
Built for a law office's file room, and usable for any "files on shelves" register. It works
well on a tablet at the rack.

The register is one Google Sheet with **one tab per rack**. The app searches every rack at once,
suggests the next free file number when a new file comes in, and keeps track of files taken out.

## Features

- **Search all racks at once.** Search by name, case number, place, court or file number. Results
  are grouped by rack, with filters for place, court and status, and a plain list view.
- **"Did you mean"** suggestions for misspelt or sound-alike names (Sowmya / Soumya, Prabhu / Prabu).
- **Voice search** where the browser allows it. Otherwise it falls back to Windows dictation
  (Win + H).
- **Add File:**
  - picks the rack: racks that already hold the client's files come first, then recently used ones;
  - shows each rack's status (space available / almost full / full);
  - gives the next file number in that rack;
  - warns about a possible duplicate (same case number or cause title) before saving.
- **Edit File:** correct a file's details in place. Every change is logged.
- **Take out / Returned:** record who took a file and when. The **Files out** screen lists every
  file out of the racks, longest first.
- **Clean-up report:**
  - repeated or mistyped file numbers;
  - files with no number;
  - places and courts spelt different ways;
  - racks with no status.
- **Activity log:** a hidden tab recording who added, edited, took out or returned which file.
- **Weekly backup** (optional, separate script): a dated copy of the sheet in Drive, keeping the
  last 12.

All data stays in your own Google account. The app reads and writes only the spreadsheet it is
attached to.

## The spreadsheet layout

One tab per rack ("Rack 01", "Rack 02", …, "Unsorted"). Row 1 is the header; each later row is one file:

| Column | A | B | C | D | E | F | G | H |
|---|---|---|---|---|---|---|---|---|
| Holds | File No. | Place | Court | Case No. | Cause Title | Client Name | Case Status | Date |

Notes on the layout:
- **A1** holds the rack's status: *Space Available*, *Almost Full* or *Full*. `setupRackStatus()`
  makes it a coloured dropdown.
- **File numbers** follow a rack prefix and a running number, like `01-001`. Matching ignores
  spaces and leading zeros, so `08-00009`, `08-009` and `08-9` are the same file.
- **Bundles:** a case kept in several bundles uses the same File No. on each row, with
  "Bundle 1 of 2", "Bundle 2 of 2" in Case Status. The clean-up report then doesn't flag it.
- **Notes in column A** that don't follow the rack's numbering, such as "IKEA BOX RACK 02", are
  treated as notes, not files.

`sample/` has two example racks with made-up data. Import each CSV as its own tab (File → Import
→ Upload → *Insert new sheet(s)*) and rename the tabs to "Rack 01" and "Rack 02".

## Setup

1. Create a Google Sheet laid out as above, or import the samples.
2. In the sheet: **Extensions → Apps Script**. Replace the code with `src/Code.gs`.
3. Add an HTML file named `Index` (File → + → HTML) and paste `src/Index.html` into it.
4. **Services (+) → Google Sheets API → Add.** The app uses it for fast reads across all racks.
5. Optionally, in Project Settings, show the `appsscript.json` manifest and replace it with
   `src/appsscript.json`. This also sets the web app options in step 7.
6. Run **`setupRackStatus`** once from the editor, and accept the permissions. It puts the status
   dropdown in A1 of every rack.
7. **Deploy → New deployment → Web app**:
   - *Execute as*: **User accessing the web app**. Each person works with their own access to the
     sheet, and the log shows who did what.
   - *Who has access*: whatever suits you, for example anyone in your organisation.
   - Share the sheet with the people who will use the app: **Editor** to add, edit or take out
     files, **Viewer** to search only.
8. Open the web app URL on the tablet and add it to the home screen.

### With clasp (optional)

```bash
npm install -g @google/clasp
clasp login
clasp create --type sheets --title "File Management System for Lawyers" --rootDir src
clasp push
```

`clasp create` writes `.clasp.json` with your script ID. It is git-ignored; keep it out of the
repository.

### Weekly backup (optional)

`backup/Backup.gs` runs as its **own** Apps Script project (script.google.com → New project). That
way, people using the web app are never asked for Drive access. Put the sheet's ID into
`SPREADSHEET_ID`, then run `setupBackup` once.

## Settings

At the top of `Code.gs`:

- `IGNORED_PREFIX`: a prefix some of your file numbers carry that should not count when matching.
  For example, `'ABC'` makes `ABC-14-011` the same file as `14-00011`. Set the same value at the
  top of the script in `Index.html`.
- `DEFAULT_HEADERS`, `STATUS_OPTIONS`, `MAX_RESULTS`, `TEMPLATE_RACK`: column names, rack statuses,
  the search result limit, and the rack whose dropdown colours `copyStatusDropdown()` copies.

## Licence

MIT. See [LICENSE](LICENSE).
