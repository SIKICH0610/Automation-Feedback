# Automation Feedback

A local app for writing parent-facing class feedback and pasting it into WeCom or
WhatsApp. It runs entirely on the teacher's own computer: a small web server plus a
browser interface at `http://127.0.0.1:8765`.

Nothing is ever sent automatically. The paste helper opens the right chat and fills in
the message box, then stops — a person reads it and presses Send.

## How the data is stored

`app_data/feedback.db` (SQLite) is the source of truth for classes, students, quiz
banks, and attachments. On the very first launch, an Excel workbook is imported once to
create that database; after that the app never reads it again, and editing that
spreadsheet does not change anything.

Back up the live system by copying `app_data/feedback.db` while the server is stopped.
`app_data/`, `exports/`, and `announcements/` are gitignored so student data is not
committed by accident.

## Setup

First-time setup on Windows:

```powershell
.\setup.ps1
```

First-time setup on macOS:

```bash
./setup.sh
```

This creates a local `.venv` and installs the dependencies into it, so the project never
picks up Python from somewhere else. VS Code is pointed at that interpreter.

### macOS only: Accessibility permission

WeCom paste automation on macOS drives the app through the Accessibility API, which
macOS gates behind an explicit grant. The permission has to be given to the **Python
interpreter itself**. Find its real path:

```bash
readlink -f .venv/bin/python
```

Add that exact path (not `.venv/bin/python`, which is only a symlink) in **System
Settings → Privacy & Security → Accessibility**, and switch it on.

Then **fully quit and restart the server**. A process that was already running when the
permission was granted keeps its old, unpermitted state, and every Accessibility call
fails with `osascript is not allowed assistive access`.

## Running the app

```powershell
.\.venv\Scripts\python.exe frontend_server.py
```

```bash
./.venv/bin/python frontend_server.py
```

It opens `http://127.0.0.1:8765` in the browser. Keep the terminal open while using it;
`Ctrl+C` stops the server.

## Writing feedback

### Lesson recap → the opening paragraph

The **Lesson recap 本节课内容回顾** box in the right-hand rail is what the generated
comment opens with. Your wording is used verbatim — the only thing added is the
greeting in front:

> recap: `今天的课程主要围绕三角形全等的判定、勾股定理的应用展开`
> becomes: 家长您好～今天的课程主要围绕三角形全等的判定、勾股定理的应用展开～

Left empty, a generic recap sentence fills in after the greeting instead. A recap that
already starts with 家长 is kept whole, greeting and all.

This is deliberately **separate from the Announcement box**. The announcement is text
you blast to a whole class, and reusing it here used to put things like "next week is
cancelled" at the top of every student's individual feedback.

### Generate comments

Select students, then **Generate comments**. Every message is exactly three paragraphs:
the greeting with the lesson recap; the student's classroom performance, built from the
dropdown observations and the teacher's `Remark for Student` note; and a final paragraph
with the standing homework/app/coin note, any `Homework Reflection`, and the closing
line. Results land in the `Feedback` column, where you can edit them before pasting —
click a long cell and it floats open as a full-size box in place.

Absent students are skipped. Fields marked `Not Observed` are left out.
`Additional Comment` is appended to the end of the personal paragraph.

### Quiz feedback

Pick **Quiz 1** or **Quiz 2** in the Quiz block. Each quiz has **its own recap box**,
which follows that selector, and is used as paragraph 1 of that quiz's message only —
the two quizzes are written up and sent separately.

**Generate quiz feedback** then writes into `Quiz1 Feedback` / `Quiz2 Feedback`, using
only that quiz's score and average. The quiz you picked is passed through explicitly, so
a student who has both quizzes recorded can no longer get Quiz 2's score reported inside
their Quiz 1 message.

Open **Quiz Banks** in the top navigation to edit the per-question feedback: question
numbers, titles, matching patterns, and the Chinese/English wording. Banks live in the
same database; the Python bank files seed a bank only the first time it is created, and
after that the editor is the live source.

## Sending

All paste actions open the chat, fill the message box, and stop. They never press Send.

- **Paste comments** — pastes each selected student's `Feedback`.
- **Paste quiz feedback** — pastes `Quiz1`/`Quiz2 Feedback` for the selected quiz.
- **Paste announcement** — pastes one shared message, plus any files ticked under
  **Attachments**, into each selected student's chat. Attachment runs handle one student
  at a time, since the file preview stays open for review.

Which app a student routes to comes from `Preferred Channel`, or from `Parent Language`
when that is blank (Chinese → WeCom, otherwise WhatsApp).

`Send Status`, `Send Error`, and `Last Attempt` are written back per row. Status values
are `pasted`, `needs_review`, `skipped_absent`, and `failed`.

### Check group chat status

**Check group chat status** searches for each selected student's chat by `uid` and
writes `TRUE`/`FALSE` into the `Group Chat` column. It never opens or pastes into a chat.
Rows that could not be checked are reported as `needs_review` and left unchanged rather
than guessed.

For a student with neither `Preferred Channel` nor `Parent Language` set — the usual case
right after an import — it tries WeCom first, then WhatsApp, and fills in
`Parent Language` from whichever one found the chat, so later actions route correctly
without anyone setting it by hand. If neither finds a match, the language is left blank.

The **Bulk actions 跨班级** panel in the right-hand rail runs across whole class
rosters instead of selected rows — every student of every ticked class in the current
semester. Untick any class you want to skip; unticked classes are remembered as you move
between tabs. Three actions share the class list:

- **Generate comments** — regenerates the `Feedback` column for everyone, each class
  opening with its own lesson recap. Overwrites what was there.
- **Check group chats** — the check above, semester-wide.
- **Paste comments** — the supervised paste, one student at a time, never pressing Send.

The two desktop actions take care of WeCom themselves: its window is brought up (or
restored from the Dock) before the batch starts, and minimized again when the batch
finishes, so WeCom can just sit in the background between runs. WhatsApp is still
checked up front and must already be open — the batch stops immediately with a clear
message if a needed app is unavailable, and never opens anything mid-batch. Waits are adaptive: fast machines move on as soon as WeCom's UI confirms each step,
slow machines get longer caps than the old fixed delays — budget two to four seconds
per student depending on the machine. Long runs show a live progress bar with an
estimated time remaining and a **Cancel** button — cancelling stops after the row in
flight, keeps everything already written, and still tucks WeCom away. The automation
itself must stay in the foreground while it runs: it works by sending real keystrokes,
and macOS delivers those only to the frontmost window, so true background operation is
not possible with this approach.

## Managing the roster

### Import a semester's enrollment

The **Import Students** tab reads a platform enrollment export — `.xlsx` (including the
platform's Strict OOXML variant), `.csv` (UTF-8 or GBK), or `.json` (a bare list of
records, or wrapped in an API envelope like `{"data": {"list": [...]}}`) — with the
format detected from the file's content, not its name. Required fields: `classId`,
`className`, `classTimeDescription`, `firstName`/`lastName`, `学员id`, `payStatus`,
`是否入班`. The reader is deliberately
forgiving about the file's shape: columns are matched by name (with common renames and
case/spacing differences accepted), extra columns and reordering are ignored, the header
row is located by content even under a banner row, and instruction sheets before the
data sheet are skipped — so most platform export changes need no app update. Only
`payStatus = Paid` and `是否入班 = 是` rows are imported. Choose the file, name the semester (or pick an existing
one), **Preview** to see the classes and counts found, untick anything you don't want,
then **Commit Import**.

Parent phone, email, and address are deliberately not imported — contact happens through
chat search by `uid`, so `WhatsApp Phone` is only filled in by hand for families that
need it.

Re-running the same file is safe: semesters and classes are matched by name / platform
`classId`, and students already on the roster (matched by `uid`) are left untouched
rather than duplicated.

**Overwrite existing roster** instead makes a class's roster exactly match the file. A
student on the roster but missing from the file is **permanently deleted**, so it asks
for explicit confirmation whenever the preview shows anyone would be removed. Students in
both keep everything already recorded about them — feedback, quiz scores, teacher notes,
attendance — and only their name and `Group Chat` are refreshed. A timestamped database
backup is written automatically before any deletion.

### Delete a student, a class, or a semester

- **A student**: the `×` at the end of their row. Confirms by name.
- **A class**: **Delete class** in the roster toolbar. Removes its students, columns,
  attachments, and announcement file, and confirms with the student count first.
- **A semester**: the **Delete "…"** button beside the semester dropdown. Shows exactly
  which classes and how many students would go, and makes the same automatic backup.

### Export

Both export buttons download a real `.xlsx`, scoped to the semester selected in the
top-left dropdown, one class per sheet:

- **Export Excel** → every internal column. The full round-trip copy.
- **Export Report** → the curated status view: Name, Student ID, 电话号码, 是否有群,
  是否发开课提醒, 是否发课后反馈, 第一节课反馈, 第一次quiz反馈, 第二次quiz反馈. For
  reporting, not for re-importing.

Editing a downloaded file does not change the database. Scripted callers can use
`POST /api/export` or `POST /api/export/report` with an optional `{"semester": "..."}`
body, which writes into `exports/` and returns the path instead of streaming a download.

## Building the macOS app

```bash
./build_mac.sh
```

That builds `dist/Teacher Feedback Desk.app`, signs every nested binary with the
"Think Academy Automation" certificate from the login keychain, and wraps it in
`dist/Teacher Feedback Desk.dmg`. Signing every release with that same certificate is
what lets macOS keep the Accessibility grant across updates — teachers replace the app
and their permissions carry over. (The deep re-sign step matters: a cert-signed
launcher refuses to load the bundled Python.framework while it still carries
python.org's signature, and the app dies instantly at startup.)

The packaged app keeps its data in `~/Library/Application Support/Teacher Feedback
Desk/` — database, announcements, exports, and an `app.log` with startup lines and
crash tracebacks (a windowed app has no console). To seed a fresh install with an
existing roster, copy the project's `app_data/feedback.db` into that folder's
`app_data/` before first launch; otherwise the bundled workbook is imported once.

Launch it by double-clicking (or `open`); running the binary inside `Contents/MacOS`
directly starts the server but never shows the window. On another Mac the ad-hoc
signature means the first launch needs right-click → Open. The app needs its own
Accessibility and Automation grants in System Settings — the ones given to the dev
Python don't carry over — and, as always, quit and reopen the app after granting.

## Platform support

| | Windows | macOS |
|---|---|---|
| Roster, feedback generation, quiz banks, import/export | ✅ | ✅ |
| WeCom paste + group chat check | ✅ `pywinauto` | ✅ Accessibility API |
| WhatsApp paste + group chat check | ✅ | ❌ returns `needs_review` |
| Attachment paste | ✅ | ❌ skipped with a notice |

`pywinauto` is Windows-only and is skipped by `requirements.txt` elsewhere. macOS uses
`wecom_mac.py`, which drives WeCom through `osascript`/JXA instead.

**How each platform verifies it found the right chat** — the two are genuinely different,
because the two WeCom builds expose very different things:

- **Windows.** WeCom draws its own UI rather than using real controls, so UI Automation
  exposes no readable text at all, and OCR proved unreliable (WeCom echoes the search term
  back in an always-present "search online" suggestion, and other sidebar contacts can
  coincidentally contain a name fragment). What is reliable: a genuine match adds a result
  row above that suggestion, making the dropdown measurably taller. The check diffs a
  screenshot from immediately before and after typing the uid and measures the changed
  region — no text is read. Because that dropdown is a fixed pixel size that does not
  scale with the window, each run measures its own baseline first (a uid guaranteed not to
  exist) and compares as a ratio, so it self-calibrates to the machine's window size and
  display scaling.
- **macOS.** WeCom for Mac does expose a real accessibility tree, so the check reads the
  sidebar's own selection state and matches the chat title against the student's uid and
  name directly.

## Command line

The app covers everything below; these are for scripting and debugging. They act on an
Excel workbook passed with `--workbook`, **not** on the live database, unless launched
through the app's own buttons.

Preview feedback without writing anything:

```powershell
python feedback_generator.py --sheet "Geo TTh" --row 2 --class-review "三角形全等的判定"
python feedback_generator.py --sheet "Geo TTh" --all --start-row 2 --end-row 5 --class-review "..."
python feedback_generator.py --sheet "Geo TTh" --all --class-review "..." --review-csv preview.csv
```

Add `--write` to save into the feedback column. Choose the kind with
`--feedback-type general | quiz | comprehensive`, and for quiz runs pass
`--quiz-number 1 | 2` so the right quiz is used.

Import enrollment:

```powershell
python import_enrollment.py --source-file ".\enrollment.xlsx" --semester "2026-27 School Year"
python import_enrollment.py --source-file ".\enrollment.xlsx" --semester "2026-27 School Year" --commit
python import_enrollment.py --source-file ".\enrollment.xlsx" --semester "Fall 2026" --class-id 123789 --overwrite --commit --confirm-delete
```

Supervised paste (`--mode paste-only`; the default `dry-run` only prints):

```powershell
.\.venv\Scripts\python.exe paste_sender.py --sheet "Geo TTh" --row 2 --status
.\.venv\Scripts\python.exe paste_sender.py --sheet "Geo TTh" --rows 3,5 --mode paste-only
.\.venv\Scripts\python.exe paste_sender.py --sheet "Geo TTh" --start-row 2 --end-row 10 --mode paste-only --action check-group-chat
.\.venv\Scripts\python.exe paste_sender.py --sheet "Geo TTh" --row 2 --action mass-notification --mass-message-file notice.txt --attachment ".\handout.pdf" --mode paste-only
```

`--action comment` (default) pastes each row's feedback; `--action mass-notification`
pastes one shared message everywhere; `--action check-group-chat` only checks. Add
`--channel wecom|whatsapp` to force which app a check searches, `--check-apps` to print
app availability as JSON and exit, or `--require-verification` to stop rather than paste
when the chat cannot be confirmed.

On Windows, `--wecom-exe` or the `WECOM_EXE` variable points at a non-standard WeCom
install, and `--debug-search-results` prints the candidates the search sees.

### Common options

- `--workbook` — path to the Excel file.
- `--sheet` — sheet name.
- `--row` / `--rows` / `--start-row` / `--end-row` / `--all` — which students.
- `--class-review` / `--class-review-file` — the opening paragraph.
- `--feedback-type` — `comprehensive` (default), `general`, or `quiz`.
- `--quiz-number` — `1` or `2`. Without it the quiz is inferred from each row's data.
- `--write` — save back to the workbook.

## Project files

**App**
- `frontend_server.py` — the local server and action runner.
- `frontend/` — the browser interface.
- `database_store.py` — SQLite roster, recaps, one-time Excel import, exports, backups.
- `attachment_store.py` — per-class attachment storage.
- `quiz_bank_store.py` — editable quiz banks.
- `import_enrollment.py` — enrollment import, also usable from the command line.

**Feedback wording**
- `feedback_master.py` — assembles a message from the pieces below.
- `feedback_common.py` — shared student-row, wording, and formatting helpers.
- `feedback_general.py` — regular classroom feedback.
- `feedback_quiz.py` — quiz score parsing and quiz wording.
- `feedback_generator.py` — command-line front end for the generators.
- `geometry_volume1_quiz1_comment_bank.py`, `geometry_volume1_quiz2_comment_bank.py`,
  `amc10_quiz1_comment_bank.py` — default banks used to seed the database.

**Paste automation** — one module per app per platform, with `paste_sender.py`
choosing between them at run time. The Windows modules are imported only on Windows,
so `pywinauto` is never needed elsewhere.

- `paste_sender.py` — the supervised paste CLI, job building, and the platform
  dispatch. Never sends.
- `paste_common.py` — the job/status types and clipboard helpers every robot shares.
- `wecom_win.py` / `wecom_mac.py` — WeCom, via UI Automation and the Accessibility
  API respectively.
- `whatsapp_win.py` — WhatsApp on Windows.
- `whatsapp_mac.py` — in-progress WhatsApp support for macOS; not wired up yet.
- `paste_comment.py`, `paste_mass_notification.py` — which text a row should paste.
- `paste_attachments.py` — Windows file-clipboard staging.

**Other**
- `workbook_setup.py` — adds the optional helper columns to a legacy Excel workbook.
- `setup.ps1`, `setup.sh` — first-time environment setup.
- `Geo_TTh_Student_Script_fixed_rows_only.xlsx` — the one-time import source used to
  create the database on a fresh install.
