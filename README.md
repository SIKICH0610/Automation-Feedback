# Automation Feedback

Generate one parent-facing feedback entry from the provided Excel tracker.

## Project Files

- `database_store.py`: SQLite roster storage, one-time Excel migration, temporary action workbooks, and safe exports.
- `frontend_server.py`: local browser interface and action runner.
- `frontend/`: editable roster, announcements, quiz fields, feedback generation, paste controls, and Excel export.
- `feedback_generator.py`: CLI wrapper for previewing or writing feedback into the existing `Feedback` cells.
- `feedback_common.py`: shared workbook, student-row, wording, and formatting helpers.
- `feedback_general.py`: regular classroom-feedback wording.
- `feedback_quiz.py`: quiz-score parsing and quiz-feedback wording.
- `quiz_bank_store.py`: encapsulated SQLite storage, validation, matching, and runtime access for editable quiz banks.
- `geometry_volume1_quiz1_comment_bank.py`: default Quiz 1 bank used for first-time database seeding.
- `geometry_volume1_quiz2_comment_bank.py`: default Quiz 2 bank used for first-time database seeding.
- `feedback_master.py`: master generator that can call general, quiz, or comprehensive feedback.
- `paste_sender.py`: CLI wrapper for supervised paste actions. It checks app status and never sends.
- `paste_comment.py`: row-specific comment payload selection.
- `paste_mass_notification.py`: shared mass-notification payload selection.
- `attachment_store.py`: private SQLite attachment storage for each class.
- `paste_attachments.py`: coordinate-free Windows file clipboard staging.
- `class_review_builder.py`: create or overwrite `class_review.txt` from teacher text, notes, or slides.
- `import_enrollment.py`: import a platform enrollment export into the SQLite roster, grouped into a named semester block.
- `workbook_setup.py`: prepare optional workbook columns such as `Additional Comment`.
- `openai_api.py`: shared OpenAI API helper.
- `class_review.txt`: editable class-level paragraph used as paragraph 1.
- `Geo_TTh_Student_Script_fixed_rows_only.xlsx`: one-time import source for the frontend and default workbook for direct CLI commands.

The initial Excel feedback form lives in this project folder:

```text
Geo_TTh_Student_Script_fixed_rows_only.xlsx
```

On the first frontend launch, its sheets and student rows are imported into `app_data/feedback.db`. After that migration, the frontend, generators, and paste actions use SQLite as the live source of truth. Editing the original `.xlsx` file does not change the frontend database.

Direct command-line scripts still use the workbook passed with `--workbook`. Those legacy CLI writes do not update the frontend database unless they are launched through the frontend action buttons.

## Setup

First-time setup on Windows:

```powershell
.\setup.ps1
```

This creates a local `.venv` folder and installs the required packages there. VS Code is configured to use `.venv\Scripts\python.exe` for this workspace, so it will not accidentally use Python from another project.

Optional API key for GPT-assisted workflows:

```powershell
$env:OPENAI_API_KEY="your_api_key_here"
```

## Local frontend

Start the editable workbook interface:

```powershell
.\.venv\Scripts\python.exe frontend_server.py
```

The app opens at `http://127.0.0.1:8765`. Each class sheet has its own linked UTF-8 announcement file in `announcements`. The frontend saves student edits and uploaded images or documents into `app_data/feedback.db`, generates general or quiz feedback, and runs the supervised paste-only workflow for selected rows.

Open **Quiz Banks** in the top navigation to edit question numbers, titles, matching patterns, and Chinese or English feedback. Quiz bank entries are stored in the same `app_data/feedback.db` database. The Python bank files are imported only when the corresponding database bank is first created. After that, frontend edits are the live source used by quiz-feedback generation.

The original workbook is imported only when the database is created for the first time. Use **Export Excel** in the roster toolbar when you need a spreadsheet copy. The export is written to `exports/Student_Feedback_Export.xlsx`. Editing that exported file does not change the live database.

To back up the live system, copy `app_data/feedback.db` while the frontend is stopped. The `app_data`, `exports`, and `announcements` folders are ignored by Git so private student data is not committed accidentally.

Close the terminal or press `Ctrl+C` to stop the local server.

## Prepare the workbook

Add the optional automation helper columns to every sheet:

```powershell
python workbook_setup.py
```

This prepares `Additional Comment`, `Preferred Channel`, WhatsApp routing fields, and send audit fields. Use `Additional Comment` for extra notes that should be added to the end of paragraph 2. In API mode, the note is translated when needed.

Routing columns:

- `Preferred Channel`: optional `wecom` or `whatsapp`. Blank uses `Parent Language`: Chinese -> WeCom, non-Chinese -> WhatsApp.
- `WhatsApp Phone`: parent phone number for direct WhatsApp phone URL mode.
- `WhatsApp Search Key`: optional group-chat search key. Blank uses `uid`.
- `WhatsApp Target Type`: optional `phone` or `group_search`. Blank uses `phone` when `WhatsApp Phone` is present, otherwise `group_search`.

## Import a new semester's enrollment

`import_enrollment.py` reads a platform enrollment export (one row per paid, enrolled student order, with `classId`, `className`, `classTimeDescription`, `firstName`/`lastName`, and `学员id`) and adds any classes and students it finds into the SQLite roster, grouped under a semester block you name. Only `payStatus = Paid` and `是否入班 = 是` rows are imported. Parent phone, email, and address are intentionally not imported; contact happens through WeCom/WhatsApp chat search by `uid`, so staff add `Preferred Channel` / `WhatsApp Phone` manually for a family that needs it.

Preview only, nothing is written:

```powershell
python import_enrollment.py --source-file ".\enrollment_export.xlsx" --semester "2026-27 School Year"
```

Commit the previewed changes:

```powershell
python import_enrollment.py --source-file ".\enrollment_export.xlsx" --semester "2026-27 School Year" --commit
```

Import a single new class into a semester block you already created, instead of every class in the file:

```powershell
python import_enrollment.py --source-file ".\enrollment_export.xlsx" --semester "2026-27 School Year" --class-id 123790 --commit
```

Re-running the same file (or a corrected re-export) is safe: an existing semester and class are reused by name / platform `classId`, and a student already on the roster (matched by `uid`) is left untouched rather than duplicated.

The frontend has the same workflow under the **Import Students** tab: choose the `.xlsx` file, type a new semester name (or pick an existing one from the suggestions) to add classes into it, click **Preview** to see the classes and student counts found, uncheck any class you don't want, then **Commit Import**.

### Overwrite an existing roster from a new file

Add `--overwrite` to sync a matched class's roster to exactly match the uploaded file, instead of only adding new students:

```powershell
python import_enrollment.py --source-file ".\enrollment_export.xlsx" --semester "Fall 2026" --class-id 123789 --overwrite --commit --confirm-delete
```

A student currently on that class's roster but missing from the file is **permanently deleted**; `--confirm-delete` is required whenever the preview shows anyone would be removed, as an explicit acknowledgment. A student present in both gets their name and `Group Chat` refreshed from the file, but every other field (feedback, quiz scores, teacher notes, attendance) is left untouched — overwrite only syncs who's on the roster, not what's already been recorded about them. A timestamped backup of the whole database (`app_data/feedback.*.before-overwrite-*.db`) is made automatically before any deletion runs. The frontend has the same option as the **Overwrite existing roster** checkbox on the Import Students tab, with the same confirmation step.

### Delete a semester

The semester dropdown (top-left of the class tabs, once you have more than one semester) has a **Delete "<semester name>"** button next to it. It shows exactly which classes and how many students would be permanently removed before you confirm, makes the same automatic backup, and keeps you on the roster view afterward — falling back to the first remaining class rather than reloading the page.

## Export a semester's roster

Both export buttons in the roster toolbar download a real `.xlsx` file through the browser, scoped to **the semester currently selected in the top-left dropdown**, with one class per sheet:

- **Export Excel** → `Student_Feedback_Export_<Semester>.xlsx`: every internal column, formatted from the workbook template. This is the full round-trip copy.
- **Export Report** → `Student_Report_Export_<Semester>.xlsx`: the curated status view with just Name, Student ID, 电话号码 (`WhatsApp Phone`, blank unless set), 是否有群 (`Group Chat`), 是否发开课提醒 (`Before Class Informing`), 是否发课后反馈 (whether `Send Status` is `pasted`), 第一节课反馈 (`Feedback`), 第一次quiz反馈 (`Quiz1 Feedback`), and 第二次quiz反馈 (`Quiz2 Feedback`). Meant for reporting, not for re-importing.

Editing either downloaded file does not change the database. The pre-semester classes shown as "Other Classes" export the same way, as `..._Other_Classes.xlsx`.

Scripted callers can still use `POST /api/export` or `POST /api/export/report` with an optional `{"semester": "Fall 2026"}` body, which writes the file into `exports/` and returns its path instead of streaming a download. Omitting `semester` there exports every class across all semesters.

## Create the class review file

The editable class review file lives in this project folder as `class_review.txt`.
Teachers can edit that file directly before generating feedback.
When new slides or class notes are provided, run `class_review_builder.py` again and it will overwrite `class_review.txt`.

Teacher-written text:

```powershell
python class_review_builder.py --source-text "Today we discussed triangle similarity, matching corresponding angles and sides, and setting up proof statements from diagrams." --output class_review.txt
```

From a text or PowerPoint file without the API:

```powershell
python class_review_builder.py --source-file ".\lesson_notes.txt" --output class_review.txt
python class_review_builder.py --source-file ".\lesson_slides.pptx" --output class_review.txt
```

From slides or class material with the OpenAI API:

```powershell
$env:OPENAI_API_KEY="your_api_key_here"
python class_review_builder.py --source-file ".\lesson_slides.pdf" --output class_review.txt --use-api
python class_review_builder.py --source-file ".\lesson_slides.pptx" --output class_review.txt --use-api
```

The generated `class_review.txt` is copied directly as paragraph 1 of the parent message.

## Preview one entry

```powershell
python feedback_generator.py --sheet "Geo TTh" --row 2 --class-review-file class_review.txt
```

The command previews the generated feedback without changing the workbook.

## Preview several or all entries

Preview rows 2 through 5:

```powershell
python feedback_generator.py --sheet "Geo TTh" --all --start-row 2 --end-row 5 --class-review-file class_review.txt
```

Save that preview to a spreadsheet-friendly CSV for review:

```powershell
python feedback_generator.py --sheet "Geo TTh" --all --start-row 2 --end-row 5 --class-review-file class_review.txt --review-csv review_preview.csv
```

Preview every student row in the sheet:

```powershell
python feedback_generator.py --sheet "Geo TTh" --all --class-review-file class_review.txt
```

Choose which kind of feedback to generate:

```powershell
python feedback_generator.py --sheet "Geo TTh" --row 2 --class-review-file class_review.txt --feedback-type general
python feedback_generator.py --sheet "Geo TTh" --row 2 --class-review-file class_review.txt --feedback-type quiz
python feedback_generator.py --sheet "Geo TTh" --row 2 --class-review-file class_review.txt --feedback-type comprehensive
```

`general` writes the course description plus regular classroom feedback. `quiz` writes the course description plus quiz-focused feedback. `comprehensive` writes the course description plus both quiz and regular feedback. Writing with `--write` updates the existing `Feedback` cells in the same workbook; it does not create a copy of the sheet.

## Write one entry back to Excel

```powershell
python feedback_generator.py --sheet "Geo TTh" --row 2 --class-review-file class_review.txt --write
```

This writes the generated text into the `Feedback` column for that row.

## Write all entries back to Excel

```powershell
python feedback_generator.py --sheet "Geo TTh" --all --class-review-file class_review.txt --write
```

Test a smaller range before writing everyone:

```powershell
python feedback_generator.py --sheet "Geo TTh" --all --start-row 2 --end-row 5 --class-review-file class_review.txt --write
```

## API-assisted student comments

Revise the Chinese teacher note in `Remark for Student`, preview the parent comment, and leave the sheet unchanged:

```powershell
python feedback_generator.py --sheet "Geo TTh" --row 2 --class-review-file class_review.txt --revise-remark --use-api
```

Revise the Chinese teacher note, save it back to `Remark for Student`, generate the parent comment, and write it to `Feedback`:

```powershell
python feedback_generator.py --sheet "Geo TTh" --row 2 --class-review-file class_review.txt --revise-remark --write-revised-remark --use-api --write
```

## Supervised paste helper

Check one row and the needed desktop app without pasting:

```powershell
.\.venv\Scripts\python.exe paste_sender.py --sheet "Geo TTh" --row 2 --class-review-file class_review.txt --status
```

Open the needed app if the script can find it:

```powershell
.\.venv\Scripts\python.exe paste_sender.py --sheet "Geo TTh" --row 2 --class-review-file class_review.txt --open-app
```

Search WeCom by UID, press Enter to open the first relevant result, focus the message box, and paste without sending:

```powershell
.\.venv\Scripts\python.exe paste_sender.py --sheet "Geo TTh" --row 2 --class-review-file class_review.txt --mode paste-only
```

Paste a mixed batch without sending. Rows route to WeCom or WhatsApp from the workbook; missing contacts are marked `needs_review` and the batch continues:

```powershell
.\.venv\Scripts\python.exe paste_sender.py --sheet "Geo TTh" --start-row 2 --end-row 8 --class-review-file class_review.txt --mode paste-only
```

Paste one shared mass notification to each selected chat without sending:

```powershell
.\.venv\Scripts\python.exe paste_sender.py --sheet "Geo TTh" --start-row 2 --end-row 8 --action mass-notification --mass-message-file notice.txt --mode paste-only
```

Stage an announcement and local files in one chat without sending:

```powershell
.\.venv\Scripts\python.exe paste_sender.py --sheet "Geo TTh" --row 2 --action mass-notification --mass-message-file notice.txt --attachment ".\handout.pdf" --attachment ".\diagram.png" --mode paste-only
```

The frontend provides the same workflow through **Add files** and **Paste announcement**. Attachments are staged with the Windows file clipboard, so the process does not depend on window coordinates. The chat preview remains open for manual review and the robot never presses Send. Because that preview blocks navigation, attachment runs process one selected student at a time.

`--action comment` is the default and pastes each row's `Feedback` value. `--action mass-notification` pastes the same shared text for every selected row. Shared text can come from `--mass-message`, `--mass-message-file`, or, if neither is provided, `--class-review-file`.

The paste helper recognizes window titles containing `WeCom`, `企业微信`, or `WXWork` by default. It does not press Enter after pasting, and it does not use calculated screen-position clicks. If WeCom / 企业微信 is installed in a custom location, pass `--wecom-exe "C:\path\to\WXWork.exe"` or set `WECOM_EXE`. If Enter cannot open the result on a computer, try `--ui-control-result-open` or `--manual-result-click`. Add `--require-verification` if you want the script to stop whenever it cannot verify the chat by UI text.

Paste multiple specific rows automatically without sending:

```powershell
.\.venv\Scripts\python.exe paste_sender.py --sheet "Geo TTh" --rows 3,5 --class-review-file class_review.txt --mode paste-only
```

You can also use a row range:

```powershell
.\.venv\Scripts\python.exe paste_sender.py --sheet "Geo TTh" --start-row 3 --end-row 5 --class-review-file class_review.txt --mode paste-only
```

The sender writes `Send Status`, `Send Error`, and `Last Attempt` unless `--no-status-write` is passed. Status values include `pasted`, `needs_review`, `skipped_absent`, and `failed`.

For WhatsApp group search, WhatsApp Desktop or an active WhatsApp Web browser tab can be used. If WhatsApp is in a browser tab, make that tab active before running paste-only, or run `--open-app` to open `https://web.whatsapp.com/`.

To inspect which safe WeCom search candidates the script sees:

```powershell
python paste_sender.py --sheet "Geo TTh" --row 2 --class-review-file class_review.txt --debug-search-results
```

## Check group chat status

`--action check-group-chat` searches WeCom or WhatsApp for each selected row's chat by `uid`, the same safe search used by `--debug-search-results`, but never opens or pastes into a chat. It writes `TRUE`/`FALSE` back to the existing `Group Chat` column: `TRUE` when a matching chat is found, `FALSE` when it is not. Rows that could not be checked (app not available, WhatsApp phone-target rows, etc.) are left unchanged and reported as `needs_review` instead of being guessed.

For a row with no `Preferred Channel` and no `Parent Language` set (the common case for freshly imported students), the check searches WeCom first, since most families are Chinese-speaking; only if that finds nothing does it check WhatsApp. Whichever channel actually finds the chat also fills in `Parent Language` (`Chinese` for WeCom, `English` for WhatsApp), so future comment/paste actions route correctly without anyone having to set it by hand. If neither channel finds a match, `Parent Language` is left blank rather than guessed. Rows that already have a channel or language configured are checked on that single channel only, unchanged from before.

**WeCom verification note:** WeCom renders its entire UI as custom-drawn graphics rather than real controls, so Windows' UI Automation exposes no readable text from it at all, regardless of what's on screen — and reading the screen via OCR turned out to be unreliable too, since WeCom always echoes the raw search term back in a "Search for mobile number/email online: ..." suggestion even when nothing matches, and the sidebar's many other real contacts can coincidentally contain a name fragment being checked, producing false positives for uids that don't exist at all.

What's reliable instead: a genuine search result adds an extra result row (avatar, title, subtitle, timestamp) above that always-present suggestion, making the search dropdown measurably taller than it is for a non-match. The check never opens a chat and never reads any text — it just diffs a screenshot taken immediately before and after typing the uid, and measures the height of the changed region.

That dropdown turns out to be a fixed pixel size that does not scale with the WeCom window (confirmed by resizing a real window and remeasuring — a match stayed exactly the same pixel height at two different window sizes), so neither a raw pixel threshold nor a window-height fraction is portable across different window sizes or different computers' display scaling. Instead, each run measures its own live baseline once (searching a uid that is guaranteed not to exist) and compares every real check against that baseline as a ratio, so it self-calibrates to whatever window size or DPI the computer running it happens to have, rather than assuming a fixed number will transfer from the machine this was developed on.

```powershell
.\.venv\Scripts\python.exe paste_sender.py --sheet "Geo TTh" --start-row 2 --end-row 10 --mode paste-only --action check-group-chat
```

The frontend has the same action as the **Check group chat status** button next to **Paste comments**, which checks the student rows you have selected in the current class.

### Checking several classes at once

The **Bulk group chat check** panel in the right-hand rail works on whole class rosters instead of selected rows. It lists every class in the semester you are currently viewing with its student count, all ticked by default; untick any you want to skip and press **Check group chats**. Unticked classes are remembered while you move between class tabs.

This runs one class at a time under the hood and writes every class's results back together at the end. Before starting anything it checks once that WeCom or WhatsApp actually has an open window and stops immediately with a clear message if neither does — otherwise every student would separately retry a launch that cannot succeed, which is slow and makes WhatsApp open a browser tab per student. For that same reason the batch itself never auto-opens an app: have the ones you need already open. Budget roughly three seconds per student.

## Options

- `--workbook`: Path to the Excel file. Defaults to `./Geo_TTh_Student_Script_fixed_rows_only.xlsx`.
- `--sheet`: Sheet name. Defaults to `Geo TTh`.
- `--row`: Excel row number. Row `2` is the first student row.
- `--all`: Generate feedback for every student row in the sheet.
- `--start-row`: First row for `--all`. Defaults to `2`.
- `--end-row`: Last row for `--all`. Omit to continue through the sheet.
- `--review-csv`: Save generated preview rows to a UTF-8 CSV with row, UID, student, status, revised remark, and feedback columns.
- `--class-review`: What the class covered today. This becomes the first paragraph.
- `--class-review-file`: Optional text file containing the class review.
- `--feedback-type`: `comprehensive`, `general`, or `quiz`. Defaults to `comprehensive`.
- `--use-api`: Use OpenAI to polish the student-specific parent comment.
- `--revise-remark`: Use OpenAI to revise `Remark for Student` first.
- `--write-revised-remark`: Save the revised remark back to the sheet.
- `--model`: OpenAI model name. Defaults to `gpt-5.5`.
- `--write`: Save the generated text back to the workbook.

Absent students are skipped and no feedback comment is generated for them.
Fields marked `Not Observed` are left out of the message.
`Additional Comment` is appended to the end of paragraph 2.
