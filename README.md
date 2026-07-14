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
- `feedback_master.py`: master generator that can call general, quiz, or comprehensive feedback.
- `paste_sender.py`: CLI wrapper for supervised paste actions. It checks app status and never sends.
- `paste_comment.py`: row-specific comment payload selection.
- `paste_mass_notification.py`: shared mass-notification payload selection.
- `class_review_builder.py`: create or overwrite `class_review.txt` from teacher text, notes, or slides.
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

The app opens at `http://127.0.0.1:8765`. Each class sheet has its own linked UTF-8 announcement file in `announcements`. The frontend saves student edits into `app_data/feedback.db`, generates general or quiz feedback, and runs the supervised paste-only workflow for selected rows.

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
