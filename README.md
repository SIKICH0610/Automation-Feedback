# Automation Feedback

Generate one parent-facing feedback entry from the provided Excel tracker.

## Project Files

- `feedback_generator.py`: preview or write feedback for one row, a row range, or the whole sheet.
- `paste_sender.py`: supervised one-row desktop paste helper for WeCom / 企业微信. It checks app status and never sends.
- `class_review_builder.py`: create or overwrite `class_review.txt` from teacher text, notes, or slides.
- `workbook_setup.py`: prepare optional workbook columns such as `Additional Comment`.
- `openai_api.py`: shared OpenAI API helper.
- `class_review.txt`: editable class-level paragraph used as paragraph 1.
- `Geo_TTh_Student_Script_fixed_rows_only.xlsx`: local workbook copy used by default.
- `docs/commands.md`: quick command reference.

The Excel feedback form also lives in this project folder:

```text
Geo_TTh_Student_Script_fixed_rows_only.xlsx
```

By default, the scripts read and write this local copy.

## Setup

Install dependencies once:

```powershell
python -m pip install -r requirements.txt
```

Optional API key for GPT-assisted workflows:

```powershell
$env:OPENAI_API_KEY="your_api_key_here"
```

## Prepare the workbook

Add the optional `Additional Comment` column to every sheet:

```powershell
python workbook_setup.py
```

Use this column for extra notes that should be added to the end of paragraph 2. In API mode, the note is translated when needed.

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
python paste_sender.py --sheet "Geo TTh" --row 2 --class-review-file class_review.txt --status
```

Open the needed app if the script can find it:

```powershell
python paste_sender.py --sheet "Geo TTh" --row 2 --class-review-file class_review.txt --open-app
```

Search WeCom by UID, press Enter to open the first relevant result, focus the message box, and paste without sending:

```powershell
python paste_sender.py --sheet "Geo TTh" --row 2 --class-review-file class_review.txt --mode paste-only
```

The paste helper recognizes window titles containing `WeCom`, `企业微信`, or `WXWork` by default. It does not press Enter after pasting. If WeCom / 企业微信 is installed in a custom location, pass `--wecom-exe "C:\path\to\WXWork.exe"` or set `WECOM_EXE`. If Enter cannot open the result on a computer, try `--ui-control-result-open`, `--coordinate-result-click`, or `--manual-result-click`. Add `--require-verification` if you want the script to stop whenever it cannot verify the chat by UI text.

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
- `--use-api`: Use OpenAI to polish the student-specific parent comment.
- `--revise-remark`: Use OpenAI to revise `Remark for Student` first.
- `--write-revised-remark`: Save the revised remark back to the sheet.
- `--model`: OpenAI model name. Defaults to `gpt-5.5`.
- `--write`: Save the generated text back to the workbook.

Absent students are skipped and no feedback comment is generated for them.
Fields marked `Not Observed` are left out of the message.
`Additional Comment` is appended to the end of paragraph 2.
