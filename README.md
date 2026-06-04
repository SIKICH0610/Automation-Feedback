# Automation Feedback Prototype

Generate one parent-facing feedback entry from the provided Excel tracker.

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

## Write one entry back to Excel

```powershell
python feedback_generator.py --sheet "Geo TTh" --row 2 --class-review-file class_review.txt --write
```

This writes the generated text into the `Feedback` column for that row.

## API-assisted student comments

Revise the Chinese teacher note in `Remark for Student`, preview the parent comment, and leave the sheet unchanged:

```powershell
python feedback_generator.py --sheet "Geo TTh" --row 2 --class-review-file class_review.txt --revise-remark --use-api
```

Revise the Chinese teacher note, save it back to `Remark for Student`, generate the parent comment, and write it to `Feedback`:

```powershell
python feedback_generator.py --sheet "Geo TTh" --row 2 --class-review-file class_review.txt --revise-remark --write-revised-remark --use-api --write
```

## Options

- `--workbook`: Path to the Excel file. Defaults to `~/Downloads/Geo_TTh_Student_Script_fixed_rows_only.xlsx`.
- `--sheet`: Sheet name. Defaults to `Geo TTh`.
- `--row`: Excel row number. Row `2` is the first student row.
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
