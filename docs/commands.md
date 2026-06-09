# Commands

## Setup

Install dependencies:

```powershell
.\setup.ps1
```

Optional API key:

```powershell
$env:OPENAI_API_KEY="your_api_key_here"
```

## Class Review

Edit `class_review.txt` directly, or overwrite it from text:

```powershell
python class_review_builder.py --source-text "Today we discussed triangle similarity..." --output class_review.txt
```

## Feedback Preview

One student:

```powershell
python feedback_generator.py --sheet "Geo TTh" --row 2 --class-review-file class_review.txt
```

Small range:

```powershell
python feedback_generator.py --sheet "Geo TTh" --all --start-row 2 --end-row 5 --class-review-file class_review.txt
```

Small range with review CSV:

```powershell
python feedback_generator.py --sheet "Geo TTh" --all --start-row 2 --end-row 5 --class-review-file class_review.txt --review-csv review_preview.csv
```

All students:

```powershell
python feedback_generator.py --sheet "Geo TTh" --all --class-review-file class_review.txt
```

## Feedback Write

Write one row:

```powershell
python feedback_generator.py --sheet "Geo TTh" --row 2 --class-review-file class_review.txt --write
```

Write all rows:

```powershell
python feedback_generator.py --sheet "Geo TTh" --all --class-review-file class_review.txt --write
```

## Supervised Paste

Check readiness:

```powershell
python paste_sender.py --sheet "Geo TTh" --row 2 --class-review-file class_review.txt --status
```

Open the needed app:

```powershell
python paste_sender.py --sheet "Geo TTh" --row 2 --class-review-file class_review.txt --open-app
```

Paste one WeCom / 企业微信 row without sending. The script searches by UID, opens the best matching result, focuses the message box, and pastes:

```powershell
python paste_sender.py --sheet "Geo TTh" --row 2 --class-review-file class_review.txt --mode paste-only
```

Paste a mixed WeCom / WhatsApp batch without sending. Missing contacts are marked `needs_review` and the batch continues:

```powershell
python paste_sender.py --sheet "Geo TTh" --start-row 2 --end-row 8 --class-review-file class_review.txt --mode paste-only
```

The supervised paste path uses keyboard shortcuts and UI Automation controls. It does not use calculated screen-position clicks, so resizing the WeCom or WhatsApp window should not change where the robot tries to act.

Try a WhatsApp Web row, such as Aditya on row 12:

```powershell
python paste_sender.py --sheet "Geo TTh" --row 12 --class-review-file class_review.txt --mode paste-only
```

WhatsApp group search uses the same pattern as WeCom: `Ctrl+F`, paste the UID/search key, press `Enter`, then paste only if the search box is no longer active.

Paste selected rows:

```powershell
python paste_sender.py --sheet "Geo TTh" --rows 3,5 --class-review-file class_review.txt --mode paste-only
```

Debug safe WeCom search candidates:

```powershell
python paste_sender.py --sheet "Geo TTh" --row 2 --class-review-file class_review.txt --debug-search-results
```

