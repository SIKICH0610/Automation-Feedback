# Commands

## Setup

Install dependencies:

```powershell
python -m pip install -r requirements.txt
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

