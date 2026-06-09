# Project Notes

## Goal

Build a feedback automation workflow for class rosters:

- teachers update the Excel tracker
- teachers edit or regenerate `class_review.txt`
- the script previews one student, a row range, or the whole class
- the script can write final parent-facing comments back to the workbook

## Current Workflow

1. Edit `class_review.txt` for the class-level paragraph.
2. Add student observations in `Geo_TTh_Student_Script_fixed_rows_only.xlsx`.
3. Preview a small range:

   ```powershell
   python feedback_generator.py --sheet "Geo TTh" --all --start-row 2 --end-row 5 --class-review-file class_review.txt
   ```

4. Write the whole class when the preview looks right:

   ```powershell
   python feedback_generator.py --sheet "Geo TTh" --all --class-review-file class_review.txt --write
   ```

## API Roadmap

- Use GPT to summarize uploaded slides into `class_review.txt`.
- Use GPT to revise Chinese teacher notes in `Remark for Student`.
- Use GPT to translate and blend `Additional Comment` into paragraph 2.
- Batch generation can now export a review CSV before writing to Excel.


