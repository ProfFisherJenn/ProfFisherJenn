# Word to Canvas QTI Converter - Usage Guide

## Setup (Same as Your PPTX Converter!)

Your QTI converter uses the **exact same folder structure** as your doc→pptx converter:

```
Your Folder/
├── convert_docs_to_pptx_auto_template_with_log_v2.py  ← Your existing script
├── convert_docs_to_qti.py                              ← New QTI converter
├── input_docs/                                         ← Shared input folder
├── output_pptx/                                        ← PPTX outputs
├── output_qti/                                         ← QTI outputs
├── conversion_log.txt                                  ← PPTX log
└── qti_conversion_log.txt                              ← QTI log
```

## Workflow

### For Lesson Content (PPTX):
1. Drop lesson Word docs into `input_docs/`
2. Run `convert_docs_to_pptx_auto_template_with_log_v2.py`
3. Get presentations in `output_pptx/`
4. Source files automatically deleted from `input_docs/`

### For Quiz Content (QTI):
1. Drop quiz Word docs into `input_docs/`
2. Run `convert_docs_to_qti.py`
3. Get QTI zips in `output_qti/`
4. Source files automatically deleted from `input_docs/`

**Same folder, same workflow - just different output formats!**

## Running the QTI Converter

**Windows:** Double-click `convert_docs_to_qti.py`

**Mac/Linux:** 
```bash
python3 convert_docs_to_qti.py
```

The script will:
- ✓ Process all .docx files in `input_docs/`
- ✓ Create `filename_qti.zip` in `output_qti/`
- ✓ Delete source files from `input_docs/`
- ✓ Log everything to `qti_conversion_log.txt`
- ✓ Open the output folder when done

## Requirements (One-Time Setup)

If you haven't already installed these for your PPTX converter:
```bash
pip install python-docx lxml
```

## Your Converted Quiz Results

✓ All 8 questions successfully parsed from M3K1_SelectionTools.docx
✓ All question types supported (Multiple Choice, True/False, Multiple Answer, Matching)
✓ All correct answers properly marked
✓ Ready to import into Canvas

## How to Import QTI Files into Canvas

1. Go to your Canvas course → **Settings** → **Import Course Content**
2. Choose **"QTI .zip file"** as the import type
3. Upload the `_qti.zip` file from `output_qti/`
4. Click **Import**
5. Questions appear in **Quizzes** (Classic Quizzes)

## Document Format Requirements

Your Word documents should be formatted like your sample (M3K1_SelectionTools.docx):

### Required Elements:
- **Section headers:** "Multiple Choice", "True/False", "Multiple Answer", "Matching"
- **Questions:** Each on its own line (can end with ? or .)
- **Answer choices:** Each on separate lines
- **Correct answers:** Either:
  - Bold formatting on the correct choice(s) in Word, OR
  - "Correct answer:" line after each question

### Example Format:
```
Multiple Choice

What is the primary use of Photoshop?
Raster image editing
Vector graphics
Video editing
3D modeling

Correct answer: Raster image editing

True/False

True or False: CMYK is used for web design.
True
False

Correct answer: False
```

## Troubleshooting

**Problem:** No files processed
- Check that .docx files are in `input_docs/` folder
- Make sure files aren't open in Word

**Problem:** Some questions not importing
- Verify section headers match exactly (Multiple Choice, True/False, etc.)
- Check that correct answers are bolded OR have "Correct answer:" lines

**Problem:** Import fails in Canvas
- Make sure you selected "QTI .zip file" as import type (not "Canvas quiz")
- Try re-running the converter

**Problem:** Script won't run
- Install requirements: `pip install python-docx lxml`
- For Windows: Make sure Python is installed and in PATH

## About Classic vs. New Quizzes

This imports to **Classic Quizzes** (not New Quizzes). This is the most reliable method for QTI imports in Canvas. Classic Quizzes are fully functional, stable, and preferred by many instructors.

## Benefits of This Setup

✓ **Same workflow** as your PPTX converter
✓ **Shared input folder** - one place to drop all docs
✓ **Batch processing** - handle multiple quizzes at once
✓ **Auto cleanup** - source files removed after conversion
✓ **Detailed logging** - track what happened
✓ **Auto-naming** - output files named after source files
