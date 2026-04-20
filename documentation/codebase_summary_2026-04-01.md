# HFP Codebase Summary

Date stamped: 2026-04-01

This repository is a small PowerPoint automation workspace focused on reformatting and cleaning an existing presentation. It is not a general application with reusable modules; it is a task-oriented set of scripts and presentation assets built around a specific file-processing workflow.

## What It Does

The codebase automates three main steps for PowerPoint files:

1. Apply styling or theme elements from a source presentation to a target presentation.
2. Extract slide text from the target presentation into a text file for review.
3. Run spelling correction on slide text and log the changes that were applied.

## Main Workflow

The scripts are centered around these files:

- `CDER Research Governance WOAP presentation.pptx`
  Source presentation used as the styling/theme reference.
- `HPC_Cloud_Analysis_HPC_GAB_Consolidated_Slides_2.pptx`
  Main target presentation being modified.
- `HPC_Cloud_Analysis_HPC_GAB_Consolidated_Slides_2_backup.pptx`
  Backup copy created before edits.
- `HPC_Cloud_Analysis_HPC_GAB_Consolidated_Slides_2_theme_applied.pptx`
  Output after theme/styling is applied.
- `HPC_extracted_text.txt`
  Extracted text from the target presentation for inspection.

## Script Roles

### `presentation-processing/apply_template_and_extract_text.ps1`

Uses PowerPoint COM automation to:

- open the source and target presentations
- create a backup of the target file
- attempt to apply the source presentation as a template
- extract text from each slide shape into `HPC_extracted_text.txt`
- save a themed output copy

This is the most direct "theme and extract" automation script in the repo.

### `presentation-processing/pptx_theme_apply.py`

Uses the Python `python-pptx` library as an alternative approach to:

- create a backup of the target file
- inspect the source deck for a sample image and font information
- extract text from the target deck
- apply a best-effort background image and font styling to the target deck
- save a themed output copy

This is a looser approximation of theming than the PowerPoint COM approach. It appears to be an experiment or fallback path rather than the primary production script.

### `presentation-processing/correct_spelling_and_log.ps1`

Uses PowerPoint and Word COM automation to:

- open the themed presentation
- inspect text in each shape
- send text through Word spelling checks
- replace misspelled words with the first suggested correction
- save a corrected copy
- write a correction log to `HPC_corrections_log.txt`

### `presentation-processing/correct_spelling_and_log_improved.ps1`

This is effectively a cleaned-up version of `presentation-processing/correct_spelling_and_log.ps1`. The behavior is nearly the same, but some log formatting is safer and more explicit.

### `presentation-processing/correct_spelling_and_log_basic.ps1`

Another variation of the correction script. It is very similar to the fixed version, with slightly reduced console/UI setup.

## Overall Structure

This codebase is best understood as a one-off document-processing toolkit with:

- PowerShell as the main automation layer
- COM automation for Microsoft PowerPoint and Word
- one Python prototype for manipulating `.pptx` files without Office COM
- hard-coded absolute file paths tied to this local workspace

The current folder layout separates the repo by function:

- `presentation-processing/` for theme application, text extraction, spelling correction, and the Python prototype
- `presentation-generation/` for deck-building and slide-image export scripts
- `documentation/` for repository notes and summaries

## Important Characteristics

- The scripts are tightly coupled to specific filenames and directories.
- They assume Microsoft Office is installed and COM automation is available.
- The correction logic is best-effort and automatically accepts Word's first spelling suggestion.
- The code is procedural and task-specific rather than modular.
- The repo now groups scripts and notes into dedicated folders, but the automation remains task-specific rather than modular.

## Review Notes

- `presentation-processing/correct_spelling_and_log.ps1`, `presentation-processing/correct_spelling_and_log_improved.ps1`, and `presentation-processing/correct_spelling_and_log_basic.ps1` contain substantial duplication and could be consolidated.
- Hard-coded absolute paths make the scripts less portable.
- Error handling is present but minimal; failures are usually logged or printed rather than fully recovered from.
- The Python theming approach does not truly transfer a PowerPoint theme; it approximates theme changes by copying a background image and font settings.

## Bottom Line

This repository is a specialized PowerPoint cleanup workflow for a specific presentation set. Its purpose is to restyle a target deck, extract slide text, and apply automated spelling corrections while preserving backups and logs.
