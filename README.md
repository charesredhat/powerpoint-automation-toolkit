# Personal Windows Automation

This repository contains small Windows automation scripts for working with PowerPoint presentations.

The main workflow is:

1. Apply a source presentation's template to a target presentation.
2. Extract text from the target presentation into a text file.
3. Run spelling correction against slide text using Microsoft Word.
4. Generate and refresh the FDA Unified HPC deck and supporting slide images.

## Repository Contents

- `apply_theme_and_extract.ps1`
  Applies a source presentation as a template to a target presentation, creates a backup, and extracts slide text.
- `apply_corrections_and_log.ps1`
  Runs spelling correction on a PowerPoint file and logs the changes.
- `apply_corrections_and_log_fixed.ps1`
  Same general workflow as the correction script above, with slightly cleaner log formatting.
- `apply_corrections_and_log_run.ps1`
  A lighter variation of the correction script.
- `pptx_theme_apply.py`
  Experimental Python-based alternative for background/font application and text extraction.
- `Presentation-Docs/generate_unified_hpc_ppt.ps1`
  Rebuilds the FDA Unified HPC PowerPoint deck from the shared template and slide text definitions.
- `Presentation-Docs/create_slide_images.ps1`
  Regenerates the supporting PNG slide images used by the Unified HPC deck.
- `info/codebase_summary_2026-04-01.md`
  Short codebase summary and review notes.

## Requirements

- Windows
- Microsoft PowerPoint installed
- Microsoft Word installed
- PowerShell

Optional for the Python script:

- Python 3
- `python-pptx`

## PowerShell Usage

### 1. Apply Theme And Extract Text

```powershell
.\apply_theme_and_extract.ps1 `
  -SourcePresentation "C:\Slides\template.pptx" `
  -TargetPresentation "C:\Slides\target.pptx"
```

Optional parameters:

- `-BackupPath`
- `-OutputPresentation`
- `-ExtractedTextPath`

If you omit them, the script creates default paths based on the target file.

Example:

```powershell
.\apply_theme_and_extract.ps1 `
  -SourcePresentation "C:\Slides\template.pptx" `
  -TargetPresentation "C:\Slides\target.pptx" `
  -OutputPresentation "C:\Slides\target_theme_applied.pptx" `
  -ExtractedTextPath "C:\Slides\extracted_text.txt"
```

### 2. Apply Spelling Corrections

```powershell
.\apply_corrections_and_log_fixed.ps1 `
  -TargetPresentation "C:\Slides\target_theme_applied.pptx"
```

Optional parameters:

- `-OutputPresentation`
- `-LogPath`
- `-ShowPowerPoint`

Example:

```powershell
.\apply_corrections_and_log_fixed.ps1 `
  -TargetPresentation "C:\Slides\target_theme_applied.pptx" `
  -OutputPresentation "C:\Slides\target_corrected.pptx" `
  -LogPath "C:\Slides\corrections_log.txt" `
  -ShowPowerPoint $false
```

### 3. Build the Unified HPC Deck

```powershell
.\Presentation-Docs\generate_unified_hpc_ppt.ps1
```

By default, this uses:

- Template: `.\FDA_PP_Final_Use_This - 16x9 version.pptx`
- Output: `.\Presentation-Docs\Unified_HPC_Deck_Output.pptx`

You can point the output to any destination filename you want:

```powershell
.\Presentation-Docs\generate_unified_hpc_ppt.ps1 `
  -OutputPath "C:\Path\To\Output\Unified_HPC_Deck_Output.pptx"
```

### 4. Regenerate Unified HPC Slide Images

```powershell
.\Presentation-Docs\create_slide_images.ps1
```

By default, the script writes PNG files to:

- `.\Presentation-Docs\Slide_Images`

## Outputs

Depending on the script and options used, the automation will create:

- a backup `.pptx`
- a themed output `.pptx`
- an extracted text `.txt` file
- a corrected output `.pptx`
- a corrections log `.txt`
- a rebuilt Unified HPC `.pptx`
- refreshed Unified HPC slide image `.png` files

## Notes

- The PowerShell scripts are parameterized and do not require hard-coded local paths.
- The Python script is still an experimental helper and currently uses hard-coded paths. It should be treated as a prototype until it is refactored.
- PowerPoint theme application is best-effort. Some formatting details may not transfer perfectly.
- The spelling workflow automatically uses Word's first suggestion for detected spelling issues.
- The Unified HPC scripts are meant to be run from the repository root or with full paths, and they assume the template file and `Presentation-Docs` folder are present in the repository layout shown above.

## Typical Flow

```powershell
.\apply_theme_and_extract.ps1 `
  -SourcePresentation "C:\Slides\template.pptx" `
  -TargetPresentation "C:\Slides\target.pptx"

.\apply_corrections_and_log_fixed.ps1 `
  -TargetPresentation "C:\Slides\target_theme_applied.pptx"
```

For the Unified HPC deck refresh:

```powershell
.\Presentation-Docs\create_slide_images.ps1
.\Presentation-Docs\generate_unified_hpc_ppt.ps1 `
  -OutputPath "C:\Path\To\Output\output_powerpoint.pptx"
```
