# Powerpoint Automation Toolkit

This repository contains small Windows automation scripts for working with PowerPoint presentations.

The main workflow is:

1. Apply a source presentation's template to a target presentation.
2. Extract text from the target presentation into a text file.
3. Run spelling correction against slide text using Microsoft Word.
4. Generate and refresh the FDA powerpoint deck and supporting slide images.

## Repository Contents

- `presentation-processing/apply_template_and_extract_text.ps1`
  Applies a source presentation as a template to a target presentation, creates a backup, and extracts slide text.
- `presentation-processing/correct_spelling_and_log.ps1`
  Runs spelling correction on a PowerPoint file and logs the changes.
- `presentation-processing/correct_spelling_and_log_improved.ps1`
  Same general workflow as the correction script above, with slightly cleaner log formatting.
- `presentation-processing/correct_spelling_and_log_basic.ps1`
  A lighter variation of the correction script.
- `presentation-processing/pptx_theme_apply.py`
  Experimental Python-based alternative for background/font application and text extraction.
- `presentation-generation/build_presentation_deck.ps1`
  Rebuilds the FDA powerpoint PowerPoint deck from the shared template and slide text definitions.
- `presentation-generation/export_slide_images.ps1`
  Regenerates the supporting PNG slide images used by the powerpoint deck.
- `presentation-generation/edit_jpg_and_update_pptx.py`
  Applies JSON-driven edits to JPG/PNG files and inserts the edited image into a PowerPoint deck.
- `presentation-generation/sample_image_edits.json`
  Minimal example edit specification for the Python image update helper.
- `documentation/codebase_summary_2026-04-01.md`
  Short codebase summary and review notes.
- `requirements.txt`
  Python dependencies for the image and PowerPoint helper.

## Requirements

- Windows
- Microsoft PowerPoint installed
- Microsoft Word installed
- PowerShell

Optional for the Python script:

- Python 3
- `python-pptx`
- `Pillow`

Install the Python dependencies with:

```powershell
python -m pip install -r requirements.txt
```

## PowerShell Usage

### 1. Apply Theme And Extract Text

```powershell
.\presentation-processing\apply_template_and_extract_text.ps1 `
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
.\presentation-processing\apply_template_and_extract_text.ps1 `
  -SourcePresentation "C:\Slides\template.pptx" `
  -TargetPresentation "C:\Slides\target.pptx" `
  -OutputPresentation "C:\Slides\target_theme_applied.pptx" `
  -ExtractedTextPath "C:\Slides\extracted_text.txt"
```

### 2. Apply Spelling Corrections

```powershell
.\presentation-processing\correct_spelling_and_log_improved.ps1 `
  -TargetPresentation "C:\Slides\target_theme_applied.pptx"
```

Optional parameters:

- `-OutputPresentation`
- `-LogPath`
- `-ShowPowerPoint`

Example:

```powershell
.\presentation-processing\correct_spelling_and_log_improved.ps1 `
  -TargetPresentation "C:\Slides\target_theme_applied.pptx" `
  -OutputPresentation "C:\Slides\target_corrected.pptx" `
  -LogPath "C:\Slides\corrections_log.txt" `
  -ShowPowerPoint $false
```

### 3. Build the powerpoint Deck

```powershell
.\presentation-generation\build_presentation_deck.ps1
```

By default, this uses:

- Template: `.\FDA_PP_Final_Use_This - 16x9 version.pptx`
- Output: `.\presentation-generation\Presentation_Deck_Output.pptx`

You can point the output to any destination filename you want:

```powershell
.\presentation-generation\build_presentation_deck.ps1 `
  -OutputPath "C:\Path\To\Output\output_powerpoint.pptx"
```

### 4. Regenerate powerpoint Slide Images

```powershell
.\presentation-generation\export_slide_images.ps1
```

By default, the script writes PNG files to:

- `.\presentation-generation\Slide_Images`

### 5. Edit a JPG/PNG and Insert It Into a PowerPoint Deck

Use `edit_jpg_and_update_pptx.py` after the existing generation scripts when you need to apply repeatable image edits and place the corrected image into a presentation. The script has two independent parts:

- Image editing: reads a source JPG/PNG, applies a JSON edit specification, and writes a corrected image.
- PowerPoint updating: places a JPG/PNG as a full-slide image in an existing `.pptx`.

You can run either part by itself, or run both parts in one command.

Install dependencies first:

```powershell
python -m pip install -r requirements.txt
```

Example edit-only run:

```powershell
python .\presentation-generation\edit_jpg_and_update_pptx.py `
  --source-image "C:\Slides\Slide3-marked-up.jpg" `
  --edit-spec ".\presentation-generation\sample_image_edits.json" `
  --output-image "C:\Slides\Slide3-clean.jpg"
```

Example insert an already edited image into a deck without changing the image:

```powershell
python .\presentation-generation\edit_jpg_and_update_pptx.py `
  --source-image "C:\Slides\Slide3-clean.jpg" `
  --presentation ".\presentation-generation\Presentation_Deck_Output.pptx" `
  --slide-number 3 `
  --insert-mode replace `
  --picture-fit contain `
  --output-presentation "C:\Slides\Presentation_Deck_Output_with_slide3.pptx"
```

Example replace slide 3 with the edited image:

```powershell
python .\presentation-generation\edit_jpg_and_update_pptx.py `
  --source-image "C:\Slides\Slide3-marked-up.jpg" `
  --edit-spec ".\presentation-generation\sample_image_edits.json" `
  --output-image "C:\Slides\Slide3-clean.jpg" `
  --presentation ".\presentation-generation\Presentation_Deck_Output.pptx" `
  --slide-number 3 `
  --insert-mode replace `
  --picture-fit contain `
  --output-presentation "C:\Slides\Presentation_Deck_Output_with_slide3.pptx"
```

Common arguments:

- `--source-image`: source JPG/PNG. Required for image editing and also used as the insertion image when no `--output-image` is provided.
- `--edit-spec`: JSON file describing the image edits to apply.
- `--output-image`: corrected image path. Required when using `--edit-spec`.
- `--presentation`: existing PowerPoint file to update.
- `--output-presentation`: destination PowerPoint file. Required when using `--presentation`.
- `--slide-number`: 1-based slide number for `replace` and `insert-after`.
- `--insert-mode`: `replace`, `insert-after`, or `append`.
- `--picture-fit`: `contain`, `cover`, or `stretch`.

Supported edit operations in the JSON file:

- `cover`: fill a rectangular area, useful for removing old text.
- `rectangle`: draw an outlined or filled rectangle.
- `text`: add text with optional bold, color, wrapping, and anchor.
- `paste`: place another image into a box.
- `crop`: crop the source image.
- `resize`: resize the working image.

Minimal edit specification:

```json
{
  "quality": 95,
  "operations": [
    {
      "type": "cover",
      "box": [0, 0, 1152, 52],
      "fill": "white"
    },
    {
      "type": "text",
      "text": "FDA HPC RESOURCES TODAY",
      "position": [576, 29],
      "font_size": 32,
      "bold": true,
      "fill": "#141c2d",
      "anchor": "mm"
    }
  ]
}
```

The `box` format is `[x, y, width, height]` in pixels. Text `position` is `[x, y]`; `anchor: "mm"` centers the text on that point. Relative image paths in `paste` operations are resolved relative to the JSON file.

PowerPoint insertion modes:

- `replace`: clear an existing slide and place the image on it.
- `insert-after`: add a new image slide after the requested slide number.
- `append`: add a new image slide to the end of the deck.

Picture fit options:

- `contain`: preserve aspect ratio and fit the entire image inside the slide.
- `cover`: preserve aspect ratio and fill the slide, cropping overflow.
- `stretch`: resize to the exact slide size.

## Outputs

Depending on the script and options used, the automation will create:

- a backup `.pptx`
- a themed output `.pptx`
- an extracted text `.txt` file
- a corrected output `.pptx`
- a corrections log `.txt`
- a rebuilt powerpoint `.pptx`
- refreshed powerpoint slide image `.png` files
- an edited slide image `.jpg` or `.png`
- an updated presentation containing the edited image slide

## Notes

- The PowerShell scripts are parameterized and do not require hard-coded local paths.
- The Python image/PPT helper is parameterized and should be run from the repository root or with full paths.
- PowerPoint theme application is best-effort. Some formatting details may not transfer perfectly.
- The spelling workflow automatically uses Word's first suggestion for detected spelling issues.
- The generation scripts are meant to be run from the repository root or with full paths, and they assume the template file and `presentation-generation` folder are present in the repository layout shown above.

## Typical Flow

```powershell
.\presentation-processing\apply_template_and_extract_text.ps1 `
  -SourcePresentation "C:\Slides\template.pptx" `
  -TargetPresentation "C:\Slides\target.pptx"

.\presentation-processing\correct_spelling_and_log_improved.ps1 `
  -TargetPresentation "C:\Slides\target_theme_applied.pptx"
```

For the powerpoint deck refresh:

```powershell
.\presentation-generation\export_slide_images.ps1
.\presentation-generation\build_presentation_deck.ps1 `
  -OutputPath "C:\Path\To\Output\output_powerpoint.pptx"

python .\presentation-generation\edit_jpg_and_update_pptx.py `
  --source-image "C:\Slides\Slide3-marked-up.jpg" `
  --edit-spec ".\presentation-generation\sample_image_edits.json" `
  --output-image "C:\Slides\Slide3-clean.jpg" `
  --presentation "C:\Path\To\Output\output_powerpoint.pptx" `
  --slide-number 3 `
  --insert-mode replace `
  --output-presentation "C:\Path\To\Output\output_powerpoint_with_edits.pptx"
```
