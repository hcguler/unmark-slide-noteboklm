# 🧹 unmark-slide-notebooklm

A Python script that automatically detects and removes the **NotebookLM watermark logo** from every slide in a `.pptx` PowerPoint file — without touching any other content.

---

## How It Works

When you export a presentation from [Google NotebookLM](https://notebooklm.google.com/), a small **"NotebookLM" logo** is embedded in the bottom-right corner of the background image of each slide. This script:

1. Opens the `.pptx` file using `python-pptx`
2. Extracts the background image from each slide
3. Automatically detects the logo region via pixel color analysis
4. Fills the logo area with the surrounding background color (sampled intelligently)
5. Re-embeds the cleaned image back into the slide
6. Saves the result as a new `.pptx` file

> ✅ All text, shapes, fonts, animations, and other elements remain completely intact.  
> ✅ Shared background images (used across multiple slides) are only processed once.

---

## Requirements

- Python 3.8+
- [python-pptx](https://python-pptx.readthedocs.io/)
- [Pillow](https://pillow.readthedocs.io/)
- [NumPy](https://numpy.org/)

Install all dependencies at once:

```bash
pip install python-pptx Pillow numpy
```

---

## Installation

```bash
git clone https://github.com/your-username/remove-notebooklm-logo.git
cd remove-notebooklm-logo
pip install python-pptx Pillow numpy
```

---

## Usage

### Basic — output auto-named

```bash
python remove_notebooklm_logo.py presentation.pptx
```
Output: `presentation_cleaned.pptx` (same folder)

---

### Specify output path

```bash
python remove_notebooklm_logo.py presentation.pptx output/clean_deck.pptx
```

---

### Verbose mode — see per-slide detail

```bash
python remove_notebooklm_logo.py presentation.pptx --verbose
```

Sample output:
```
INFO  Input:  presentation.pptx
INFO  Output: presentation_cleaned.pptx
INFO  Processing slides...

INFO    Slide 1: Logo removed ✓
INFO    Slide 2: Logo removed ✓
...

──────────────────────────────────────────────────
  DONE
  Slides processed : 11
  Logos removed    : 11
  Unchanged slides : 0

  ✓  Saved to: presentation_cleaned.pptx
──────────────────────────────────────────────────
```

---

### Dry run — detect without saving

```bash
python remove_notebooklm_logo.py presentation.pptx --dry-run
```

Useful to check whether logos are detected before committing to a save.

---

### Adjust detection sensitivity

If the logo is not detected (e.g. different background color), increase `--tolerance`:

```bash
python remove_notebooklm_logo.py presentation.pptx --tolerance 50
```

Default is `35`. Range: `10` (strict) → `80` (aggressive).

---

### Increase erase padding

Add extra pixels around the erased area to ensure no logo remnants:

```bash
python remove_notebooklm_logo.py presentation.pptx --padding 12
```

Default is `6`.

---

## All Options

| Flag | Default | Description |
|------|---------|-------------|
| `INPUT` | *(required)* | Path to the `.pptx` file to clean |
| `OUTPUT` | `<input>_cleaned.pptx` | Path to save the cleaned file |
| `-v`, `--verbose` | off | Show debug info per slide |
| `--dry-run` | off | Detect logos without saving |
| `--tolerance N` | `35` | Color deviation threshold for logo detection |
| `--padding N` | `6` | Extra pixels to erase around logo bounds |

---

## Examples

```bash
# Most common usage
python remove_notebooklm_logo.py my_deck.pptx

# With custom output path
python remove_notebooklm_logo.py slides/deck.pptx output/deck_final.pptx

# Check detection first
python remove_notebooklm_logo.py deck.pptx --dry-run

# Verbose output for debugging
python remove_notebooklm_logo.py deck.pptx --verbose

# Increase tolerance for difficult cases
python remove_notebooklm_logo.py deck.pptx --tolerance 55 --padding 10
```

---

## Troubleshooting

**"No logos were found"**
- Run with `--verbose` to see what the script detects
- Try `--tolerance 50` or higher
- Confirm the file was actually exported from NotebookLM (logo is in the bottom-right of background images)

**Slight color mismatch after removal**
- Try `--tolerance 40` for a slightly larger detection area
- The script samples the surrounding background color automatically; results are best on solid/uniform backgrounds

**The script crashes on my file**
- Make sure you're on Python 3.8+ (`python --version`)
- Reinstall dependencies: `pip install --upgrade python-pptx Pillow numpy`
- Open an issue and attach your `.pptx` if the problem persists

---

## Limitations

- Only removes logos that are **embedded in the slide background image** (the standard NotebookLM export format)
- Does not handle logos placed as separate shape layers on top of slides (uncommon)
- Works best when the logo is on a solid or near-uniform background color
- Not tested on password-protected `.pptx` files

---

## Project Structure

```
remove-notebooklm-logo/
├── remove_notebooklm_logo.py   # Main script
└── README.md                   # This file
```

---

## License

MIT — free to use, modify, and distribute.

---

## Contributing

Pull requests are welcome! If you encounter a `.pptx` where detection fails, please open an issue and describe the slide layout (no need to share confidential content).
