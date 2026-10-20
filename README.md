# PUBLICATION_HELPERS

Small Python utilities that help with publishing and writing scientific papers.

## MDPI to author-year citations

`mdpi_to_elsevier_citation.py` converts numbered citations such as `[1]`, `[2, 3]`, or `[4–7]` in a Word manuscript into author-year citations based on the manuscript's reference list.

Example:

```bash
python mdpi_to_elsevier_citation.py PAPER_mdpi.docx -o PAPER_apa7.docx
```

The script supports common reference formats including personal authors, institutional authors, standards, webpages, and APA/MDPI-style entries.

### Requirements

```bash
pip install python-docx
```

> Use a different output filename so the original manuscript is preserved.

# DSL_TOOLS — YOLO Dataset Splitter & Trainer

A simple Python GUI for combining multiple YOLO datasets, splitting them into **train, validation, and test sets without copying files**, and training Ultralytics YOLO models.

## Features

- **Multi-folder selection** — combine images and labels from multiple directories.
- **Adjustable dataset splits** — linked sliders and text inputs for Train / Val / Test percentages.
- **Ignore empty images** — optionally exclude images with missing or empty annotation files.
- **No file duplication** — generate `train.txt`, `val.txt`, `test.txt`, and `autosplit_data.yaml`.
- **YOLO training** — supports YOLO 8, 11, and 26, including fine-tuning existing `.pt` models.
- **Training configuration** — epochs or hours, batch size, image size, CPU/GPU, and output paths.
- **Compact CLI output** — cleaner epoch-level training progress.

## Installation

Requires Python with Tkinter support.

```bash
pip install ultralytics
```

Missing Python dependencies can also be installed automatically by the script.

## Usage

Launch the GUI:

```bash
python DSL_TOOLS.py
```

1. Select one or more dataset folders.
2. Configure train/validation/test percentages.
3. Enable or disable **IGNORE EMPTY IMAGES**.
4. Choose your YOLO model and training parameters.
5. Start splitting and training directly from the GUI.

Both mixed image/annotation folders and separate `images/` and `labels/` directories are supported.

**Note:** Original images and labels remain untouched. All datasets must use compatible class IDs.

## Additional Modes

```bash
# Split datasets without training
python DSL_TOOLS.py --autosplit

# Original YoloSplitter workflow
python DSL_TOOLS.py --legacy
```

## Dependencies

- Python
- Ultralytics YOLO
- PyTorch
- Tkinter
- PyYAML

---

Built for efficient YOLO dataset preparation and model training without unnecessary file duplication.
