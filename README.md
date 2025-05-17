# Ø76 mm Disc Drawing Generator

A Windows-only Python script that drives **AutoCAD®** invisibly and produces a fully-annotated drawing:

* Ø 76 mm circle with dashed centre-lines
* 3 mm-tall rectangular bar (labeled **0.15 ± 0.05 mm**)
* Stacked limits **+0.0 / –0.2** above the diameter value
* Closed-filled magenta arrows, white text
* Note block ("VISUAL CRITERIA …") at lower left

Outputs are written to **`outputs/`**:

* `outputs/disc_76.dwg` – drawing
* `outputs/disc_76.log` – run-time log

---

## Prerequisites

| Requirement                     | Notes                                             |
| ------------------------------- | ------------------------------------------------- |
| **AutoCAD desktop ≥ 2016**      | Any full edition (LT/OEM may lack COM automation) |
| **Python 3.8 – 3.12 (Windows)** | Bitness must match AutoCAD (32/64-bit)            |
| **pywin32**                     | Install with `pip install -r requirements.txt`    |

---

## Installation

```powershell
git clone https://github.com/your-org/disc-76-dwg.git
cd disc-76-dwg

python -m venv venv
venv\Scripts\activate

pip install -r requirements.txt
```

---

## Usage

```powershell
python generate_disc_drawing.py
```

The script will:

1. Launch AutoCAD in the background
2. Build the drawing
3. Save files in **outputs/**
4. Log progress to **outputs/disc\_76.log**
5. Close AutoCAD cleanly

---

## Configuration

Edit the constants near the top of **`generate_disc_drawing.py`** to fine-tune the output.

| Variable     | Description                                          | Default |
| ------------ | ---------------------------------------------------- | ------- |
| `CIRCLE_DIA` | Disc diameter (mm)                                   | `76.0`  |
| `BAR_HEIGHT` | Visible bar height (mm)                              | `3.0`   |
| `LABEL_THK`  | Thickness shown in the dim text (mm)                 | `0.15`  |
| `TXT_H`      | Dimension text height (mm)                           | `2.5`   |
| `TOL_SCALE`  | Tolerance text height factor                         | `0.9`   |
| `DX_DIA`     | Horizontal nudge for diameter tolerance (× `TXT_H`)  | `-4`    |
| `DX_THK`     | Horizontal nudge for thickness tolerance (× `TXT_H`) | `-7.5`  |

---

## License

MIT — see `LICENSE`.
