# FinanceApp

This repository provides a simple tool for allocating loan collections to a submission sheet.

## Requirements

- Python 3.8+
- `pandas`
- `tkinterdnd2` (optional, for drag-and-drop support)

Install requirements with:

```bash
pip install pandas tkinterdnd2
```

## Usage

Run the GUI from the repository root:

```bash
python loan_allocator_gui.py
```

Select the *Submission* and *Collected* Excel files. After processing, choose where to save `output.xlsx`. A summary of records processed and totals is displayed.

If some payments cannot be matched to any outstanding instalment, the tool reports the unallocated balance in the summary.

