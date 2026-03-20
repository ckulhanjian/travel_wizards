# Travel Wizards — Desktop Tools

Two standalone desktop utilities for invoice processing. Each compiles to a single `.exe` (Windows) via GitHub Actions.

---

## Tools

| Folder | Tool | Description |
|---|---|---|
| `invoice_processor/` | **Invoice Processor** | Renames & stamps PDF invoices with overlay + back page |
| `hotel_invoice_processor/` | **Hotel Invoice Processor** | Generates hotel invoice `.xlsx` files from the Hotels report |

---

## Builds

Executables are produced automatically on every push to `main` by GitHub Actions (`.github/workflows/build.yml`).

Download the latest builds from the [**Releases**](../../releases) page or the **Actions → Artifacts** tab.

---

## Repository Layout

```
travel-wizards-tools/
├── .github/
│   └── workflows/
│       └── build.yml               ← single workflow builds both tools
├── invoice_processor/
│   ├── invoice_processor.py        ← main script
│   ├── overlay.pdf                 ← bundled asset (add to repo)
│   ├── backside.pdf                ← bundled asset (add to repo)
│   ├── invoice_processor.spec      ← PyInstaller spec
│   └── requirements.txt
├── hotel_invoice_processor/
│   ├── hotel_invoice_processor.py  ← main script
│   ├── Hotel_Invoice-Template.xlsx ← bundled asset (add to repo)
│   ├── hotel_invoice_processor.spec
│   └── requirements.txt
└── README.md
```

---

## Adding a New Template (Hotel Invoice)

1. Place the new `.xlsx` template file in `hotel_invoice_processor/`
2. Add it to the `TEMPLATES` dict in `hotel_invoice_processor.py`
3. Add the filename to the `datas` list in `hotel_invoice_processor.spec`
4. Commit — the build runs automatically

---

## Local Development

```bash
# Invoice Processor
cd invoice_processor
pip install -r requirements.txt
python invoice_processor.py

# Hotel Invoice Processor
cd hotel_invoice_processor
pip install -r requirements.txt
python hotel_invoice_processor.py
```

---

## Building Locally

```bash
pip install pyinstaller

# Invoice Processor
cd invoice_processor
pyinstaller invoice_processor.spec

# Hotel Invoice Processor
cd hotel_invoice_processor
pyinstaller hotel_invoice_processor.spec
```

Output lands in `dist/` inside each tool folder.
