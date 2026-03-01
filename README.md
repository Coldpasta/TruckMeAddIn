# Excel Visual Diff – Git-style comparison right inside Excel

**Visual diff tool for Excel spreadsheets** – see exactly what changed between two files, with colorful highlights, navigation, and detailed summary – all without leaving Excel.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![GitHub stars](https://img.shields.io/github/stars/Coldpasta/TruckMeAddIn?style=social)](https://github.com/Coldpasta/TruckMeAddIn)
[![Product Hunt](https://img.shields.io/badge/Launch%20on-Product%20Hunt-orange?logo=producthunt)](https://www.producthunt.com/) <!-- update link after launch -->

## Screenshots

![Demo – before and after diff](assets/screenshot-before-after.png)  
*(Zmiany podświetlone: zielony = dodane, czerwony = usunięte, żółty = zmodyfikowane)*

![Task pane with change list](assets/screenshot-taskpane-list.png)  
*(Lista wszystkich zmian z możliwością skoku do komórki)*

## Features

- **Beautiful visual diff** directly in your worksheet – cells are highlighted in familiar git-style colors
- **Compares values AND formulas** (detects both value changes and formula edits)
- **Handles row/column insertions and deletions** using smart matching (LCS-based)
- **Navigation** – Previous / Next buttons + clickable list of changes
- **Clear highlights** with one click
- **Trial mode** – 5 free comparisons to try before buying
- **Two pricing plans**:
  - **Basic** – $11.99/month – unlimited comparisons + core diff features
  - **Pro** – $19.99/month – export diff to new sheet/PDF, format changes detection (coming soon), priority support

## How to Install (Sideload – for testing)

1. Download the latest release or clone this repository
2. Open Excel → **Insert** → **My Add-ins** → **Developer** tab → **Upload My Add-in**
3. Select the file `manifest.prod.xml` from the project
4. The "Show Diff" button should appear in the Home tab ribbon
5. Click it → task pane opens → enjoy!

**Note:** Sideload works on Windows, Mac and Excel Online. Full AppSource publication coming soon.

## Quick Demo

Watch 90-second demo:  
[Loom Video – How to use Excel Visual Diff](https://www.loom.com/share/YOUR_LOOM_LINK_HERE)  
*(nagraj i wklej link przed publikacją)*

## Pricing & Plans

- **Free Trial** → 5 comparisons – no credit card required
- **Basic** → $11.99 / month – unlimited use, core diff features
- **Pro** → $19.99 / month – advanced export, format diff (coming), priority updates

Upgrade directly from the add-in after trial ends.

## Roadmap (what’s coming next)

- Compare two uploaded files (not only open workbook vs upload)
- Detect formatting changes (bold, color, borders)
- Export diff report to new worksheet or PDF
- Support for VBA code comparison
- Git-like commit history inside Excel
- Dark mode support
- Better performance for very large files (>50k cells)

## Contributing

Pull requests are welcome!  
Especially welcome:
- Bug fixes
- Performance improvements for large sheets
- New features from roadmap

## License

MIT License – see [LICENSE](LICENSE) file.

## Built with

- Office JS API
- React + Fluent UI
- SheetJS (xlsx)
- Firebase (trial licensing)
- Netlify Functions (backend)

## Polish / Polski

**Excel Visual Diff** – dodatek do Excela, który pokazuje zmiany między arkuszami tak, jak git diff w kodzie: kolory, nawigacja, lista zmian.  
Jeśli pracujesz z dużymi modelami finansowymi, audytami, kontrolingiem – to narzędzie znacznie ułatwia życie.

Chętnie przyjmę feedback po polsku na Issues lub mailowo.

---

Made with ❤️ by [Lukas / @Coldpasta]  
Questions? → Lmacewicz@gmail.com
