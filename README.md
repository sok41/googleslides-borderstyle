# Slide Table Styler

**English** | [日本語](README.ja.md)

A Google Slides add-on (built with Google Apps Script) that applies clean, consistent designs to your tables in one click — borders, background colors, and text styles together.

Instead of formatting tables one by one, pick a preset and click **Apply** to make every table in your deck look consistent.

- Website: https://slide-table-styler.zkuma.com/
- [Privacy Policy](https://slide-table-styler.zkuma.com/privacy.html) · [Terms of Service](https://slide-table-styler.zkuma.com/terms.html)
- Support: [GitHub Issues](https://github.com/sok41/slide-table-styler/issues)

## Features

- Apply one of 4 design presets in one click
- Change the accent color: pick one of your slide's theme colors, or any color
- Choose where to apply: the selected table, all tables on the current slide, or all tables in the presentation
- Set horizontal alignment (left or center) and vertical alignment (top or middle), or keep the existing alignment
- Automatically right-align columns that contain only numbers
- Set a minimum row height and make columns equal width
- Change inner lines to solid, dotted, or dashed
- Set the font and the header/body font sizes
- Choose which columns are bold (e.g. `1,4`)
- Turn bold on or off for the header row (the first row)
- Available in English and Japanese (Japanese when Chrome's display language is Japanese, English otherwise)

### Presets

| Preset | Look |
|---|---|
| Clean Gray | Light gray header row. No vertical lines; rows separated by light gray lines |
| Navy Stripe | Navy header row with white text. Light blue striped rows and a thick navy bottom line |
| Minimal (Rules Only) | No vertical lines. Thick black lines at the top and bottom, a black line under the header, and light lines between rows |
| Dotted Rows | Lightly tinted header row. Dotted lines between rows, with thick lines under the header and at the bottom |

## How to use

1. Open a presentation and choose **Extensions → Slide Table Styler → Open sidebar**.
2. Click inside a table, or select the table itself (not needed when applying to a whole slide or presentation).
3. In the sidebar, choose a preset and where to apply it, and adjust alignment and text settings if needed.
4. Click **Apply**.

The first time you run it, you'll be asked to authorize the add-on with your Google account. To undo, press Ctrl+Z (⌘+Z on Mac).

### Display language

- The sidebar, messages, and errors follow Chrome's display language (Japanese if it's Japanese, English otherwise).
- The item in the Extensions menu follows your Google account's language setting, because Chrome's language isn't available when the menu is created.
- All text lives in `MESSAGES` in [src/Code.js](src/Code.js) (menu and messages) and `I18N` in [src/Sidebar.html](src/Sidebar.html) (sidebar).

### What happens when you apply

- The first row is treated as the header row.
- When alignment is set to "No change", each cell keeps its current alignment.
- With "Right-align columns that contain only numbers" on, any column whose cells (excluding the header) all contain numbers is right-aligned, including its header. Numbers with symbols or units such as `1,234` `-12.5%` `$500` `▲300` `(1,200)` are included. Empty cells are ignored.
- Tables inside groups are not included when applying to a whole slide or presentation.
- If the font field is empty, the font isn't changed. The same goes for empty font size fields.
- Empty cells and cells hidden by merging keep their text formatting.

## Development setup

The add-on is developed and tested as an editor add-on in a standalone Apps Script project.

### Requirements

- Node.js
- [clasp](https://github.com/google/clasp) (`npm install -g @google/clasp`, or use `npx @google/clasp`)

### Steps

1. Clone this repository and log in to clasp.

   ```sh
   clasp login
   ```

2. Create a standalone Apps Script project.

   ```sh
   clasp create --type standalone --title "Slide Table Styler" --rootDir src
   ```

   clasp overwrites `src/appsscript.json` with an empty manifest at this point, so restore it with `git checkout -- src/appsscript.json`.

3. Push the code.

   ```sh
   clasp push
   ```

4. In the Apps Script editor, open **Deploy → Test deployments**, choose **Editor Add-on**, and create a test with a presentation that contains tables.
5. Run the test. **Slide Table Styler** appears in that presentation's **Extensions** menu.

The Slides API (advanced service) and the OAuth scopes are already configured in [src/appsscript.json](src/appsscript.json).

## Project structure

```
src/
├── appsscript.json   Manifest (V8 runtime, Slides API v1, OAuth scopes)
├── Code.js           Menu, preset definitions, table formatting, display language
└── Sidebar.html      Sidebar UI
docs/                 Website on GitHub Pages: home, privacy policy, terms of service
assets/icon/          Icons
```

## Adding or changing presets

Edit only `PRESETS` in [src/Code.js](src/Code.js). The sidebar's list and previews are generated from it.

- Write `name` in both Japanese (`ja`) and English (`en`).
- Colors can be `'#RRGGBB'`, `'accent'` (the accent color), or `'accent/10'` (the accent color mixed into white at 10%). Set the preset's default accent color in `accent`.
- `headerText: 'auto'` picks white or dark gray, whichever is easier to read on the header background.
- `borders` are applied in order. `color: null` makes a line transparent, and `range: 'header'` limits a line to the header row. Use `dash: 'DOT'` (dotted) or `'DASH'` (dashed) to set the line style.

The sidebar preview (`resolvePreset` in [src/Sidebar.html](src/Sidebar.html)) resolves colors with the same rules. Keep both in sync when you change them.

## Development rules

- Edit code only under `src/`. Don't edit directly in the Apps Script editor — the local files are the source of truth.
- After a change, run `clasp push` and check it with the test deployment.
