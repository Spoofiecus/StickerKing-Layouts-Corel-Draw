# Sticker King - Audit Feedback & Reality Check

Branch: `cline/6n34nwxf`. This is a read-only audit of the **actual GitHub state**, re-derived because the prior session's work is not present on GitHub (see §0).

## 0. IMPORTANT correction: the archived work IS on a PRIVATE repo I cannot access

**Correction (supersedes the original §0):** the previous session's work was pushed to a **private** repo:
`Spoofiecus/Sticker-King---Scripts-extentions-and-apps` (branch `cline/sq349hc2`). It is NOT lost.

This sandbox authenticates as a GitHub App (`cline-cloud[bot]`) whose installation covers **only**
`Spoofiecus/StickerKing-Layouts-Corel-Draw` (`gh api /installation/repositories` returns exactly that one repo).
The other Sticker King repos are visible only because they are **public**. The private
`Sticker-King---Scripts-extentions-and-apps` repo therefore returns **404** here (`gh`, `git ls-remote`).

=> To continue that work, the app needs access to that repo (grant the Cline GitHub App access to it, or make it public).
Everything below is still a valid audit of the PUBLIC repos; it does not replace the private repo's contents.

## 1. What actually exists (9 sticker repos, no monorepo)

Pricing Calculator Chrome extension exists as **6 near-identical copies**:
`Stickerking-chrome-extention` (v1, monolithic) -> `Experimental-layout-generator`, `Corel_Layouts-quotes` (both embed the same monolithic copy) -> `ChromeExtentionV2` (modular `js/`) -> `StickerkingChrome-Extentionv3` (+jspdf Save-PDF) -> `StickerkingExtention-Finalizations` (**canonical**, newest).
Manifest name/version identical in all six: `Sticker King Pricing Calculator` / `1.2`.

CorelDRAW generators = **4 independent implementations** (all plain-text VBA despite `.gms`/`.gsm`):
- `StickerKing-Layouts-Corel-Draw` (`Corel Layouts.gms` == `LayoutGeneratorV2`; `CorelLayoutsOrigional`)
- `LayoutGenerator` (`AdvancedStickerLayout.gsm` + `StickerLayoutGeneratorOrigional.gsm`; `top-left-cutting` branch is byte-identical to `main`)
- `ZoopZoops` (`LayoutGenerator.gms` - a cutline-grid builder, different tool)
- `Corel_Layouts-quotes` (`Corel_Layout_generator.gsm`, `Price Check.xls` binary)

## 2. Verified findings

### 2.1 Pricing Calculator

**Reference numbers intended:** 650mm roll, **R460/m2** (R420 is the old price), 1mm bleed, 15% VAT, R100 min order, R0.20 min sticker.

| Item | Old copies (r1 / Experimental / Corel) | Canonical (Finalizations) |
|---|---|---|
| price/m2 fallback | **420** (`sidebar.js:82`) WRONG | 460 (`js/config.js:1`) OK |
| stickers/row | `>=0.95 -> ceil` (`sidebar.js:98,108`) WRONG | `Math.floor` (`js/calculator.js:14,23`) OK |
| orientation tie | tie -> vertical (`sidebar.js:116`) | tie -> horizontal (`calculator.js:36`) OK |
| row area | hard-coded `0.65*H` (`sidebar.js:100,110`) | `ROLL_WIDTH/1000` OK |
| Infinity guard | partial | present OK |

Reproduced the user's exact bugs from the OLD code:
- **70x105 -> 6/row @ R3.49** (both orientations price identically; `price < P_vertical` is false -> picks vertical/landscape).
- **130x50 -> 5/row @ R2.99** (`650/131 = 4.96`, `.96 >= .95` -> ceil to 5; but `5*130 + 4*1mm = 654mm > 650mm`, physically impossible).

Canonical (Finalizations) already computes the correct, physically-valid answers: **70x105 -> 9/row @ R3.49**, **130x50 -> 12/row @ R3.24** (portrait 130 across -> 4/row). So the fix is: **delete the 5 stale copies and keep/polish Finalizations.**

Remaining real defects in ALL copies:
- `sidebar.html:22` loads `src="logo.png"` but the file is `Logo.png` -> **404 on case-sensitive systems** (P1)
- `manifest.json` uses invalid key **`default_icons`** (should be `default_icon`) -> toolbar icon override ignored (P2)
- dead bundled assets (Font Awesome `css/all.min.css`, `webfonts/*`, `*-icon.png`, `remove.png`) - page uses a **CDN** instead, breaking offline/CSP (P2)
- stray **`webfonts/11.pdf`** (200KB) shipped inside 5 copies (P2)
- finalizations **dropped** the Save-as-PDF feature that v3 had (P2 regression)

### 2.2 CorelDRAW generators

- `Corel Layouts.gms` == `LayoutGeneratorV2` (whitespace-only diff) - same file twice in one repo.
- `CorelLayoutsOrigional` == `LayoutGenerator/StickerLayoutGeneratorOrigional.gsm` - identical across repos.
- `Corel_Layout_generator.gsm` is an early snapshot of `AdvancedStickerLayout.gsm` (which is newer/canonical).
- Duplicate `Const ROLL_WIDTH_MM = 650` declared twice per advanced file.
- ZoopZoops: `Const BATCH_HEIGHT = 300` vs advanced `BATCH_MAX_HEIGHT = 350` (conflicting); its `INSTRUCTIONS.md` promises an InputBox for batch height that the code never shows; error handler can call `EndCommandGroup` without a matching `BeginCommandGroup`.
- No script contains `.Delete`/`.Remove` - none can delete artwork (contradicting a claim in the archived session).
- Advanced scripts create hundreds of shapes with **no** `BeginCommandGroup` -> not undoable as a unit.

### 2.3 Branding / structure
- Folder **`Sticker King  Calculator`** (double space) duplicated in V2, Experimental, Corel_Layouts-quotes.
- READMEs are single-line repo names; "Sticker King" branding is **inconsistent** inside tools.
- Claimed typos ("Genertors", "Lable", "Label Sheet Creator", "Canva -> CorelDRAW Converter") **do not exist** on GitHub.

## 3. Recommended fix plan

**P0**
1. Consolidate the 6 calculator copies into ONE canonical repo (keep `StickerkingExtention-Finalizations`); delete/archive the rest.
2. Re-add the physically-validated fit rule + deterministic tie-break as the single source of truth in that copy.
3. Pick ONE canonical Corel generator: `LayoutGeneratorV2` (simple) + `AdvancedStickerLayout.gsm` (pricing/batch); archive the duplicates.

**P1**
4. Fix `logo.png` -> `Logo.png`, `default_icons` -> `default_icon`.
5. Re-implement / restore Save-as-PDF if it is a wanted feature.

**P2**
6. Remove dead bundled assets + stray `webfonts/11.pdf`; de-duplicate `Sticker King  Calculator` folder.
7. Real READMEs per repo; unify "Sticker King" branding.

## 4. Products from the archive that do not exist and would need to be built from scratch
- Canva PDF Fixer / Canva->CorelDRAW converter
- Order Form Creator Chrome extension
- The pricing test CSV + a shared business-config file
