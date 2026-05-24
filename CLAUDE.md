@AGENTS.md

---

# Defy Field Execute — Current State

## Project Location
`C:\Users\CarlDosSantos-(OUTER\Projects\defy-field-execute`

## Tech Stack
- Next.js (latest), React 19, TypeScript, Tailwind CSS
- Vercel deployment (auto-deploy from master)
- SharePoint integration via Microsoft Graph API (images + data)
- Perigee raw export (Excel) as input
- JSZip for PPTX generation, XLSX for Excel parsing, sharp for image processing
- GitHub: https://github.com/CarlofSaints/Defy-Field-Execute

## What The App Does
Field execution tool for Defy. Reps upload Perigee visit data (Excel exports), and the app generates various reports:
- **Beko Stand Report** — PowerPoint slides showing store stand images (2-up per slide)
- **Red Flag Report** — Issues flagged during visits
- **Training Feedback Report** — Training session feedback
- Other report types via the build menu

Reports pull images from SharePoint folders, embed them in generated PPTX files.

---

## ACTIVE BUG: Beko Stand Report — PPTX Repair Dialog (UNRESOLVED)

### Client Complaint
"I'm still having an issue with the Beko Stand Report. When using the OJ tool, it always needs to repair the file and it doesn't seem to repair fully. Images are still rotated and/or stretched, and even if I manually correct this, it doesn't seem to save these changes. The file size is also much larger now than it used to be."

### What's Been Done (3 commits, all deployed, bug PERSISTS)

#### Commit 24df62b — Image processing + aspect ratio + Content_Types
- Added `sharp` dependency for image processing
- `processImage()` — auto-rotates EXIF, resizes to max 1920px, compresses JPEG quality 80 with mozjpeg
- `fitImageEmu()` — calculates aspect-ratio-correct EMU dimensions instead of forcing 5080000x5080000 square
- `buildContentSlideXml()` — accepts `img1Dims`/`img2Dims`, centres images in their column
- Content_Types: changed check from `!includes("jpg") && !includes("jpeg")` to just `!includes("jpg")`
- ZIP compression bumped from level 6 to level 9

#### Commit eb9745e — Strip changesInfo from Content_Types
- `[Content_Types].xml` declared `ppt/changesInfos/changesInfo1.xml` but the file was already being skipped during ZIP copy
- Added regex to strip `<Override>` entries containing "changesInfo"

#### Commit 721f958 — Strip customXml, rebuild docProps and _rels
- Strip `customXml/` files (SharePoint metadata artifacts) from output ZIP
- Strip customXml references from `presentation.xml.rels` and `[Content_Types].xml`
- Remove `docProps/custom.xml` (SharePoint custom properties, orphaned without customXml/)
- Rebuild `docProps/app.xml` with correct slide count (template hardcodes 3)
- Rebuild `_rels/.rels` without custom-properties reference to removed custom.xml

### What's Been Verified
- TypeScript compiles clean
- Consistency audit: no orphan part references in Content_Types, no missing files referenced by .rels
- All customXml, changesInfo, custom.xml stripped
- docProps/app.xml slide count is dynamic
- _rels/.rels rebuilt clean (only presentation.xml, core.xml, app.xml)

### What HASN'T Been Tried Yet
- **Opening a generated PPTX in a hex editor / XML viewer** to see exactly what PowerPoint's repair log says it fixed
- **PowerPoint's repair log** — after repair, check File > Info for repair details, or look at the Event Viewer on Windows for Office repair events
- **Testing with zero images** — does the repair dialog appear even with no embedded images? This would isolate whether it's the image embedding or the base template structure
- **Comparing template vs output** — diff every XML file between the template PPTX and a generated output to find structural differences PowerPoint objects to
- **Checking slide layout/master references** — each slide XML references a slideLayout via rId1. If the layout .rels or the layout itself has issues, PowerPoint flags it
- **Checking the presentation.xml sldIdLst** — the slide ID list in presentation.xml must match the actual slides. `buildPresXml()` rebuilds this but may have ID conflicts with other IDs in the file
- **Theme/slideMaster validation** — the template's slideMasters and slideLayouts reference each other; if any chain is broken, repair triggers
- **Checking if the TEMPLATE ITSELF triggers repair** — open the raw `public/PPT Template AM.pptx` in PowerPoint and see if it also needs repair. If so, the template is the root cause.

### Key File
`lib/reports/stand-report.ts` — the entire PPTX generation pipeline

### Architecture of stand-report.ts
1. Parse Perigee Excel → extract rows with store/rep/date/image URLs
2. Group rows by store (2 images per slide)
3. Fetch images from SharePoint by matching URL filename to SP file listing
4. Process each image through sharp (rotate, resize, compress)
5. Load template PPTX (`public/PPT Template AM.pptx`)
6. Copy template files (skipping slides, changesInfos, customXml, manifests)
7. Generate cover slide (slide1.xml) + N content slides + closing slide
8. Rebuild presentation.xml, presentation.xml.rels, Content_Types, docProps/app.xml, _rels/.rels
9. Output as ZIP buffer

### Template PPTX Structure (`public/PPT Template AM.pptx`)
Contains SharePoint co-authoring artifacts that cause issues:
- `customXml/` — 3 items with SharePoint content type metadata
- `ppt/changesInfos/changesInfo1.xml` — co-authoring tracking
- `docProps/custom.xml` — SharePoint custom properties (ContentTypeId etc.)
- `_rels/.rels` references custom.xml
- `docProps/app.xml` hardcodes `<Slides>3</Slides>`
- Template has 3 slides (slide1-3), 2 slide layouts, 1 slide master

---

## Uncommitted Changes (Other Files)
These are NOT related to the stand report bug:
- `lib/reports/build-menu.ts` — 110 lines changed (previous work)
- `lib/reports/training-feedback.ts` — 414 lines changed (previous work)

---

## Key Files

### Reports
- `lib/reports/stand-report.ts` — Beko Stand Report PPTX generator (active bug)
- `lib/reports/build-menu.ts` — Report type menu/builder
- `lib/reports/training-feedback.ts` — Training feedback report
- `public/PPT Template AM.pptx` — PPTX template for stand reports

### SharePoint Integration
- `lib/graph-oj.ts` — Microsoft Graph API helpers (listFilesInSPFolder, downloadSPFileById)
- `lib/appSettings.ts` — App settings (includes image folder paths)
- `lib/spUrlParser.ts` — SharePoint URL parsing

### App
- `app/` — Next.js app directory with pages for reports, settings, etc.
