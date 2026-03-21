# PPT Preview Design

**Date:** 2026-03-21
**Status:** Approved

## Overview

Add in-app PPT/PPTX preview to the existing file preview panel. Currently PPT files show a "no preview" placeholder prompting the user to open in an external app. This spec describes a hybrid strategy: use LibreOffice CLI when available for high-quality rendering, otherwise fall back to a pptx2json-based slide renderer.

## Goals

- Show PPT content inside the app without requiring any mandatory external dependency
- Use LibreOffice for pixel-accurate PDF rendering when it is installed
- Provide a usable text + image fallback when LibreOffice is absent
- Integrate naturally into the existing `OfficeDocViewer` / `PDFViewer` infrastructure

## Non-Goals

- Bundling LibreOffice into the app distribution
- Animated transitions or slide effects
- Editable PPT content
- Pixel-perfect layout reproduction in the fallback renderer
- Exhaustive temp-file cleanup in V1 (see Temp File Strategy below)

---

## Architecture

### Conversion Flow

```
Open PPT file
    |
    v
[Main Process] detect LibreOffice installation
    |
    +-- found --> soffice --headless --convert-to pdf <file> --outdir <tmpdir>
    |               --> return { to: 'ppt-pdf', result: { success: true, data: '/tmp/xxx.pdf' } }
    |
    +-- not found / conversion fails --> pptx2json.toJson(filePath)
                        --> extract slides: title + body texts + embedded images (Base64)
                        --> return { to: 'ppt-json', result: { success: true, data: PPTJsonData } }
```

The renderer calls `ppt-pdf` first. Any failure (including `LIBRE_OFFICE_NOT_FOUND` and conversion errors) causes it to fall back to `ppt-json`. The fallback hint badge (`preview.ppt.fallbackHint`) is shown whenever the pptx2json path is used, regardless of the specific failure reason.

---

## Backend Changes (Main Process)

### 1. `src/process/utils/libreofficeUtils.ts` (new)

Exports one function:

```ts
export async function findLibreOfficeBin(): Promise<string | null>;
```

Checks the following paths in order and returns the first executable found, or `null`:

| Platform | Paths checked                                                                                                |
| -------- | ------------------------------------------------------------------------------------------------------------ |
| macOS    | `/Applications/LibreOffice.app/Contents/MacOS/soffice`, `/usr/local/bin/soffice`                             |
| Windows  | `C:\Program Files\LibreOffice\program\soffice.exe`, `C:\Program Files (x86)\LibreOffice\program\soffice.exe` |
| Linux    | `/usr/bin/soffice`, `/usr/local/bin/soffice`, `/snap/bin/libreoffice`                                        |

Result is cached in memory for the process lifetime (detection runs at most once).

> **Directory note:** `src/process/utils/` currently exceeds the 10-child limit. Place `libreofficeUtils.ts` in a new sub-directory `src/process/utils/conversion/`, and also move `previewUtils.ts` there. Update all import paths accordingly.

### 2. `conversionService.pptToPdf(filePath)` (new method)

- Calls `findLibreOfficeBin()`. If `null`, throws `Error('LIBRE_OFFICE_NOT_FOUND')`.
- Runs via `safeExecFile` (the project wrapper at `src/process/utils/safeExec.ts`, which uses `spawn + detached: true` to avoid `SIGTTOU` in the Electron main process) with a 30-second timeout:
  ```ts
  await safeExecFile(soffice, ['--headless', '--convert-to', 'pdf', '--outdir', tmpDir, filePath], { timeout: 30_000 });
  ```
- Output PDF path is `path.join(tmpDir, path.basename(filePath, ext) + '.pdf')`.
- Uses a **content-hash-based fixed filename** (`<sha256(filePath+mtime)>.pdf`) so that converting the same file twice reuses the cached PDF instead of accumulating new files.
- Returns `ConversionResult<string>` where `data` is the absolute path to the generated PDF.

### 3. Temp File Strategy

- PDF files are written to `os.tmpdir()` under a fixed hash-based filename (same input → same output path, no accumulation).
- On `app.on('before-quit')`, the main process deletes all `*.pdf` files it created during the session (tracked in an in-memory set).
- This is a known V1 limitation: files are not cleaned up if the app crashes.

### 4. PDF File Access from Renderer

The generated PDF path is an absolute `file://` path on the user's machine. The `PDFViewer` component already loads PDFs via `file://` URLs. Confirm that the app's `BrowserWindow.webPreferences` does not set `webSecurity: false` — existing PDFViewer behaviour (which already serves local files) is the precedent; no additional work needed here unless a new CSP restriction is introduced.

### 5. `conversionService.pptToJson(filePath)` (fix existing)

Current issues:

- Contains debug `console.log` statements — remove all.
- Slide key detection uses nested path lookup that never matches the flat key map returned by pptx2json.

Fix slide extraction:

```ts
const slideKeys = Object.keys(json)
  .filter((k) => /^ppt\/slides\/slide\d+\.xml$/.test(k))
  .sort((a, b) => {
    const numA = parseInt(a.match(/slide(\d+)/)?.[1] ?? '0', 10);
    const numB = parseInt(b.match(/slide(\d+)/)?.[1] ?? '0', 10);
    return numA - numB;
  });
```

Add image extraction per slide:

- Load `ppt/slides/_rels/slideN.xml.rels` from the same flat json map.
- Resolve `r:embed` references to `ppt/media/*` entries (which are Buffers).
- Convert to `data:<mime>;base64,<...>` strings.
- Attach as `images: string[]` in the slide data.

### 6. `PPTSlideData` type update (`src/common/types/conversion.ts`)

Replace the opaque `content: any` with structured fields:

```ts
export interface PPTSlideData {
  slideNumber: number;
  title: string;
  texts: string[]; // body text paragraphs (excluding title)
  images: string[]; // data URLs of embedded images
}
```

### 7. `DocumentConversionTarget` and response union update

```ts
// Add 'ppt-pdf' to the union:
export type DocumentConversionTarget = 'markdown' | 'excel-json' | 'ppt-json' | 'ppt-pdf';

// Add to DocumentConversionResponse union:
// (data is the absolute file path of the generated PDF)
| { to: 'ppt-pdf'; result: ConversionResult<string> }
```

### 8. `documentBridge` (`src/process/bridge/documentBridge.ts`)

Add new case `'ppt-pdf'` to the switch:

```ts
case 'ppt-pdf': {
  if (!ensureExtension(filePath, PPT_EXTENSIONS)) {
    return unsupportedResult(to, 'Only PowerPoint files can be converted to PDF');
  }
  const result = await conversionService.pptToPdf(filePath);
  return { to, result };
}
```

---

## Frontend Changes (Renderer)

### 1. `PPTViewer.tsx` (new)

**Path:** `src/renderer/pages/conversation/Preview/components/viewers/PPTViewer.tsx`

**Props:**

```ts
interface PPTViewerProps {
  filePath?: string;
  hideToolbar?: boolean;
}
```

**State machine:**

```
idle
  -> loading (try ppt-pdf first)
     -> pdfReady (pdfPath: string)  --> render <PDFViewer filePath={pdfPath} hideToolbar />
     -> fallbackLoading (try ppt-json on any ppt-pdf failure)
        -> slidesReady (slides: PPTSlideData[], currentIndex: number)
        -> error
```

**Fallback slide card layout (16:9 aspect ratio):**

- Header bar: slide number badge (`3 / 10`), open-in-system Button (Arco `Button` component, not raw `<button>`)
- Content area:
  - Title: `text-18px font-semibold text-t-primary mb-12px`
  - Text paragraphs: `text-14px text-t-secondary leading-relaxed`
  - Images: `max-h-200px object-contain` grid (max 2 per row)
  - If no content: `text-13px text-t-tertiary` placeholder (`preview.ppt.noContent`)
- Navigation footer: Arco `Button` components for Prev / Next, keyboard `ArrowLeft` / `ArrowRight`

**Fallback hint badge:** When showing pptx2json-rendered slides (i.e., LibreOffice was unavailable or failed), show `preview.ppt.fallbackHint` in the toolbar in `text-t-tertiary` style.

**LibreOffice indicator:** When showing the PDF path, toolbar shows `preview.ppt.viaLibreOffice` badge in `text-t-tertiary`.

### 2. `OfficeDocViewer.tsx` (update)

Replace the `if (docType === 'ppt')` block that renders the placeholder with:

```tsx
if (docType === 'ppt') {
  return <PPTViewer filePath={filePath} hideToolbar={hideToolbar} />;
}
```

### 3. Export `PPTViewer` from viewers index

In `src/renderer/pages/conversation/Preview/components/viewers/index.ts`, add:

```ts
export { default as PPTViewer } from './PPTViewer';
```

> **Directory note:** `viewers/` currently has 10 children. Adding `PPTViewer.tsx` brings it to 11, exceeding the project's 10-child limit. Group the office-document viewers into a sub-directory: `src/renderer/pages/conversation/Preview/components/viewers/office/` containing `OfficeDocViewer.tsx`, `ExcelViewer.tsx`, and the new `PPTViewer.tsx`. Update all import paths and re-export from `viewers/index.ts`.

### 4. i18n keys

Add to `preview.json` in **all six locale directories** (`en-US`, `zh-CN`, `zh-TW`, `ja-JP`, `ko-KR`, `tr-TR`):

| Key                  | en-US                                 | zh-CN                            | zh-TW                            | ja-JP                                          | ko-KR                            | tr-TR                                   |
| -------------------- | ------------------------------------- | -------------------------------- | -------------------------------- | ---------------------------------------------- | -------------------------------- | --------------------------------------- |
| `ppt.loading`        | Loading presentation...               | 正在加载演示文稿...              | 正在載入簡報...                  | プレゼンテーションを読み込み中...              | 프레젠테이션 로딩 중...          | Sunum yükleniyor...                     |
| `ppt.converting`     | Converting with LibreOffice...        | 正在通过 LibreOffice 转换...     | 正在透過 LibreOffice 轉換...     | LibreOffice で変換中...                        | LibreOffice로 변환 중...         | LibreOffice ile dönüştürülüyor...       |
| `ppt.fallbackHint`   | Basic preview (LibreOffice not found) | 基础预览（未检测到 LibreOffice） | 基本預覽（未偵測到 LibreOffice） | 基本プレビュー（LibreOffice が見つかりません） | 기본 미리보기 (LibreOffice 없음) | Temel önizleme (LibreOffice bulunamadı) |
| `ppt.slideCount`     | `{{current}} / {{total}}`             | `{{current}} / {{total}}`        | `{{current}} / {{total}}`        | `{{current}} / {{total}}`                      | `{{current}} / {{total}}`        | `{{current}} / {{total}}`               |
| `ppt.prevSlide`      | Previous                              | 上一张                           | 上一張                           | 前へ                                           | 이전                             | Önceki                                  |
| `ppt.nextSlide`      | Next                                  | 下一张                           | 下一張                           | 次へ                                           | 다음                             | Sonraki                                 |
| `ppt.noContent`      | This slide has no text content        | 此幻灯片无文字内容               | 此投影片無文字內容               | このスライドにテキストはありません             | 이 슬라이드에 텍스트가 없습니다  | Bu slaytda metin içeriği yok            |
| `ppt.viaLibreOffice` | via LibreOffice                       | via LibreOffice                  | via LibreOffice                  | via LibreOffice                                | via LibreOffice                  | via LibreOffice                         |

---

## Error Handling

| Scenario                                         | Behavior                                                      |
| ------------------------------------------------ | ------------------------------------------------------------- |
| LibreOffice not found                            | Silently fall back to ppt-json; show `ppt.fallbackHint` badge |
| LibreOffice found but conversion fails/times out | Silently fall back to ppt-json; show `ppt.fallbackHint` badge |
| ppt-json parsing fails                           | Show error state + Arco Button "Open in system app"           |
| Slide has no title or body text                  | Show `ppt.noContent` placeholder in slide card                |
| filePath is undefined                            | Show `preview.errors.missingFilePath`                         |

---

## Testing

- Unit test `findLibreOfficeBin` with mocked `fs.access` (all paths absent → null, first path present → returns it)
- Unit test `pptToJson` slide extraction with a minimal synthetic flat-key json map
- Unit test `pptToPdf` error path (`LIBRE_OFFICE_NOT_FOUND`) when detection returns null
- Component test `PPTViewer`: mock IPC to return ppt-pdf success → verify PDFViewer rendered; mock ppt-pdf failure + ppt-json success → verify slide cards rendered with `ppt.fallbackHint` badge
- Component test: keyboard navigation (`ArrowLeft`/`ArrowRight`) changes current slide index

---

## File Change Summary

| File                                                                             | Change                                                                                |
| -------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------- |
| `src/process/utils/conversion/libreofficeUtils.ts`                               | New (new sub-directory)                                                               |
| `src/process/utils/conversion/previewUtils.ts`                                   | Moved from `src/process/utils/previewUtils.ts`                                        |
| `src/process/services/conversionService.ts`                                      | Add `pptToPdf`, fix `pptToJson`, track created temp files                             |
| `src/process/bridge/documentBridge.ts`                                           | Add `ppt-pdf` case, add before-quit cleanup handler                                   |
| `src/common/types/conversion.ts`                                                 | Update `PPTSlideData`, `DocumentConversionTarget`, `DocumentConversionResponse` union |
| `src/renderer/pages/conversation/Preview/components/viewers/PPTViewer.tsx`       | New                                                                                   |
| `src/renderer/pages/conversation/Preview/components/viewers/index.ts`            | Export PPTViewer                                                                      |
| `src/renderer/pages/conversation/Preview/components/viewers/OfficeDocViewer.tsx` | Route `docType === 'ppt'` to PPTViewer                                                |
| `src/renderer/services/i18n/locales/en-US/preview.json`                          | Add `ppt.*` keys                                                                      |
| `src/renderer/services/i18n/locales/zh-CN/preview.json`                          | Add `ppt.*` keys                                                                      |
| `src/renderer/services/i18n/locales/zh-TW/preview.json`                          | Add `ppt.*` keys                                                                      |
| `src/renderer/services/i18n/locales/ja-JP/preview.json`                          | Add `ppt.*` keys                                                                      |
| `src/renderer/services/i18n/locales/ko-KR/preview.json`                          | Add `ppt.*` keys                                                                      |
| `src/renderer/services/i18n/locales/tr-TR/preview.json`                          | Add `ppt.*` keys                                                                      |
