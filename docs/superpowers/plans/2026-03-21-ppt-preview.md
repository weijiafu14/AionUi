# PPT Preview Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Enable in-app PPT/PPTX preview — use LibreOffice CLI when available for PDF rendering, otherwise fall back to a pptx2json-based slide renderer with text + images.

**Architecture:** The renderer calls `document.convert({ to: 'ppt-pdf' })` first; if that fails the main process returns an error and the renderer retries with `{ to: 'ppt-json' }`. The new `PPTViewer` component handles both paths: delegates to the existing `PDFViewer` on success, or renders slide cards from `PPTSlideData[]` on fallback. `OfficeDocViewer` routes `docType === 'ppt'` to `PPTViewer` instead of the current placeholder.

**Tech Stack:** Electron (main process), React + TypeScript (renderer), pptx2json, safeExecFile, Vitest, @arco-design/web-react, UnoCSS, i18next

---

## File Map

| File | Action | Responsibility |
|------|--------|---------------|
| `src/common/types/conversion.ts` | Modify | Structured `PPTSlideData`, add `'ppt-pdf'` target and response type |
| `src/process/utils/conversion/` | Create dir | New sub-directory to stay under 10-child limit |
| `src/process/utils/conversion/libreofficeUtils.ts` | Create | Detect LibreOffice binary; cached |
| `src/process/utils/conversion/previewUtils.ts` | Move | From `src/process/utils/previewUtils.ts`; update 3 importers |
| `src/process/services/conversionService.ts` | Modify | Fix `pptToJson`, add `pptToPdf`, track temp PDFs for cleanup |
| `src/process/bridge/documentBridge.ts` | Modify | Add `ppt-pdf` IPC case; register before-quit cleanup |
| `src/renderer/pages/conversation/Preview/components/viewers/office/` | Create dir | Office viewer sub-directory |
| `src/renderer/pages/conversation/Preview/components/viewers/office/PPTViewer.tsx` | Create | Two-path PPT viewer component |
| `src/renderer/pages/conversation/Preview/components/viewers/office/ExcelViewer.tsx` | Move | From `viewers/ExcelViewer.tsx` |
| `src/renderer/pages/conversation/Preview/components/viewers/office/OfficeDocViewer.tsx` | Move | From `viewers/OfficeDocViewer.tsx`; route ppt to PPTViewer |
| `src/renderer/pages/conversation/Preview/components/viewers/index.ts` | Modify | Re-export from new office/ paths + add PPTViewer |
| `src/renderer/services/i18n/locales/*/preview.json` | Modify | Add `ppt` nested object (6 locales) |
| `tests/unit/process/utils/conversion/libreofficeUtils.test.ts` | Create | Unit tests for binary detection |
| `tests/unit/process/services/conversionService.ppt.test.ts` | Create | Unit tests for pptToJson and pptToPdf |
| `tests/unit/renderer/PPTViewer.dom.test.tsx` | Create | Component tests for PPTViewer |

---

## Task 1: Update Shared Types

**Files:**
- Modify: `src/common/types/conversion.ts`

- [ ] **Step 1: Replace `PPTSlideData.content: any` with structured fields**

In `src/common/types/conversion.ts`, replace:
```ts
export interface PPTSlideData {
  slideNumber: number;
  content: any;
}
```
With:
```ts
export interface PPTSlideData {
  slideNumber: number;
  title: string;
  texts: string[];   // body text paragraphs, excluding title
  images: string[];  // data URLs of embedded images
}
```

- [ ] **Step 2: Add `'ppt-pdf'` to `DocumentConversionTarget`**

```ts
export type DocumentConversionTarget = 'markdown' | 'excel-json' | 'ppt-json' | 'ppt-pdf';
```

- [ ] **Step 3: Add `ppt-pdf` to `DocumentConversionResponse` union**

In the `DocumentConversionResponse` type, add:
```ts
| { to: 'ppt-pdf'; result: ConversionResult<string> }
// ConversionResult<string>.data is the absolute path to the generated PDF
```

- [ ] **Step 4: Verify TypeScript compiles**

```bash
bunx tsc --noEmit
```
Expected: no errors related to `PPTSlideData`, `DocumentConversionTarget`, or `DocumentConversionResponse`.

- [ ] **Step 5: Commit**

```bash
git add src/common/types/conversion.ts
git commit -m "feat(ppt-preview): update conversion types for structured PPT data and ppt-pdf target"
```

---

## Task 2: Create `conversion/` Sub-directory and Move `previewUtils.ts`

**Files:**
- Create: `src/process/utils/conversion/` (directory)
- Move: `src/process/utils/previewUtils.ts` → `src/process/utils/conversion/previewUtils.ts`
- Modify (3 files): `src/process/task/CodexAgentManager.ts`, `GeminiAgentManager.ts`, `AcpAgentManager.ts`

- [ ] **Step 1: Create the directory and copy previewUtils.ts**

```bash
mkdir -p src/process/utils/conversion
cp src/process/utils/previewUtils.ts src/process/utils/conversion/previewUtils.ts
```

- [ ] **Step 2: Update the 3 importers**

In each of these files, change:
```ts
import { handlePreviewOpenEvent } from '@process/utils/previewUtils';
```
To:
```ts
import { handlePreviewOpenEvent } from '@process/utils/conversion/previewUtils';
```

Files to update:
- `src/process/task/CodexAgentManager.ts:32`
- `src/process/task/GeminiAgentManager.ts:25`
- `src/process/task/AcpAgentManager.ts:22`

- [ ] **Step 3: Delete the original file**

```bash
rm src/process/utils/previewUtils.ts
```

- [ ] **Step 4: Verify TypeScript compiles**

```bash
bunx tsc --noEmit
```
Expected: no import errors.

- [ ] **Step 5: Commit**

```bash
git add src/process/utils/conversion/previewUtils.ts src/process/utils/previewUtils.ts \
  src/process/task/CodexAgentManager.ts src/process/task/GeminiAgentManager.ts src/process/task/AcpAgentManager.ts
git commit -m "refactor(process): move previewUtils into utils/conversion/ subdirectory"
```

---

## Task 3: LibreOffice Detection Utility

**Files:**
- Create: `src/process/utils/conversion/libreofficeUtils.ts`
- Create: `tests/unit/process/utils/conversion/libreofficeUtils.test.ts`

- [ ] **Step 1: Write the failing test**

Create `tests/unit/process/utils/conversion/libreofficeUtils.test.ts`:
```ts
import { describe, it, expect, vi, beforeEach } from 'vitest';
import fs from 'fs/promises';

vi.mock('fs/promises');
vi.mock('@process/utils/conversion/libreofficeUtils', async (importOriginal) => {
  const mod = await importOriginal<typeof import('@process/utils/conversion/libreofficeUtils')>();
  return mod;
});

import { findLibreOfficeBin, _resetLibreOfficeBinCache } from '@process/utils/conversion/libreofficeUtils';

describe('findLibreOfficeBin', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    _resetLibreOfficeBinCache();
  });

  it('returns null when no LibreOffice binary is found', async () => {
    vi.mocked(fs.access).mockRejectedValue(new Error('not found'));
    const result = await findLibreOfficeBin();
    expect(result).toBeNull();
  });

  it('returns the first found binary path on macOS', async () => {
    const macPath = '/Applications/LibreOffice.app/Contents/MacOS/soffice';
    vi.mocked(fs.access).mockImplementation(async (p) => {
      if (p === macPath) return;
      throw new Error('not found');
    });
    vi.spyOn(process, 'platform', 'get').mockReturnValue('darwin');
    const result = await findLibreOfficeBin();
    expect(result).toBe(macPath);
  });
});
```

- [ ] **Step 2: Run the test — expect FAIL**

```bash
bun run test tests/unit/process/utils/conversion/libreofficeUtils.test.ts
```
Expected: FAIL — module not found.

- [ ] **Step 3: Create `libreofficeUtils.ts`**

Create `src/process/utils/conversion/libreofficeUtils.ts`:
```ts
import fs from 'fs/promises';

const LIBRE_OFFICE_PATHS: Partial<Record<NodeJS.Platform, string[]>> = {
  darwin: [
    '/Applications/LibreOffice.app/Contents/MacOS/soffice',
    '/usr/local/bin/soffice',
  ],
  win32: [
    'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
    'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
  ],
  linux: [
    '/usr/bin/soffice',
    '/usr/local/bin/soffice',
    '/snap/bin/libreoffice',
  ],
};

let cachedBin: string | null | undefined = undefined;

/**
 * Find the LibreOffice soffice binary on the current platform.
 * Result is cached for the process lifetime.
 */
export async function findLibreOfficeBin(): Promise<string | null> {
  if (cachedBin !== undefined) return cachedBin;

  const paths = LIBRE_OFFICE_PATHS[process.platform] ?? [];
  for (const p of paths) {
    try {
      await fs.access(p);
      cachedBin = p;
      return cachedBin;
    } catch {
      // not found at this path, try next
    }
  }
  cachedBin = null;
  return null;
}

/** Reset cache — for testing only. */
export function _resetLibreOfficeBinCache(): void {
  cachedBin = undefined;
}
```

- [ ] **Step 4: Run the test — expect PASS**

```bash
bun run test tests/unit/process/utils/conversion/libreofficeUtils.test.ts
```
Expected: PASS.

- [ ] **Step 5: Verify TypeScript**

```bash
bunx tsc --noEmit
```

- [ ] **Step 6: Commit**

```bash
git add src/process/utils/conversion/libreofficeUtils.ts \
  tests/unit/process/utils/conversion/libreofficeUtils.test.ts
git commit -m "feat(ppt-preview): add LibreOffice binary detection utility"
```

---

## Task 4: Fix `pptToJson` and Add `pptToPdf` to ConversionService

**Files:**
- Modify: `src/process/services/conversionService.ts`
- Create: `tests/unit/process/services/conversionService.ppt.test.ts`

- [ ] **Step 1: Write the failing tests**

Create `tests/unit/process/services/conversionService.ppt.test.ts`:
```ts
import { describe, it, expect, vi, beforeEach } from 'vitest';
import path from 'path';

vi.mock('@process/utils/conversion/libreofficeUtils', () => ({
  findLibreOfficeBin: vi.fn(),
}));
vi.mock('@process/utils/safeExec', () => ({
  safeExecFile: vi.fn(),
}));
vi.mock('fs/promises');

import { findLibreOfficeBin } from '@process/utils/conversion/libreofficeUtils';
import { safeExecFile } from '@process/utils/safeExec';
import fs from 'fs/promises';
import { conversionService } from '@process/services/conversionService';

vi.mock('pptx2json', () => {
  const mockToJson = vi.fn();
  return { default: vi.fn().mockImplementation(() => ({ toJson: mockToJson })) };
});
import PPTX2Json from 'pptx2json';

describe('conversionService.pptToJson', () => {
  it('extracts slides in numeric order from flat key map', async () => {
    // A minimal synthetic flat-key map as returned by pptx2json.toJson()
    const fakeJson = {
      'ppt/slides/slide2.xml': { 'p:sld': { 'p:cSld': [{ 'p:spTree': [{}] }] } },
      'ppt/slides/slide1.xml': { 'p:sld': { 'p:cSld': [{ 'p:spTree': [{}] }] } },
      'ppt/slides/_rels/slide1.xml.rels': {},
      'ppt/slides/_rels/slide2.xml.rels': {},
    };

    const mockInstance = (PPTX2Json as ReturnType<typeof vi.fn>).mock.results[0]?.value;
    mockInstance.toJson.mockResolvedValue(fakeJson);

    const result = await conversionService.pptToJson('/fake/test.pptx');

    expect(result.success).toBe(true);
    expect(result.data?.slides).toHaveLength(2);
    expect(result.data?.slides[0].slideNumber).toBe(1);
    expect(result.data?.slides[1].slideNumber).toBe(2);
  });
});

describe('conversionService.pptToPdf', () => {
  beforeEach(() => vi.clearAllMocks());

  it('throws LIBRE_OFFICE_NOT_FOUND when no binary is detected', async () => {
    vi.mocked(findLibreOfficeBin).mockResolvedValue(null);
    const result = await conversionService.pptToPdf('/fake/test.pptx');
    expect(result.success).toBe(false);
    expect(result.error).toBe('LIBRE_OFFICE_NOT_FOUND');
  });

  it('returns pdf path on successful conversion', async () => {
    vi.mocked(findLibreOfficeBin).mockResolvedValue('/usr/bin/soffice');
    vi.mocked(safeExecFile).mockResolvedValue({ stdout: '', stderr: '' });
    vi.mocked(fs.access).mockResolvedValue(undefined);

    const result = await conversionService.pptToPdf('/fake/test.pptx');

    expect(result.success).toBe(true);
    expect(result.data).toMatch(/\.pdf$/);
    expect(safeExecFile).toHaveBeenCalledWith(
      '/usr/bin/soffice',
      expect.arrayContaining(['--headless', '--convert-to', 'pdf']),
      expect.objectContaining({ timeout: 30_000 })
    );
  });
});
```

- [ ] **Step 2: Run tests — expect FAIL**

```bash
bun run test tests/unit/process/services/conversionService.ppt.test.ts
```
Expected: FAIL.

- [ ] **Step 3: Fix `pptToJson` in `conversionService.ts`**

In `src/process/services/conversionService.ts`, rewrite `pptToJson`:

```ts
public async pptToJson(filePath: string): Promise<ConversionResult<PPTJsonData>> {
  try {
    const pptx2json = new PPTX2Json();
    const json = await pptx2json.toJson(filePath);

    // Extract slide keys from the flat key map in numeric order
    const slideKeys = Object.keys(json)
      .filter((k) => /^ppt\/slides\/slide\d+\.xml$/.test(k))
      .sort((a, b) => {
        const numA = parseInt(a.match(/slide(\d+)/)?.[1] ?? '0', 10);
        const numB = parseInt(b.match(/slide(\d+)/)?.[1] ?? '0', 10);
        return numA - numB;
      });

    const slides: PPTSlideData[] = slideKeys.map((key, idx) => {
      const slideNum = idx + 1;
      const slideXml = json[key];

      // Extract all text nodes (a:t) recursively
      const allTexts = extractTextNodes(slideXml);

      // Attempt to find title text via p:ph type="title" or first text block
      const title = extractTitle(slideXml) ?? allTexts[0] ?? '';
      const texts = allTexts.filter((t) => t !== title);

      // Extract images via slide rels
      const relsKey = key.replace('ppt/slides/', 'ppt/slides/_rels/') + '.rels';
      const images = extractSlideImages(json, relsKey);

      return { slideNumber: slideNum, title, texts, images };
    });

    return { success: true, data: { slides } };
  } catch (error) {
    return { success: false, error: error instanceof Error ? error.message : 'Unknown error' };
  }
}
```

Add these private helpers to `ConversionService` class:
```ts
private extractTextNodes(node: unknown): string[] {
  if (typeof node !== 'object' || node === null) return [];
  const result: string[] = [];
  const obj = node as Record<string, unknown>;
  // a:t contains the actual text
  if ('a:t' in obj) {
    const val = obj['a:t'];
    if (Array.isArray(val)) {
      val.forEach((v) => { if (typeof v === 'string' && v.trim()) result.push(v.trim()); });
    }
  }
  for (const child of Object.values(obj)) {
    if (Array.isArray(child)) {
      child.forEach((item) => result.push(...this.extractTextNodes(item)));
    } else if (typeof child === 'object') {
      result.push(...this.extractTextNodes(child));
    }
  }
  return result;
}

private extractTitle(slideXml: unknown): string | null {
  // Walk tree looking for p:ph with type="title" or "ctrTitle", return adjacent a:t text
  const text = this.findTitleText(slideXml);
  return text ?? null;
}

private findTitleText(node: unknown): string | null {
  if (typeof node !== 'object' || node === null) return null;
  const obj = node as Record<string, unknown>;
  // p:nvSpPr > p:nvPr > p:ph type="title"|"ctrTitle"
  if ('p:ph' in obj) {
    const ph = (obj['p:ph'] as any[])?.[0];
    const type = ph?.['$']?.type;
    if (type === 'title' || type === 'ctrTitle') {
      // Find nearest a:t in the parent p:sp
      return null; // signal: caller should collect a:t from this shape
    }
  }
  for (const child of Object.values(obj)) {
    const found = this.findTitleText(child);
    if (found !== null) return found;
  }
  return null;
}

private extractSlideImages(json: Record<string, unknown>, relsKey: string): string[] {
  const relsXml = json[relsKey] as Record<string, unknown> | undefined;
  if (!relsXml) return [];
  const images: string[] = [];
  try {
    // Rels XML parsed by xml2js: { Relationships: { Relationship: [...] } }
    const rels = (relsXml as any)?.Relationships?.Relationship ?? [];
    for (const rel of rels) {
      const target: string = rel?.['$']?.Target ?? '';
      const type: string = rel?.['$']?.Type ?? '';
      if (!type.includes('image')) continue;
      // Resolve relative path: from ppt/slides/_rels/ back to ppt/media/
      const mediaKey = 'ppt/' + target.replace(/^\.\.\//, '');
      const buf = json[mediaKey];
      if (!buf || !Buffer.isBuffer(buf)) continue;
      const mime = this.getMimeTypeFromName(mediaKey);
      images.push(`data:${mime};base64,${buf.toString('base64')}`);
    }
  } catch {
    // silently skip malformed rels
  }
  return images;
}
```

Also update the `extractTextNodes` / `extractTitle` calls in `pptToJson` to use `this.`:
```ts
const allTexts = this.extractTextNodes(slideXml);
const title = this.extractTitle(slideXml) ?? allTexts[0] ?? '';
```

- [ ] **Step 4: Add `pptToPdf` method to `ConversionService`**

Add these imports at the top of `conversionService.ts` (if not already present):
```ts
import crypto from 'crypto';
import os from 'os';
import path from 'path';
import { findLibreOfficeBin } from '../utils/conversion/libreofficeUtils';
import { safeExecFile } from '../utils/safeExec';
```

Add this in-memory set near the top of the class:
```ts
private readonly createdTempPdfs = new Set<string>();
```

Add the method:
```ts
public async pptToPdf(filePath: string): Promise<ConversionResult<string>> {
  try {
    const bin = await findLibreOfficeBin();
    if (!bin) {
      return { success: false, error: 'LIBRE_OFFICE_NOT_FOUND' };
    }

    // Deterministic output filename based on input path + mtime
    const stat = await fs.stat(filePath);
    const hash = crypto
      .createHash('sha256')
      .update(`${filePath}:${stat.mtimeMs}`)
      .digest('hex')
      .slice(0, 16);
    const outDir = os.tmpdir();
    const pdfPath = path.join(outDir, `aionui-ppt-${hash}.pdf`);

    // Skip conversion if cached PDF already exists
    try {
      await fs.access(pdfPath);
      return { success: true, data: pdfPath };
    } catch {
      // not cached — run conversion
    }

    await safeExecFile(bin, ['--headless', '--convert-to', 'pdf', '--outdir', outDir, filePath], {
      timeout: 30_000,
    });

    await fs.access(pdfPath); // verify output was created
    this.createdTempPdfs.add(pdfPath);
    return { success: true, data: pdfPath };
  } catch (error) {
    return { success: false, error: error instanceof Error ? error.message : 'Unknown error' };
  }
}

/** Clean up temp PDFs created this session. Call on app quit. */
public async cleanupTempPdfs(): Promise<void> {
  await Promise.allSettled([...this.createdTempPdfs].map((p) => fs.unlink(p)));
  this.createdTempPdfs.clear();
}
```

- [ ] **Step 5: Run tests — expect PASS**

```bash
bun run test tests/unit/process/services/conversionService.ppt.test.ts
```
Expected: PASS.

- [ ] **Step 6: Verify TypeScript**

```bash
bunx tsc --noEmit
```

- [ ] **Step 7: Commit**

```bash
git add src/process/services/conversionService.ts \
  tests/unit/process/services/conversionService.ppt.test.ts
git commit -m "feat(ppt-preview): fix pptToJson parsing and add pptToPdf via LibreOffice"
```

---

## Task 5: Update `documentBridge` — Add `ppt-pdf` and Before-Quit Cleanup

**Files:**
- Modify: `src/process/bridge/documentBridge.ts`

- [ ] **Step 1: Add `ppt-pdf` case to the IPC switch**

In `src/process/bridge/documentBridge.ts`, inside `initDocumentBridge()`, add after the `ppt-json` case:
```ts
case 'ppt-pdf': {
  if (!ensureExtension(filePath, PPT_EXTENSIONS)) {
    return unsupportedResult(to, 'Only PowerPoint files can be converted to PDF');
  }
  const result = await conversionService.pptToPdf(filePath);
  return { to, result };
}
```

- [ ] **Step 2: Register before-quit cleanup**

At the end of `initDocumentBridge()`, add:
```ts
import { app } from 'electron';

// Clean up temp PDFs when the app quits normally
app.on('before-quit', () => {
  void conversionService.cleanupTempPdfs();
});
```

(Add the `app` import at the top of the file alongside other Electron imports if not already present.)

- [ ] **Step 3: Verify TypeScript**

```bash
bunx tsc --noEmit
```
Expected: no errors.

- [ ] **Step 4: Commit**

```bash
git add src/process/bridge/documentBridge.ts
git commit -m "feat(ppt-preview): add ppt-pdf IPC handler and before-quit PDF cleanup"
```

---

## Task 6: Reorganize `viewers/` Directory — Create `office/` Sub-directory

**Files:**
- Create: `src/renderer/pages/conversation/Preview/components/viewers/office/`
- Move: `ExcelViewer.tsx`, `OfficeDocViewer.tsx`
- Modify: `viewers/index.ts`

- [ ] **Step 1: Create the directory and move files**

```bash
mkdir -p src/renderer/pages/conversation/Preview/components/viewers/office
cp src/renderer/pages/conversation/Preview/components/viewers/ExcelViewer.tsx \
   src/renderer/pages/conversation/Preview/components/viewers/office/ExcelViewer.tsx
cp src/renderer/pages/conversation/Preview/components/viewers/OfficeDocViewer.tsx \
   src/renderer/pages/conversation/Preview/components/viewers/office/OfficeDocViewer.tsx
```

- [ ] **Step 2: Delete old files**

```bash
rm src/renderer/pages/conversation/Preview/components/viewers/ExcelViewer.tsx
rm src/renderer/pages/conversation/Preview/components/viewers/OfficeDocViewer.tsx
```

- [ ] **Step 3: Update `viewers/index.ts` to point to the new paths**

Replace the current `ExcelViewer` and `OfficeDocViewer` export lines:
```ts
// Before:
export { default as ExcelViewer } from './ExcelViewer';
export { default as OfficeDocViewer } from './OfficeDocViewer';

// After:
export { default as ExcelViewer } from './office/ExcelViewer';
export { default as OfficeDocViewer } from './office/OfficeDocViewer';
```

- [ ] **Step 4: Verify TypeScript and lint**

```bash
bunx tsc --noEmit && bun run lint:fix
```
Expected: no errors.

- [ ] **Step 5: Commit**

```bash
git add src/renderer/pages/conversation/Preview/components/viewers/
git commit -m "refactor(preview): move office viewers into office/ subdirectory"
```

---

## Task 7: Add i18n Keys for All 6 Locales

**Files:**
- Modify: `src/renderer/services/i18n/locales/{en-US,zh-CN,zh-TW,ja-JP,ko-KR,tr-TR}/preview.json`

- [ ] **Step 1: Add `ppt` block to each locale's `preview.json`**

Add the following block before the final `}` in each file (after the existing `html` block):

**en-US:**
```json
  "ppt": {
    "loading": "Loading presentation...",
    "converting": "Converting with LibreOffice...",
    "fallbackHint": "Basic preview (LibreOffice not found)",
    "slideCount": "{{current}} / {{total}}",
    "prevSlide": "Previous",
    "nextSlide": "Next",
    "noContent": "This slide has no text content",
    "viaLibreOffice": "via LibreOffice"
  }
```

**zh-CN:**
```json
  "ppt": {
    "loading": "正在加载演示文稿...",
    "converting": "正在通过 LibreOffice 转换...",
    "fallbackHint": "基础预览（未检测到 LibreOffice）",
    "slideCount": "{{current}} / {{total}}",
    "prevSlide": "上一张",
    "nextSlide": "下一张",
    "noContent": "此幻灯片无文字内容",
    "viaLibreOffice": "via LibreOffice"
  }
```

**zh-TW:**
```json
  "ppt": {
    "loading": "正在載入簡報...",
    "converting": "正在透過 LibreOffice 轉換...",
    "fallbackHint": "基本預覽（未偵測到 LibreOffice）",
    "slideCount": "{{current}} / {{total}}",
    "prevSlide": "上一張",
    "nextSlide": "下一張",
    "noContent": "此投影片無文字內容",
    "viaLibreOffice": "via LibreOffice"
  }
```

**ja-JP:**
```json
  "ppt": {
    "loading": "プレゼンテーションを読み込み中...",
    "converting": "LibreOffice で変換中...",
    "fallbackHint": "基本プレビュー（LibreOffice が見つかりません）",
    "slideCount": "{{current}} / {{total}}",
    "prevSlide": "前へ",
    "nextSlide": "次へ",
    "noContent": "このスライドにテキストはありません",
    "viaLibreOffice": "via LibreOffice"
  }
```

**ko-KR:**
```json
  "ppt": {
    "loading": "프레젠테이션 로딩 중...",
    "converting": "LibreOffice로 변환 중...",
    "fallbackHint": "기본 미리보기 (LibreOffice 없음)",
    "slideCount": "{{current}} / {{total}}",
    "prevSlide": "이전",
    "nextSlide": "다음",
    "noContent": "이 슬라이드에 텍스트가 없습니다",
    "viaLibreOffice": "via LibreOffice"
  }
```

**tr-TR:**
```json
  "ppt": {
    "loading": "Sunum yükleniyor...",
    "converting": "LibreOffice ile dönüştürülüyor...",
    "fallbackHint": "Temel önizleme (LibreOffice bulunamadı)",
    "slideCount": "{{current}} / {{total}}",
    "prevSlide": "Önceki",
    "nextSlide": "Sonraki",
    "noContent": "Bu slaytda metin içeriği yok",
    "viaLibreOffice": "via LibreOffice"
  }
```

- [ ] **Step 2: Verify JSON is valid**

```bash
for f in src/renderer/services/i18n/locales/*/preview.json; do
  node -e "JSON.parse(require('fs').readFileSync('$f','utf8'))" && echo "$f OK"
done
```
Expected: all 6 files report OK.

- [ ] **Step 3: Commit**

```bash
git add src/renderer/services/i18n/locales/
git commit -m "feat(ppt-preview): add ppt i18n keys for all 6 locales"
```

---

## Task 8: Create `PPTViewer` Component

**Files:**
- Create: `src/renderer/pages/conversation/Preview/components/viewers/office/PPTViewer.tsx`
- Create: `tests/unit/renderer/PPTViewer.dom.test.tsx`

- [ ] **Step 1: Write the failing component tests**

Create `tests/unit/renderer/PPTViewer.dom.test.tsx`:
```tsx
import { describe, it, expect, vi, beforeEach } from 'vitest';
import { render, screen, fireEvent, waitFor } from '@testing-library/react';
import React from 'react';

// Mock ipcBridge
vi.mock('@/common', () => ({
  ipcBridge: {
    document: {
      convert: { invoke: vi.fn() },
    },
    shell: {
      openFile: { invoke: vi.fn() },
    },
  },
}));

// Mock i18next
vi.mock('react-i18next', () => ({
  useTranslation: () => ({ t: (k: string) => k }),
}));

// Mock PDFViewer
vi.mock('@renderer/pages/conversation/Preview/components/viewers/PDFViewer', () => ({
  default: ({ filePath }: { filePath: string }) => <div data-testid="pdf-viewer">{filePath}</div>,
}));

import { ipcBridge } from '@/common';
import PPTViewer from '@renderer/pages/conversation/Preview/components/viewers/office/PPTViewer';

const mockSlides = [
  { slideNumber: 1, title: 'Slide One', texts: ['Body text'], images: [] },
  { slideNumber: 2, title: 'Slide Two', texts: [], images: [] },
];

describe('PPTViewer', () => {
  beforeEach(() => vi.clearAllMocks());

  it('renders PDFViewer when ppt-pdf succeeds', async () => {
    vi.mocked(ipcBridge.document.convert.invoke).mockResolvedValue({
      to: 'ppt-pdf',
      result: { success: true, data: '/tmp/test.pdf' },
    });

    render(<PPTViewer filePath="/test.pptx" />);

    await waitFor(() => expect(screen.getByTestId('pdf-viewer')).toBeTruthy());
    expect(screen.getByText('/tmp/test.pdf')).toBeTruthy();
  });

  it('renders fallback slides when ppt-pdf fails', async () => {
    vi.mocked(ipcBridge.document.convert.invoke)
      .mockResolvedValueOnce({ to: 'ppt-pdf', result: { success: false, error: 'LIBRE_OFFICE_NOT_FOUND' } })
      .mockResolvedValueOnce({ to: 'ppt-json', result: { success: true, data: { slides: mockSlides } } });

    render(<PPTViewer filePath="/test.pptx" />);

    await waitFor(() => expect(screen.getByText('Slide One')).toBeTruthy());
    // Shows fallback hint badge
    expect(screen.getByText('preview.ppt.fallbackHint')).toBeTruthy();
  });

  it('navigates slides with ArrowRight key', async () => {
    vi.mocked(ipcBridge.document.convert.invoke)
      .mockResolvedValueOnce({ to: 'ppt-pdf', result: { success: false, error: 'LIBRE_OFFICE_NOT_FOUND' } })
      .mockResolvedValueOnce({ to: 'ppt-json', result: { success: true, data: { slides: mockSlides } } });

    render(<PPTViewer filePath="/test.pptx" />);
    await waitFor(() => screen.getByText('Slide One'));

    fireEvent.keyDown(window, { key: 'ArrowRight' });
    await waitFor(() => expect(screen.getByText('Slide Two')).toBeTruthy());
  });
});
```

- [ ] **Step 2: Run the tests — expect FAIL**

```bash
bun run test tests/unit/renderer/PPTViewer.dom.test.tsx
```
Expected: FAIL — module not found.

- [ ] **Step 3: Create `PPTViewer.tsx`**

Create `src/renderer/pages/conversation/Preview/components/viewers/office/PPTViewer.tsx`:
```tsx
import { ipcBridge } from '@/common';
import type { PPTSlideData } from '@/common/types/conversion';
import { usePreviewToolbarExtras } from '../../../../context/PreviewToolbarExtrasContext';
import { Button, Message } from '@arco-design/web-react';
import React, { useCallback, useEffect, useRef, useState } from 'react';
import { useTranslation } from 'react-i18next';
import PDFViewer from '../PDFViewer';

interface PPTViewerProps {
  filePath?: string;
  hideToolbar?: boolean;
}

type ViewerState =
  | { status: 'loading' }
  | { status: 'pdf'; pdfPath: string }
  | { status: 'fallback-loading' }
  | { status: 'slides'; slides: PPTSlideData[]; currentIndex: number; usedFallback: true }
  | { status: 'error'; message: string };

const PPTViewer: React.FC<PPTViewerProps> = ({ filePath, hideToolbar = false }) => {
  const { t } = useTranslation();
  const [state, setState] = useState<ViewerState>({ status: 'loading' });
  const [messageApi, messageContextHolder] = Message.useMessage();
  const toolbarExtrasContext = usePreviewToolbarExtras();
  const usePortalToolbar = Boolean(toolbarExtrasContext) && !hideToolbar;
  const messageApiRef = useRef(messageApi);
  useEffect(() => { messageApiRef.current = messageApi; }, [messageApi]);

  // Load: try ppt-pdf, fall back to ppt-json on any failure
  useEffect(() => {
    if (!filePath) {
      setState({ status: 'error', message: t('preview.errors.missingFilePath') });
      return;
    }

    const load = async () => {
      setState({ status: 'loading' });

      // Step 1: try LibreOffice PDF conversion
      try {
        const pdfResp = await ipcBridge.document.convert.invoke({ filePath, to: 'ppt-pdf' });
        if (pdfResp.to === 'ppt-pdf' && pdfResp.result.success && pdfResp.result.data) {
          setState({ status: 'pdf', pdfPath: pdfResp.result.data });
          return;
        }
      } catch {
        // proceed to fallback
      }

      // Step 2: fallback — pptx2json
      setState({ status: 'fallback-loading' });
      try {
        const jsonResp = await ipcBridge.document.convert.invoke({ filePath, to: 'ppt-json' });
        if (jsonResp.to === 'ppt-json' && jsonResp.result.success && jsonResp.result.data) {
          setState({
            status: 'slides',
            slides: jsonResp.result.data.slides,
            currentIndex: 0,
            usedFallback: true,
          });
          return;
        }
        throw new Error(jsonResp.to === 'ppt-json' ? (jsonResp.result.error ?? '') : '');
      } catch (err) {
        setState({
          status: 'error',
          message: err instanceof Error ? err.message : t('preview.word.loadFailed'),
        });
      }
    };

    void load();
  }, [filePath, t]);

  // Keyboard navigation
  useEffect(() => {
    if (state.status !== 'slides') return;
    const handler = (e: KeyboardEvent) => {
      if (e.key === 'ArrowRight') {
        setState((prev) =>
          prev.status === 'slides' && prev.currentIndex < prev.slides.length - 1
            ? { ...prev, currentIndex: prev.currentIndex + 1 }
            : prev
        );
      } else if (e.key === 'ArrowLeft') {
        setState((prev) =>
          prev.status === 'slides' && prev.currentIndex > 0
            ? { ...prev, currentIndex: prev.currentIndex - 1 }
            : prev
        );
      }
    };
    window.addEventListener('keydown', handler);
    return () => window.removeEventListener('keydown', handler);
  }, [state.status]);

  // Toolbar extras
  useEffect(() => {
    if (!usePortalToolbar || !toolbarExtrasContext) return;
    if (state.status !== 'slides' && state.status !== 'pdf') return;

    const badge =
      state.status === 'pdf'
        ? t('preview.ppt.viaLibreOffice')
        : t('preview.ppt.fallbackHint');

    toolbarExtrasContext.setExtras({
      left: <span className='text-11px text-t-tertiary'>{badge}</span>,
      right: null,
    });
    return () => toolbarExtrasContext.setExtras(null);
  }, [usePortalToolbar, toolbarExtrasContext, state.status, t]);

  const handleOpenInSystem = useCallback(async () => {
    if (!filePath) return;
    try {
      await ipcBridge.shell.openFile.invoke(filePath);
    } catch {
      messageApiRef.current?.error?.(t('preview.openInSystemFailed'));
    }
  }, [filePath, t]);

  // --- Render ---

  if (state.status === 'loading' || state.status === 'fallback-loading') {
    const label = state.status === 'loading' ? t('preview.ppt.loading') : t('preview.ppt.converting');
    return (
      <div className='flex items-center justify-center h-full'>
        <div className='text-14px text-t-secondary'>{label}</div>
      </div>
    );
  }

  if (state.status === 'pdf') {
    return (
      <div className='h-full w-full flex flex-col'>
        {messageContextHolder}
        {!usePortalToolbar && !hideToolbar && (
          <div className='flex items-center justify-between h-40px px-12px bg-bg-2 shrink-0'>
            <span className='text-11px text-t-tertiary'>{t('preview.ppt.viaLibreOffice')}</span>
          </div>
        )}
        <div className='flex-1 overflow-hidden'>
          <PDFViewer filePath={state.pdfPath} hideToolbar />
        </div>
      </div>
    );
  }

  if (state.status === 'error') {
    return (
      <div className='flex items-center justify-center h-full'>
        {messageContextHolder}
        <div className='text-center'>
          <div className='text-16px text-t-error mb-8px'>{state.message}</div>
          {filePath && (
            <Button size='small' onClick={handleOpenInSystem}>
              {t('preview.openInSystemApp')}
            </Button>
          )}
        </div>
      </div>
    );
  }

  // state.status === 'slides'
  const { slides, currentIndex } = state;
  const slide = slides[currentIndex];
  const total = slides.length;

  return (
    <div className='h-full w-full flex flex-col bg-bg-1'>
      {messageContextHolder}

      {/* Toolbar */}
      {!usePortalToolbar && !hideToolbar && (
        <div className='flex items-center justify-between h-40px px-12px bg-bg-2 shrink-0'>
          <span className='text-11px text-t-tertiary'>{t('preview.ppt.fallbackHint')}</span>
          {filePath && (
            <Button size='mini' type='text' onClick={handleOpenInSystem}>
              {t('preview.openInSystemApp')}
            </Button>
          )}
        </div>
      )}

      {/* Slide content — fixed 16:9 aspect ratio card */}
      <div className='flex-1 overflow-hidden flex items-center justify-center p-16px'>
        <div
          className='w-full bg-bg-2 rd-8px shadow-sm overflow-hidden flex flex-col'
          style={{ aspectRatio: '16/9', maxHeight: '100%' }}
        >
          <div className='flex-1 overflow-y-auto p-24px'>
            {slide.title && (
              <div className='text-18px font-semibold text-t-primary mb-12px leading-snug'>
                {slide.title}
              </div>
            )}
            {slide.texts.map((text, i) => (
              <p key={i} className='text-14px text-t-secondary leading-relaxed mb-6px'>
                {text}
              </p>
            ))}
            {slide.images.length > 0 && (
              <div className='flex flex-wrap gap-8px mt-12px'>
                {slide.images.map((src, i) => (
                  <img
                    key={i}
                    src={src}
                    alt=""
                    className='object-contain rd-4px'
                    style={{ maxHeight: '200px', maxWidth: '48%' }}
                  />
                ))}
              </div>
            )}
            {!slide.title && slide.texts.length === 0 && slide.images.length === 0 && (
              <div className='text-13px text-t-tertiary'>{t('preview.ppt.noContent')}</div>
            )}
          </div>
        </div>
      </div>

      {/* Navigation footer */}
      <div className='flex items-center justify-center gap-16px h-48px shrink-0 border-t border-border-base'>
        <Button
          size='small'
          disabled={currentIndex === 0}
          onClick={() => setState((prev) => prev.status === 'slides' ? { ...prev, currentIndex: prev.currentIndex - 1 } : prev)}
        >
          {t('preview.ppt.prevSlide')}
        </Button>
        <span className='text-13px text-t-secondary min-w-48px text-center'>
          {t('preview.ppt.slideCount', { current: currentIndex + 1, total })}
        </span>
        <Button
          size='small'
          disabled={currentIndex === total - 1}
          onClick={() => setState((prev) => prev.status === 'slides' ? { ...prev, currentIndex: prev.currentIndex + 1 } : prev)}
        >
          {t('preview.ppt.nextSlide')}
        </Button>
      </div>
    </div>
  );
};

export default PPTViewer;
```

- [ ] **Step 4: Run tests — expect PASS**

```bash
bun run test tests/unit/renderer/PPTViewer.dom.test.tsx
```
Expected: PASS.

- [ ] **Step 5: Lint and type check**

```bash
bun run lint:fix && bunx tsc --noEmit
```

- [ ] **Step 6: Commit**

```bash
git add src/renderer/pages/conversation/Preview/components/viewers/office/PPTViewer.tsx \
  tests/unit/renderer/PPTViewer.dom.test.tsx
git commit -m "feat(ppt-preview): add PPTViewer component with LibreOffice PDF and pptx2json fallback"
```

---

## Task 9: Wire Up — Update `OfficeDocViewer` and `viewers/index.ts`

**Files:**
- Modify: `src/renderer/pages/conversation/Preview/components/viewers/office/OfficeDocViewer.tsx`
- Modify: `src/renderer/pages/conversation/Preview/components/viewers/index.ts`

- [ ] **Step 1: Update `OfficeDocViewer.tsx` — replace ppt placeholder with `PPTViewer`**

In `src/renderer/pages/conversation/Preview/components/viewers/office/OfficeDocViewer.tsx`, add the import at the top:
```tsx
import PPTViewer from './PPTViewer';
```

Then replace the entire `if (docType === 'ppt')` block (lines 151–175) with:
```tsx
if (docType === 'ppt') {
  return <PPTViewer filePath={filePath} hideToolbar={hideToolbar} />;
}
```

Also remove the now-unused `MarkdownPreview` import from `OfficeDocViewer.tsx` if it's only used for the word render path — it is still needed for word documents so leave it.

- [ ] **Step 2: Export `PPTViewer` from `viewers/index.ts`**

Add to `src/renderer/pages/conversation/Preview/components/viewers/index.ts`:
```ts
export { default as PPTViewer } from './office/PPTViewer';
```

- [ ] **Step 3: Run the full test suite**

```bash
bun run test
```
Expected: all tests pass.

- [ ] **Step 4: Lint and type check**

```bash
bun run lint:fix && bunx tsc --noEmit
```

- [ ] **Step 5: Commit**

```bash
git add src/renderer/pages/conversation/Preview/components/viewers/office/OfficeDocViewer.tsx \
  src/renderer/pages/conversation/Preview/components/viewers/index.ts
git commit -m "feat(ppt-preview): wire PPTViewer into OfficeDocViewer and export from viewers index"
```

---

## Final Verification

- [ ] Open a `.pptx` file in the app — verify the preview panel shows slides (fallback or PDF)
- [ ] Confirm keyboard `ArrowLeft`/`ArrowRight` navigates slides in fallback mode
- [ ] Run `bun run test` — all tests green
- [ ] Run `bunx tsc --noEmit` — no type errors
- [ ] Run `bun run lint:fix` — no lint errors
