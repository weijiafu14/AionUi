import { describe, it, expect, vi, beforeEach } from 'vitest';

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
  // Use regular function (not arrow) so the mock works when called with `new`
  return {
    default: vi.fn().mockImplementation(function () {
      return { toJson: mockToJson };
    }),
  };
});
import PPTX2Json from 'pptx2json';

describe('conversionService.pptToJson', () => {
  it('extracts slides in numeric order from flat key map', async () => {
    // A minimal synthetic flat-key map as returned by pptx2json.toJson()
    const fakeJson = {
      'ppt/slides/slide2.xml': {
        'p:sld': { 'p:cSld': [{ 'p:spTree': [{}] }] },
      },
      'ppt/slides/slide1.xml': {
        'p:sld': { 'p:cSld': [{ 'p:spTree': [{}] }] },
      },
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
