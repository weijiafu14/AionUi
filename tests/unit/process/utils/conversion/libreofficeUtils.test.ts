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
