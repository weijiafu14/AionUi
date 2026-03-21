import fs from "fs/promises";

const LIBRE_OFFICE_PATHS: Partial<Record<NodeJS.Platform, string[]>> = {
  darwin: [
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/usr/local/bin/soffice",
  ],
  win32: [
    "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
    "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
  ],
  linux: [
    "/usr/bin/soffice",
    "/usr/local/bin/soffice",
    "/snap/bin/libreoffice",
  ],
};

let cachedBin: string | null | undefined = undefined;

/**
 * Find the LibreOffice soffice binary on the current platform.
 * Result is cached for the process lifetime.
 */
export async function findLibreOfficeBin(): Promise<string | null> {
  if (cachedBin !== undefined) {
    return cachedBin;
  }

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
