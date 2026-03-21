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
    console.log("[libreoffice] returning cached result:", cachedBin);
    return cachedBin;
  }

  const paths = LIBRE_OFFICE_PATHS[process.platform] ?? [];
  console.log(
    "[libreoffice] platform:",
    process.platform,
    "checking paths:",
    paths,
  );
  for (const p of paths) {
    try {
      await fs.access(p);
      console.log("[libreoffice] found binary at:", p);
      cachedBin = p;
      return cachedBin;
    } catch (err) {
      console.log(
        "[libreoffice] not found at:",
        p,
        (err as NodeJS.ErrnoException).code,
      );
    }
  }
  console.log("[libreoffice] no binary found on this system");
  cachedBin = null;
  return null;
}

/** Reset cache — for testing only. */
export function _resetLibreOfficeBinCache(): void {
  cachedBin = undefined;
}
