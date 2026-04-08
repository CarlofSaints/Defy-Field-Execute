/**
 * SharePoint folder path normaliser.
 *
 * Accepts any of the following and returns a library-relative path
 * (ready to pass to listFilesInSPFolder / uploadToSharePoint):
 *
 *   1. Library-relative path (already correct):
 *      "DEFY/PERIGEE - FG/2. EXTERNAL SYNC/REPORTS/PERIGEE IMAGE DOWNLOADS"
 *
 *   2. SharePoint address-bar URL with ?id= query parameter:
 *      "https://exceler8xl.sharepoint.com/Clients/Forms/AllItems.aspx?id=%2FClients%2FDEFY%2F..."
 *
 *   3. SharePoint "Copy link" share URL (/:f:/r/... format):
 *      "https://exceler8xl.sharepoint.com/:f:/r/Clients/DEFY/PERIGEE%20-%20FG/...?csf=1&e=..."
 *      "https://exceler8xl.sharepoint.com/:f:/r/sites/SiteName/Clients/DEFY/..."
 *
 * NOT supported (return the raw input — will fail gracefully at runtime):
 *   - Tokenised short share links: "https://.../:f:/g/AbCdEf12345?e=..."
 *     These require a Graph API /shares lookup to resolve. Navigate into
 *     the folder and use the address-bar URL instead.
 *
 * Isomorphic — safe to call from both server and client (no Node-only imports).
 */
export function parseSpPath(input: string, libraryName = 'Clients'): string {
  if (!input) return '';
  const raw = input.trim();
  if (!raw) return '';

  // Not a URL → assume library-relative path (just clean leading/trailing slashes)
  if (!/^https?:\/\//i.test(raw)) {
    return stripSlashes(raw);
  }

  let url: URL;
  try {
    url = new URL(raw);
  } catch {
    return raw;
  }

  // Case A: ?id=<server-relative-path> (address bar, Forms/AllItems.aspx?id=...)
  const idParam = url.searchParams.get('id');
  if (idParam) {
    return stripLibraryPrefix(idParam, libraryName);
  }

  // Case B: /:f:/r/... or /:w:/r/... or /:x:/r/... path prefixes (Copy link)
  let pathname: string;
  try {
    pathname = decodeURIComponent(url.pathname);
  } catch {
    pathname = url.pathname;
  }

  // Strip the /:f:/r/ (or similar) short-URL routing prefix
  pathname = pathname.replace(/^\/:[a-z]:\/[a-z]\//i, '/');

  return stripLibraryPrefix(pathname, libraryName);
}

/** Strip leading/trailing slashes and normalise backslashes. */
function stripSlashes(p: string): string {
  return p.replace(/\\/g, '/').replace(/^\/+/, '').replace(/\/+$/, '');
}

/**
 * Given a server-relative path like "/sites/Foo/Clients/DEFY/..." or
 * "/Clients/DEFY/...", strip the site and library prefixes to return
 * the library-relative path ("DEFY/...").
 */
function stripLibraryPrefix(path: string, library: string): string {
  let p: string;
  try {
    p = decodeURIComponent(path);
  } catch {
    p = path;
  }
  p = stripSlashes(p);

  // Strip "sites/<sitename>/" prefix if present
  p = p.replace(/^sites\/[^/]+\//i, '');

  // Strip library name prefix if present (case-insensitive, escaped)
  const libEscaped = library.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const libRegex = new RegExp(`^${libEscaped}(?:/|$)`, 'i');
  p = p.replace(libRegex, '');

  return stripSlashes(p);
}
