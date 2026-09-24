/**
 * Well-formedness check for the XML parts of a saved HWPX package.
 *
 * save_document(verify_integrity) used to look for three textual symptoms
 * (no `<?xml`, a dangling `<` at the end, `<` inside a tag). A section with a
 * mismatched close tag passed all three and was reported as
 * `integrity_verified: true`, while Hancom and every XML parser rejected it
 * (reported 2026-09-24: `<hp:p>` 5730 open / 5728 close after one edit).
 * A real parser is the only check that means what the flag claims.
 *
 * Measured on 275 Hancom-saved originals (2,046 XML parts, 168 MB): 0 false
 * rejections, ~1.5 s total. The only rejected sample was a password-protected
 * file, whose parts are ciphertext rather than XML.
 */
import { SaxesParser } from 'saxes';
import type JSZip from 'jszip';

/** First well-formedness error in `xml`, or null when it parses. */
export function xmlWellFormednessError(xml: string): string | null {
  // xmlns: true also rejects an undeclared prefix (<hs:sec> with only
  // xmlns:hp declared). Hancom declares every prefix it uses (0 rejections
  // across the 275-file corpus with this setting), while set_section_xml
  // accepted a section missing xmlns:hs.
  const parser = new SaxesParser({ xmlns: true });
  let first: string | null = null;
  parser.on('error', err => {
    if (first === null) first = err.message;
  });
  try {
    parser.write(xml).close();
  } catch (err) {
    first ??= err instanceof Error ? err.message : String(err);
  }
  return first;
}

/** Every `.xml` / `.hpf` part of the package that does not parse, as "path: error". */
export async function findMalformedXmlParts(zip: JSZip): Promise<string[]> {
  const out: string[] = [];
  const names = Object.keys(zip.files)
    .filter(name => !zip.files[name].dir && /\.(xml|hpf)$/i.test(name))
    .sort();
  for (const name of names) {
    const err = xmlWellFormednessError(await zip.file(name)!.async('string'));
    if (err) out.push(`${name}: ${err}`);
  }
  return out;
}
