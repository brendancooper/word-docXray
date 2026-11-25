export type Validation = { valid: boolean; message?: string; line?: number; column?: number };

/** Browser XML validation. DOMParser does not resolve external entities. */
export function validateXml(xml: string): Validation {
  if (!xml.trim()) return { valid: false, message: "XML is empty." };
  if (/<!DOCTYPE|<!ENTITY/i.test(xml)) return { valid: false, message: "DTD and entity declarations are not allowed." };
  // Namespace URIs are normal OOXML and are not external content. Reject explicit external relationship targets instead.
  if (/TargetMode\s*=\s*["']External["']/i.test(xml) || /<\/?(?:\w+:)?Relationship\b[^>]*\bTarget\s*=\s*["'](?:https?|file):\/\//i.test(xml)) {
    return { valid: false, message: "External relationship targets are not allowed in editable OOXML." };
  }
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const error = doc.querySelector("parsererror");
  if (!error) return { valid: true };
  const message = error.textContent?.replace(/\s+/g, " ").trim() || "Malformed XML.";
  const match = message.match(/(?:line|Line)\s*(\d+)[^\d]+(?:column|column)\s*(\d+)/);
  return { valid: false, message, line: match ? Number(match[1]) : undefined, column: match ? Number(match[2]) : undefined };
}

/** Lightweight token formatter: it never normalizes text-node whitespace or attribute values. */
export function prettyXml(xml: string): string {
  const result = validateXml(xml);
  if (!result.valid) throw new Error(result.message);
  const tokens = xml.match(/<!--[\s\S]*?-->|<\?[^]*?\?>|<!\[CDATA\[[\s\S]*?\]\]>|<[^>]+>|[^<]+/g) ?? [];
  let depth = 0;
  const out: string[] = [];
  let pendingText = false;
  for (const token of tokens) {
    if (!token.trim()) continue;
    if (token.startsWith("</")) {
      depth = Math.max(0, depth - 1);
      if (pendingText) { out[out.length - 1] += token; pendingText = false; }
      else out.push(`${"  ".repeat(depth)}${token}`);
    } else if (token.startsWith("<") && !token.startsWith("<?") && !token.startsWith("<!") && !token.endsWith("/>")) {
      out.push(`${"  ".repeat(depth)}${token}`); depth++; pendingText = false;
    } else if (token.startsWith("<")) { out.push(`${"  ".repeat(depth)}${token}`); pendingText = false; }
    else { out[out.length - 1] = `${out[out.length - 1] ?? ""}${token}`; pendingText = true; }
  }
  return out.join("\n");
}

export function inspectXml(xml: string) {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const count = (name: string) => doc.getElementsByTagNameNS("*", name).length;
  const styles = [...doc.getElementsByTagNameNS("*", "pStyle")].map(n => n.getAttribute("w:val") || n.getAttribute("val")).filter(Boolean);
  return { paragraphs: count("p"), runs: count("r"), tables: count("tbl"), styles: [...new Set(styles)] as string[] };
}
