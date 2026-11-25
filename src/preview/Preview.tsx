import { inspectXml } from "../xml/xml";

export default function Preview({ xml }: { xml: string }) {
  if (!xml) return <p className="muted">Load OOXML from Word to see its structural preview.</p>;
  try {
    const info = inspectXml(xml);
    return <section className="preview" aria-label="Structural preview">
      <p className="notice">Structural preview only — Word is the visual authority.</p>
      {Array.from({ length: Math.min(info.paragraphs, 30) }, (_, i) => <p key={i}>Paragraph {i + 1}</p>)}
      {info.tables > 0 && <p>▦ {info.tables} table{info.tables === 1 ? "" : "s"}</p>}
      {!info.paragraphs && !info.tables && <p className="muted">No paragraph or table nodes detected.</p>}
    </section>;
  } catch { return <p className="muted">Preview is unavailable until XML is valid.</p>; }
}
