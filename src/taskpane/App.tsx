import { useMemo, useState } from "react";
import Editor from "@monaco-editor/react";
import { diffLines } from "../xml/diff";
import { inspectXml, prettyXml, validateXml } from "../xml/xml";
import { applySelectionOoxml, capabilityMessage, readOoxml, Scope } from "../office/word";
import Preview from "../preview/Preview";

export default function App() {
  const [scope, setScope] = useState<Scope>("selection");
  const [xml, setXml] = useState("");
  const [snapshot, setSnapshot] = useState("");
  const [status, setStatus] = useState(capabilityMessage());
  const [busy, setBusy] = useState(false);
  const [review, setReview] = useState(false);
  const dirty = xml !== snapshot;
  const validation = useMemo(() => validateXml(xml), [xml]);
  const info = useMemo(() => xml && validation.valid ? inspectXml(xml) : undefined, [xml, validation.valid]);

  async function refresh(force = false) {
    if (dirty && !force && !window.confirm("Discard unsaved editor changes and refresh from Word?")) return;
    setBusy(true);
    try {
      const loaded = await readOoxml(scope);
      const formatted = prettyXml(loaded.xml);
      setXml(formatted); setSnapshot(formatted); setStatus(`Loaded ${loaded.label} from Word.`);
    } catch (error) { setStatus(`Could not load OOXML: ${message(error)}`); }
    finally { setBusy(false); }
  }
  function format() {
    try { setXml(prettyXml(xml)); setStatus("XML formatted. Text-node content was not normalized."); }
    catch (error) { setStatus(`Cannot format invalid XML: ${message(error)}`); }
  }
  async function apply() {
    if (scope !== "selection") return;
    if (!validation.valid) { setStatus(`Fix XML before applying: ${validation.message}`); return; }
    if (!window.confirm("Replace the current Word selection with this XML? Use Word Undo if needed.")) return;
    setBusy(true);
    try {
      await applySelectionOoxml(xml);
      setStatus("Applied to Word. Reloading Word's accepted XML…");
      await refresh(true);
    } catch (error) { setStatus(`Word rejected the XML. Your editor changes are still here: ${message(error)}`); }
    finally { setBusy(false); }
  }
  const diff = useMemo(() => diffLines(snapshot, xml), [snapshot, xml]);
  return <main>
    <header><h1>DocXray</h1><p className="status" role="status">{status}{dirty ? " · unsaved local changes" : ""}</p></header>
    <div className="toolbar">
      <label>Scope <select value={scope} onChange={e => setScope(e.target.value as Scope)} disabled={busy}><option value="selection">Selection</option><option value="body">Document body (read-only)</option></select></label>
      <button onClick={() => refresh()} disabled={busy}>Refresh from Word</button>
      <button onClick={() => { setXml(snapshot); setStatus("Local changes discarded."); }} disabled={!dirty || busy}>Discard</button>
    </div>
    <div className="grid">
      <section className="editor-pane" aria-label="OOXML editor">
        <Editor height="56vh" language="xml" theme="vs" value={xml} onChange={v => setXml(v ?? "")} options={{ minimap: { enabled: false }, wordWrap: "on", automaticLayout: true, tabSize: 2 }} />
      </section>
      <aside>
        <h2>Preview</h2><Preview xml={xml} />
        <h2>Inspector</h2>
        <dl><dt>Scope</dt><dd>{scope === "selection" ? "Current selection" : "Document body (read-only)"}</dd><dt>Size</dt><dd>{xml.length.toLocaleString()} characters</dd><dt>Validation</dt><dd className={validation.valid ? "good" : "bad"}>{xml ? validation.valid ? "Well-formed" : validation.message : "No XML loaded"}</dd>
        {info && <><dt>Paragraphs / runs / tables</dt><dd>{info.paragraphs} / {info.runs} / {info.tables}</dd><dt>Paragraph styles</dt><dd>{info.styles.join(", ") || "None detected"}</dd></>}</dl>
      </aside>
    </div>
    <footer><button onClick={format} disabled={!xml || busy}>Format XML</button><button onClick={() => setStatus(validation.valid ? "XML is well-formed and passed basic safety checks." : `Invalid XML: ${validation.message}`)} disabled={!xml}>Validate</button><button onClick={() => setReview(true)} disabled={!dirty || busy}>Review changes</button><button className="primary" onClick={apply} disabled={!dirty || !validation.valid || busy || scope !== "selection"}>Apply to Word</button></footer>
    {review && <div className="modal-backdrop" role="dialog" aria-modal="true" aria-label="Review XML changes"><section className="modal"><h2>Review changes</h2><p>Only selection replacement is enabled. Word may reject unsupported OOXML.</p><pre className="diff">{diff.map((part, i) => <span key={i} className={part.added ? "added" : part.removed ? "removed" : ""}>{part.value}</span>)}</pre><button onClick={() => setReview(false)}>Close</button></section></div>}
  </main>;
}
function message(error: unknown) { return error instanceof Error ? error.message : String(error); }
