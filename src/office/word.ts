export type Scope = "selection" | "body";
export type WordDocument = { xml: string; label: string };

function requireWord() {
  if (!window.Office) throw new Error("Open DocXray inside Microsoft Word to read or apply OOXML.");
  if (!Office.context.requirements.isSetSupported("WordApi", "1.1")) {
    throw new Error("This Word installation does not support the required WordApi 1.1 APIs.");
  }
}

export async function readOoxml(scope: Scope): Promise<WordDocument> {
  requireWord();
  return Word.run(async context => {
    const range = scope === "selection" ? context.document.getSelection() : context.document.body;
    const result = range.getOoxml();
    await context.sync();
    return { xml: result.value, label: scope === "selection" ? "Current selection" : "Document body" };
  });
}

/** Deliberately selection-only until body replacement is verified on supported Mac builds. */
export async function applySelectionOoxml(xml: string): Promise<void> {
  requireWord();
  await Word.run(async context => {
    context.document.getSelection().insertOoxml(xml, Word.InsertLocation.replace);
    await context.sync();
  });
}

export function capabilityMessage() {
  if (!window.Office) return "Browser development mode — connect to Word to load OOXML.";
  return Office.context.requirements.isSetSupported("WordApi", "1.1")
    ? "WordApi 1.1 available"
    : "WordApi 1.1 is unavailable; XML controls are disabled.";
}
