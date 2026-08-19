/* global Office, Word */
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("inspect").onclick = inspectDocument;
  }
});

function metric(label, value) {
  return `<div class="metric"><span>${label}</span><strong>${value}</strong></div>`;
}

async function inspectDocument() {
  const status = document.getElementById("status");
  const results = document.getElementById("results");
  status.textContent = "Inspecting locally inside Word…";
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      const ooxml = body.getOoxml();
      body.load("text");
      await context.sync();

      const xml = ooxml.value || "";
      const text = body.text || "";
      const count = (pattern) => (xml.match(pattern) || []).length;
      const numericCitations = (text.match(/\[(?:\d+(?:\s*[-–,]\s*\d+)*)\]/g) || []).length;
      const numbers = (text.match(/[-+]?\d+(?:[.,]\d+)*/g) || []).length;

      const liveEndNote = count(/EN\.CITE/gi);
      const liveZotero = count(/ZOTERO_ITEM/gi);
      const csl = count(/CSL_CITATION|MENDELEY CITATION/gi);
      const bookmarks = count(/bookmarkStart/gi);
      const links = count(/w:hyperlink/gi);
      const equations = count(/m:oMath(?:Para)?/gi);

      results.innerHTML =
        metric("Words", text.trim() ? text.trim().split(/\s+/).length : 0) +
        metric("Scientific numeric tokens", numbers) +
        metric("Plain numbered citations", numericCitations) +
        metric("EndNote field signatures", liveEndNote) +
        metric("Zotero field signatures", liveZotero) +
        metric("CSL/Mendeley signatures", csl) +
        metric("Bookmarks", bookmarks) +
        metric("Hyperlinks", links) +
        metric("Equation nodes", equations);
      results.hidden = false;
    });
    status.textContent = "Completed locally. Full preservation auditing and retargeting remain in the Word Journal Manuscript Converter desktop/core application.";
  } catch (err) {
    status.textContent = `Stopped: ${err.message || err}`;
  }
}
