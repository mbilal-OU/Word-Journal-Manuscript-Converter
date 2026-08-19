/* global Office, Word */
let navState = { citations: {}, references: {}, manager: "None detected", refHeadingIndex: -1 };

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("inspect").onclick = inspectDocument;
    document.getElementById("scanNavigator").onclick = scanCitationNavigator;
  }
});

function metric(label, value) {
  return `<div class="metric"><span>${escapeHtml(label)}</span><strong>${escapeHtml(String(value))}</strong></div>`;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function expandNumericGroup(group) {
  const keys = [];
  for (const tokenRaw of group.split(/\s*,\s*/)) {
    const token = tokenRaw.trim();
    const range = token.match(/^(\d+)\s*[-–]\s*(\d+)$/);
    if (range) {
      const a = Number(range[1]);
      const b = Number(range[2]);
      if (a <= b && b - a <= 100) {
        for (let i = a; i <= b; i += 1) keys.push(String(i));
      }
    } else if (/^\d+$/.test(token)) {
      keys.push(String(Number(token)));
    }
  }
  return keys;
}

function citationKeys(text) {
  const keys = [];
  const re = /\[(\d+(?:\s*[-–,]\s*\d+)*)\]/g;
  let match;
  while ((match = re.exec(text)) !== null) {
    keys.push(...expandNumericGroup(match[1]));
  }
  return keys;
}

function detectManager(xml) {
  const count = (pattern) => (xml.match(pattern) || []).length;
  const names = [];
  const endnote = count(/EN\.CITE/gi);
  const zotero = count(/ZOTERO_ITEM/gi);
  const csl = count(/CSL_CITATION|MENDELEY CITATION/gi);
  if (endnote) names.push(`EndNote (${endnote})`);
  if (zotero) names.push(`Zotero (${zotero})`);
  if (csl) names.push(`CSL/Mendeley (${csl})`);
  return names.length ? names.join(", ") : "None detected";
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

      results.innerHTML =
        metric("Words", text.trim() ? text.trim().split(/\s+/).length : 0) +
        metric("Scientific numeric tokens", numbers) +
        metric("Plain numbered citation groups", numericCitations) +
        metric("Citation manager", detectManager(xml)) +
        metric("Bookmarks", count(/bookmarkStart/gi)) +
        metric("Hyperlinks", count(/w:hyperlink/gi)) +
        metric("Equation nodes", count(/m:oMath(?:Para)?/gi));
      results.hidden = false;
    });
    status.textContent = "Completed locally. No document content was transmitted by this add-in operation.";
  } catch (err) {
    status.textContent = `Stopped: ${err.message || err}`;
  }
}

async function scanCitationNavigator() {
  const status = document.getElementById("status");
  const panel = document.getElementById("navigator");
  status.textContent = "Building citation map locally inside Word…";
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      const paragraphs = body.paragraphs;
      paragraphs.load("items/text");
      const ooxml = body.getOoxml();
      await context.sync();

      const texts = paragraphs.items.map((p) => (p.text || "").trim());
      const refHeadingIndex = texts.findIndex((t) => ["references", "bibliography"].includes(t.toLowerCase().replace(/:$/, "")));
      const references = {};
      const citations = {};

      if (refHeadingIndex >= 0) {
        let ordinal = 1;
        for (let i = refHeadingIndex + 1; i < texts.length; i += 1) {
          const text = texts[i];
          if (!text) continue;
          const explicit = text.match(/^\s*(?:\[(\d+)\]|(\d+)[.)])\s*/);
          const key = explicit ? String(Number(explicit[1] || explicit[2])) : String(ordinal);
          if (!references[key]) references[key] = { paragraphIndex: i, text };
          ordinal += 1;
        }

        for (let i = 0; i < refHeadingIndex; i += 1) {
          const keys = citationKeys(texts[i]);
          for (const key of keys) {
            if (!citations[key]) citations[key] = { paragraphIndexes: [], count: 0 };
            citations[key].count += 1;
            if (!citations[key].paragraphIndexes.includes(i)) citations[key].paragraphIndexes.push(i);
          }
        }
      }

      navState = {
        citations,
        references,
        manager: detectManager(ooxml.value || ""),
        refHeadingIndex,
      };
    });

    renderNavigator();
    panel.hidden = false;
    status.textContent = "Citation Navigator ready. Jumping only changes the Word selection; it does not rewrite citation fields.";
  } catch (err) {
    status.textContent = `Stopped: ${err.message || err}`;
  }
}

function renderNavigator() {
  const panel = document.getElementById("navigator");
  const keys = Object.keys(navState.citations).sort((a, b) => Number(a) - Number(b));
  const matched = keys.filter((k) => navState.references[k]).length;
  const unresolved = keys.length - matched;

  let html =
    metric("Manager", navState.manager) +
    metric("Unique citation keys", keys.length) +
    metric("References", Object.keys(navState.references).length) +
    metric("Unresolved", unresolved);

  if (navState.refHeadingIndex < 0) {
    html += '<p class="warning">Could not find a References or Bibliography heading.</p>';
  } else if (!keys.length) {
    html += '<p class="warning">No bracketed numbered citations were detected in the manuscript body.</p>';
  } else {
    html += '<div class="navlist">';
    for (const key of keys) {
      const ref = navState.references[key];
      html += `<div class="navrow"><div><strong>[${escapeHtml(key)}]</strong> <span class="small">${navState.citations[key].count} occurrence(s) · ${ref ? "matched" : "unresolved"}</span></div><div class="rowbuttons"><button class="miniButton" onclick="jumpToCitation('${escapeHtml(key)}')">Citation</button>${ref ? `<button class="miniButton secondaryButton" onclick="jumpToReference('${escapeHtml(key)}')">Reference</button>` : ""}</div></div>`;
    }
    html += "</div>";
  }
  panel.innerHTML = html;
}

async function jumpToParagraph(index) {
  await Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items");
    await context.sync();
    if (index < 0 || index >= paragraphs.items.length) throw new Error("Navigation target is no longer available. Rescan the document.");
    paragraphs.items[index].select();
    await context.sync();
  });
}

async function jumpToCitation(key) {
  const item = navState.citations[key];
  if (!item || !item.paragraphIndexes.length) return;
  await jumpToParagraph(item.paragraphIndexes[0]);
}

async function jumpToReference(key) {
  const item = navState.references[key];
  if (!item) return;
  await jumpToParagraph(item.paragraphIndex);
}

window.jumpToCitation = jumpToCitation;
window.jumpToReference = jumpToReference;
