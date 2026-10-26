let config = { comboBoxes: [] };

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Word) return;
  await loadConfig();
  document.getElementById("insert").onclick = insertSelected;
});

function setStatus(msg) { document.getElementById("status").textContent = msg; }

async function loadConfig() {
  try {
    const res = await fetch("assets/components.json", { cache: "no-store" });
    config = await res.json();
    const sel = document.getElementById("source");
    config.comboBoxes.forEach((c, i) => sel.add(new Option(c.title, i)));
  } catch (e) {
    setStatus("Could not load components.json: " + e.message);
  }
}

async function insertSelected() {
  const def = config.comboBoxes[document.getElementById("source").value];
  if (!def) return;
  try {
    await insertWithApi(def);
    setStatus(`Inserted "${def.title}" combo box.`);
  } catch (apiErr) {
    // Fallback: OOXML works on Word Online where the typed API is unavailable.
    try {
      await insertWithOoxml(def);
      setStatus(`Inserted "${def.title}" combo box (OOXML).`);
    } catch (e) {
      setStatus("Insert failed: " + e.message);
    }
  }
}

// Preferred: Word content control API (WordApi 1.9)
function insertWithApi(def) {
  return Word.run(async (context) => {
    const range = context.document.getSelection();
    const cc = range.insertContentControl("ComboBox");
    cc.title = def.title;
    cc.tag = def.id;
    cc.placeholderText = def.placeholder;
    def.items.forEach((it) => cc.comboBoxContentControl.addListItem(it.text, it.value));
    await context.sync();
  });
}

const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

function insertWithOoxml(def) {
  const items = def.items
    .map((it) => `<w:listItem w:displayText="${esc(it.text)}" w:value="${esc(it.value)}"/>`)
    .join("");
  const xml = `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    </Relationships></pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
      <w:sdt><w:sdtPr><w:alias w:val="${esc(def.title)}"/><w:tag w:val="${esc(def.id)}"/><w:showingPlcHdr/>
        <w:comboBox>${items}</w:comboBox></w:sdtPr>
        <w:sdtContent><w:r><w:t>${esc(def.placeholder)}</w:t></w:r></w:sdtContent></w:sdt>
    </w:p></w:body></w:document></pkg:xmlData>
  </pkg:part>
</pkg:package>`;
  return Word.run(async (context) => {
    context.document.getSelection().insertOoxml(xml, "Replace");
    await context.sync();
  });
}
