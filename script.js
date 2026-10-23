
const $ = id => document.getElementById(id);

const templates = {
  winner: { image: null, data: null, positions: {} },
  participant: { image: null, data: null, positions: {} }
};

let participants = [];
let generatedFiles = [];
let currentType = "winner";
let placementMode = false;

$("readBtn").addEventListener("click", readExcel);
$("templateType").addEventListener("change", switchTemplate);
$("templateFile").addEventListener("change", loadTemplate);
$("placeModeBtn").addEventListener("click", beginPlacement);
$("templateCanvas").addEventListener("click", placeField);
$("fieldSelect").addEventListener("change", loadFieldStyle);
$("fontSize").addEventListener("change", saveFieldStyle);
$("fontSelect").addEventListener("change", saveFieldStyle);
$("boldText").addEventListener("change", saveFieldStyle);
$("textColor").addEventListener("change", saveFieldStyle);
$("clearFieldBtn").addEventListener("click", clearField);
$("clearAllBtn").addEventListener("click", clearAll);
$("previewBtn").addEventListener("click", previewFirst);
$("generateBtn").addEventListener("click", generatePDFs);
$("zipBtn").addEventListener("click", downloadZip);
window.addEventListener("resize", renderMarkers);

function status(id, message, error = false) {
  $(id).textContent = message;
  $(id).classList.toggle("error", error);
}

function currentTemplate() {
  return templates[currentType];
}

function switchTemplate() {
  currentType = $("templateType").value;
  placementMode = false;
  $("templateFile").value = "";
  const t = currentTemplate();

  if (t.image) {
    drawTemplate();
  } else {
    $("templateCanvas").style.display = "none";
    $("emptyMessage").style.display = "block";
    $("markers").replaceChildren();
  }

  status("templateStatus", `Selected ${currentType} template.`);
  renderPositionList();
  loadFieldStyle();
}

function isWinner(position) {
  const p = String(position).trim().toLowerCase();
  return [
    "1", "1st", "first", "first place",
    "2", "2nd", "second", "second place",
    "3", "3rd", "third", "third place"
  ].includes(p);
}

function formatDate(value) {
  if (value instanceof Date && !isNaN(value)) {
    return `${String(value.getDate()).padStart(2, "0")}-${String(value.getMonth()+1).padStart(2, "0")}-${value.getFullYear()}`;
  }
  return String(value ?? "").trim();
}

async function readExcel() {
  participants = [];
  generatedFiles = [];
  $("participantsTable").replaceChildren();
  $("downloads").replaceChildren();
  $("zipBtn").disabled = true;

  const file = $("excelFile").files[0];
  if (!file) return status("excelStatus", "Choose an Excel file first.", true);

  try {
    const wb = XLSX.read(await file.arrayBuffer(), {
      type: "array", cellDates: true
    });
    const rows = XLSX.utils.sheet_to_json(
      wb.Sheets[wb.SheetNames[0]], { defval: "" }
    );

    if (!rows.length) throw new Error("The spreadsheet is empty.");
    if (rows.length > 5) throw new Error("Maximum 5 participants allowed.");

    const headers = Object.keys(rows[0]);
    const find = name => headers.find(
      h => h.trim().toLowerCase() === name.toLowerCase()
    );

    const columns = ["Name", "Event", "Date", "Email", "Position"];
    const map = Object.fromEntries(columns.map(c => [c, find(c)]));
    const missing = columns.filter(c => !map[c]);
    if (missing.length) throw new Error("Missing columns: " + missing.join(", "));

    participants = rows.map((row, index) => {
      const p = {};
      columns.forEach(c => {
        p[c] = c === "Date"
          ? formatDate(row[map[c]])
          : String(row[map[c]] ?? "").trim();
      });

      if (columns.some(c => !p[c])) {
        throw new Error(`Row ${index + 2}: fill all five columns.`);
      }
      if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(p.Email)) {
        throw new Error(`Invalid email at row ${index + 2}.`);
      }

      p.type = isWinner(p.Position) ? "winner" : "participant";
      return p;
    });

    const table = document.createElement("table");
    const head = document.createElement("tr");
    [...columns, "Certificate"].forEach(c => {
      const th = document.createElement("th");
      th.textContent = c;
      head.appendChild(th);
    });
    const thead = document.createElement("thead");
    thead.appendChild(head);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");
    participants.forEach(p => {
      const tr = document.createElement("tr");
      [...columns.map(c => p[c]), p.type].forEach(value => {
        const td = document.createElement("td");
        td.textContent = value;
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    $("participantsTable").appendChild(table);
    status("excelStatus", `Loaded ${participants.length} participant(s).`);
  } catch (error) {
    participants = [];
    status("excelStatus", error.message, true);
  }
}

function loadTemplate() {
  const file = $("templateFile").files[0];
  if (!file) return;

  if (!["image/jpeg", "image/png"].includes(file.type)) {
    return status("templateStatus", "Upload a JPG, JPEG, or PNG image.", true);
  }

  const reader = new FileReader();
  reader.onload = () => {
    const img = new Image();
    img.onload = () => {
      currentTemplate().image = img;
      currentTemplate().data = reader.result;
      drawTemplate();
      status("templateStatus", `Loaded ${file.name} (${img.naturalWidth} × ${img.naturalHeight}).`);
    };
    img.onerror = () => status("templateStatus", "Unable to load this image.", true);
    img.src = reader.result;
  };
  reader.readAsDataURL(file);
}

function drawTemplate() {
  const img = currentTemplate().image;
  if (!img) return;

  const canvas = $("templateCanvas");
  canvas.width = img.naturalWidth;
  canvas.height = img.naturalHeight;
  canvas.getContext("2d").drawImage(img, 0, 0);
  canvas.style.display = "block";
  $("emptyMessage").style.display = "none";
  renderMarkers();
}

function beginPlacement() {
  if (!currentTemplate().image) {
    return status("editorStatus", "Upload a template image first.", true);
  }
  placementMode = true;
  status("editorStatus", `Now click where ${$("fieldSelect").value} should start.`);
}

function placeField(event) {
  if (!placementMode) return;

  const canvas = $("templateCanvas");
  const rect = canvas.getBoundingClientRect();
  const field = $("fieldSelect").value;

  currentTemplate().positions[field] = {
    x: (event.clientX - rect.left) / rect.width,
    y: (event.clientY - rect.top) / rect.height,
    size: fontSize(),
    font: $("fontSelect").value,
    bold: $("boldText").checked,
    color: $("textColor").value
  };

  placementMode = false;
  status("editorStatus", `${field} position saved for ${currentType} template.`);
  renderMarkers();
  renderPositionList();
}

function fontSize() {
  return Math.max(6, Math.min(72, Number($("fontSize").value) || 24));
}

function saveFieldStyle() {
  const p = currentTemplate().positions[$("fieldSelect").value];
  if (!p) return;
  p.size = fontSize();
  p.font = $("fontSelect").value;
  p.bold = $("boldText").checked;
  p.color = $("textColor").value;
  renderPositionList();
}

function loadFieldStyle() {
  const p = currentTemplate().positions[$("fieldSelect").value];
  if (!p) return;
  $("fontSize").value = p.size;
  $("fontSelect").value = p.font;
  $("boldText").checked = p.bold;
  $("textColor").value = p.color;
}

function clearField() {
  delete currentTemplate().positions[$("fieldSelect").value];
  renderMarkers();
  renderPositionList();
  status("editorStatus", "Selected field position cleared.");
}

function clearAll() {
  currentTemplate().positions = {};
  renderMarkers();
  renderPositionList();
  status("editorStatus", "All field positions cleared for this template.");
}

function renderMarkers() {
  const holder = $("markers");
  holder.replaceChildren();

  Object.entries(currentTemplate().positions).forEach(([field, p]) => {
    const marker = document.createElement("div");
    marker.className = "marker";
    marker.style.left = `${p.x * 100}%`;
    marker.style.top = `${p.y * 100}%`;
    const label = document.createElement("span");
    label.textContent = field;
    marker.appendChild(label);
    holder.appendChild(marker);
  });
}

function renderPositionList() {
  const holder = $("positionList");
  holder.replaceChildren();
  const title = document.createElement("strong");
  title.textContent = `${currentType} positions`;
  holder.appendChild(title);

  ["Name", "Event", "Date", "Position"].forEach(field => {
    const line = document.createElement("div");
    const p = currentTemplate().positions[field];
    line.textContent = p
      ? `${field}: placed · ${p.size} pt${p.bold ? " · bold" : ""}`
      : `${field}: not placed`;
    holder.appendChild(line);
  });
}

function previewFirst() {
  if (!participants.length) {
    return status("generateStatus", "Read the spreadsheet first.", true);
  }
  const p = participants[0];
  const t = templates[p.type];

  if (!t.image) {
    return status("generateStatus", `Upload the ${p.type} template first.`, true);
  }

  makePDF(p, t).then(blob => {
    const frame = document.createElement("iframe");
    frame.src = URL.createObjectURL(blob);
    $("pdfPreview").replaceChildren(frame);
  }).catch(error => status("generateStatus", error.message, true));
}

function makePDF(person, template) {
  const { jsPDF } = window.jspdf;
  const img = template.image;
  const pdf = new jsPDF({
    orientation: img.naturalWidth >= img.naturalHeight ? "landscape" : "portrait",
    unit: "pt",
    format: "a4"
  });

  const pageW = pdf.internal.pageSize.getWidth();
  const pageH = pdf.internal.pageSize.getHeight();
  const scale = Math.min(pageW / img.naturalWidth, pageH / img.naturalHeight);
  const drawW = img.naturalWidth * scale;
  const drawH = img.naturalHeight * scale;
  const ox = (pageW - drawW) / 2;
  const oy = (pageH - drawH) / 2;

  pdf.addImage(
    template.data,
    template.data.startsWith("data:image/png") ? "PNG" : "JPEG",
    ox, oy, drawW, drawH
  );

  const values = {
    Name: person.Name,
    Event: person.Event,
    Date: person.Date,
    Position: person.Position
  };

  Object.entries(template.positions).forEach(([field, pos]) => {
    if (values[field] === undefined) return;
    pdf.setFont(pos.font, pos.bold ? "bold" : "normal");
    pdf.setFontSize(pos.size);
    pdf.setTextColor(pos.color);
    const x = ox + pos.x * drawW;
    const y = oy + pos.y * drawH + pos.size * 0.35;
    pdf.text(String(values[field]), x, y);
  });

  return pdf.output("blob");
}

async function generatePDFs() {
  if (!participants.length) {
    return status("generateStatus", "Read the spreadsheet first.", true);
  }

  generatedFiles = [];
  $("downloads").replaceChildren();
  $("zipBtn").disabled = true;
  status("generateStatus", "Generating certificates...");

  try {
    for (let i = 0; i < participants.length; i++) {
      const p = participants[i];
      const t = templates[p.type];

      if (!t.image) throw new Error(`Upload the ${p.type} template first.`);
      const missing = ["Name", "Event", "Date"].filter(f => !t.positions[f]);
      if (missing.length) {
        throw new Error(`${p.type} template: place ${missing.join(", ")} first.`);
      }

      const safeName = p.Name.replace(/[<>:"/\\|?*\x00-\x1F]/g, "_");
      generatedFiles.push({
        person: p,
        blob: makePDF(p, t),
        filename: `${String(i + 1).padStart(2, "0")}_${safeName}_certificate.pdf`
      });
    }

    generatedFiles.forEach(file => {
      const link = document.createElement("a");
      link.className = "download-link";
      link.href = URL.createObjectURL(file.blob);
      link.download = file.filename;
      link.textContent = `${file.person.Name} — ${file.person.type} certificate · Download PDF`;
      $("downloads").appendChild(link);
    });

    $("zipBtn").disabled = false;
    status("generateStatus", `Generated ${generatedFiles.length} PDF(s).`);
  } catch (error) {
    status("generateStatus", error.message, true);
  }
}

async function downloadZip() {
  if (!generatedFiles.length) return;

  $("zipBtn").disabled = true;
  try {
    const zip = new JSZip();
    generatedFiles.forEach(file => zip.file(file.filename, file.blob));
    const blob = await zip.generateAsync({ type: "blob" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "certificates.zip";
    link.click();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  } catch (error) {
    status("generateStatus", "Could not create ZIP: " + error.message, true);
  } finally {
    $("zipBtn").disabled = false;
  }
}
