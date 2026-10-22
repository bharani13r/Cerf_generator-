
const $ = id => document.getElementById(id);

let participants = [];
let templateBytes = null;
let templateType = null;
let templateImage = null;
let pdfPage = null;
let pdfViewport = null;
let pdfDocument = null;
let positions = {};
let generatedFiles = [];
let previewScale = 1;

if (window.pdfjsLib) {
    pdfjsLib.GlobalWorkerOptions.workerSrc =
        "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
}

// Read Excel file
$("readExcel").addEventListener("click", async () => {
    const file = $("excelFile").files[0];

    if (!file) {
        $("excelStatus").textContent = "Please select an Excel file.";
        return;
    }

    try {
        const buffer = await file.arrayBuffer();
        const workbook = XLSX.read(buffer, {
            type: "array",
            cellDates: true
        });

        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, {
            defval: ""
        });

        if (!rows.length) {
            throw new Error("The Excel sheet is empty.");
        }

        const headers = Object.keys(rows[0]);

        const findHeader = name =>
            headers.find(h => h.trim().toLowerCase() === name.toLowerCase());

        const nameKey = findHeader("Name");
        const eventKey = findHeader("Event");
        const dateKey = findHeader("Date");
        const emailKey = findHeader("Email");

        if (!nameKey || !eventKey || !dateKey || !emailKey) {
            throw new Error(
                "Excel must contain Name, Event, Date and Email columns."
            );
        }

        if (rows.length > 5) {
            throw new Error(
                "Maximum limit is 5 participants. Reduce the Excel rows."
            );
        }

        participants = rows.map((row, index) => {
            const name = String(row[nameKey]).trim();
            const event = String(row[eventKey]).trim();
            const date = formatDate(row[dateKey]);
            const email = String(row[emailKey]).trim();

            if (!name || !event || !date || !email) {
                throw new Error(
                    `Missing details in participant row ${index + 2}.`
                );
            }

            if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
                throw new Error(
                    `Invalid email address in row ${index + 2}.`
                );
            }

            return {
                Name: name,
                Event: event,
                Date: date,
                Email: email
            };
        });

        $("participantList").innerHTML =
            "<h3>Participants</h3>" +
            participants.map((p, i) =>
                `<p>${i + 1}. ${escapeHTML(p.Name)} — ${escapeHTML(p.Email)}</p>`
            ).join("");

        $("excelStatus").textContent =
            `Successfully read ${participants.length} participant(s).`;

    } catch (error) {
        participants = [];
        $("participantList").innerHTML = "";
        $("excelStatus").textContent = error.message;
    }
});

function formatDate(value) {
    if (value instanceof Date && !isNaN(value)) {
        return value.toLocaleDateString("en-GB");
    }

    return String(value).trim();
}

function escapeHTML(value) {
    return String(value).replace(/[&<>"']/g, c => ({
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#39;"
    })[c]);
}

// Upload and preview template
$("templateFile").addEventListener("change", async () => {
    const file = $("templateFile").files[0];

    if (!file) return;

    try {
        templateBytes = await file.arrayBuffer();
        templateType = null;
        templateImage = null;
        pdfDocument = null;
        pdfPage = null;
        pdfViewport = null;
        positions = {};
        generatedFiles = [];

        $("downloadLinks").innerHTML = "";
        $("downloadZipBtn").disabled = true;
        $("positionStatus").textContent = "";

        $("pdfPreview").style.display = "none";
        $("imagePreview").style.display = "none";
        $("previewMessage").style.display = "none";

        if (
            file.type === "application/pdf" ||
            file.name.toLowerCase().endsWith(".pdf")
        ) {
            templateType = "pdf";

            const loadingTask = pdfjsLib.getDocument({
                data: new Uint8Array(templateBytes)
            });

            pdfDocument = await loadingTask.promise;
            pdfPage = await pdfDocument.getPage(1);

            const baseViewport = pdfPage.getViewport({
                scale: 1
            });

            previewScale = Math.min(800 / baseViewport.width, 1);

            pdfViewport = pdfPage.getViewport({
                scale: previewScale
            });

            const canvas = $("pdfPreview");

            canvas.width = Math.ceil(pdfViewport.width);
            canvas.height = Math.ceil(pdfViewport.height);
            canvas.style.display = "block";

            await pdfPage.render({
                canvasContext: canvas.getContext("2d"),
                viewport: pdfViewport
            }).promise;

        } else if (file.type.startsWith("image/")) {
            templateType = "image";

            templateImage = new Image();
            templateImage.src = URL.createObjectURL(file);

            await new Promise((resolve, reject) => {
                templateImage.onload = resolve;
                templateImage.onerror = reject;
            });

            const img = $("imagePreview");

            img.src = templateImage.src;
            img.style.display = "block";

            previewScale = Math.min(
                800 / templateImage.naturalWidth,
                1
            );

            img.style.width =
                `${templateImage.naturalWidth * previewScale}px`;

        } else {
            throw new Error("Choose a PDF, JPG or PNG file.");
        }

        $("templateStatus").textContent =
            "Template loaded successfully.";

    } catch (error) {
        templateBytes = null;
        templateType = null;

        $("templateStatus").textContent =
            "Could not load template: " + error.message;

        $("previewMessage").style.display = "block";
    }
});

// Click preview to position selected field
$("pdfPreview").addEventListener("click", event =>
    setPosition(event, "pdf")
);

$("imagePreview").addEventListener("click", event =>
    setPosition(event, "image")
);

function setPosition(event, type) {
    if (!templateType || type !== templateType) return;

    const rect = event.currentTarget.getBoundingClientRect();

    const x = (event.clientX - rect.left) / rect.width;
    const y = (event.clientY - rect.top) / rect.height;

    const field = $("fieldSelect").value;

    positions[field] = { x, y };

    $("positionStatus").textContent =
        `${field} position saved. Select another field and click its position.`;
}

// Generate PDFs
$("generateBtn").addEventListener("click", async () => {
    $("generateStatus").textContent = "";
    $("downloadLinks").innerHTML = "";
    $("downloadZipBtn").disabled = true;
    generatedFiles = [];

    if (!participants.length) {
        $("generateStatus").textContent =
            "Please load participant details first.";
        return;
    }

    if (!templateBytes || !templateType) {
        $("generateStatus").textContent =
            "Please upload a certificate template.";
        return;
    }

    for (const field of ["Name", "Event", "Date"]) {
        if (!positions[field]) {
            $("generateStatus").textContent =
                `Please select and position ${field} on the template first.`;
            return;
        }
    }

    $("generateBtn").disabled = true;

    try {
        for (const person of participants) {
            const bytes = await createCertificate(person);

            const filename =
                safeFilename(person.Name) + "_certificate.pdf";

            generatedFiles.push({
                filename,
                bytes
            });
        }

        $("downloadLinks").innerHTML = generatedFiles.map((file, i) => {
            const url = URL.createObjectURL(
                new Blob([file.bytes], {
                    type: "application/pdf"
                })
            );

            return `<a class="download-item" href="${url}" download="${escapeHTML(file.filename)}">Download certificate ${i + 1}: ${escapeHTML(file.filename)}</a>`;
        }).join("");

        $("downloadZipBtn").disabled = false;

        $("generateStatus").textContent =
            `Successfully generated ${generatedFiles.length} certificate(s).`;

    } catch (error) {
        $("generateStatus").textContent =
            "Generation failed: " + error.message;

    } finally {
        $("generateBtn").disabled = false;
    }
});

// Create certificate with bold participant name
async function createCertificate(person) {
    const { PDFDocument, StandardFonts, rgb } = PDFLib;

    const pdf = templateType === "pdf"
        ? await PDFDocument.load(templateBytes)
        : await PDFDocument.create();

    let page;

    if (templateType === "pdf") {
        // Use the first page of the uploaded PDF.
        page = pdf.getPages()[0];

    } else {
        // Convert image template into a PDF page.
        const image = templateImage;

        const imageBytes = new Uint8Array(
            await (await fetch(image.src)).arrayBuffer()
        );

        let embeddedImage;

        if ($("templateFile").files[0].type === "image/png") {
            embeddedImage = await pdf.embedPng(imageBytes);
        } else {
            embeddedImage = await pdf.embedJpg(imageBytes);
        }

        page = pdf.addPage([
            image.naturalWidth,
            image.naturalHeight
        ]);

        page.drawImage(embeddedImage, {
            x: 0,
            y: 0,
            width: image.naturalWidth,
            height: image.naturalHeight
        });
    }

    // Load regular and bold fonts separately.
    const regularFont = await pdf.embedFont(
        StandardFonts.Helvetica
    );

    const boldFont = await pdf.embedFont(
        StandardFonts.HelveticaBold
    );

    const { width, height } = page.getSize();

    const fontSize = Number($("fontSize").value) || 24;

    for (const field of ["Name", "Event", "Date"]) {
        const pos = positions[field];
        const text = person[field];

        // Name uses bold font; Event and Date use regular font.
        const font = field === "Name"
            ? boldFont
            : regularFont;

        const textWidth = font.widthOfTextAtSize(
            text,
            fontSize
        );

        // Click position represents the approximate horizontal center.
        const x = Math.max(
            0,
            Math.min(
                width - textWidth,
                pos.x * width - textWidth / 2
            )
        );

        const y = Math.max(
            0,
            Math.min(
                height - fontSize,
                height - pos.y * height - fontSize / 2
            )
        );

        page.drawText(text, {
            x,
            y,
            size: fontSize,
            font,
            color: rgb(0, 0, 0)
        });
    }

    return await pdf.save();
}

function safeFilename(name) {
    return String(name)
        .replace(/[<>:"/\\|?*\x00-\x1F]/g, "_")
        .trim() || "participant";
}

// Download all PDFs as ZIP
$("downloadZipBtn").addEventListener("click", async () => {
    if (!generatedFiles.length) return;

    const zip = new JSZip();

    generatedFiles.forEach(file => {
        zip.file(file.filename, file.bytes);
    });

    const blob = await zip.generateAsync({
        type: "blob"
    });

    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");

    link.href = url;
    link.download = "certificates.zip";
    link.click();

    URL.revokeObjectURL(url);
});
