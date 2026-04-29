/** @odoo-module **/
import {Dialog} from "@web/core/dialog/dialog";
import {Component, xml} from "@odoo/owl";


export class DocumentPreview extends Component {

    static components = {Dialog};

    downloadFile() {
        const url = `/web/content/${this.props.attachmentId}?download=true`;

        const a = document.createElement("a");
        a.href = url;
        a.download = this.props.title;

        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
    }
}

DocumentPreview.template = xml`
<Dialog size="'fullscreen'" technical="false"
    class="'o_doc_preview_dialog'"
    contentClass="'o_full_screen_preview'">

<t t-set-slot="header">
    <div class="header">
        <h5 class="mb-0 text-truncate text-white" t-esc="props.title"/>
    
        <div class="section-2">
            <button class="download-btn btn btn-secondary" t-on-click="downloadFile">
                <i class="fa fa-download"/> Download
            </button>
            <button class="btn-close"
                    t-on-click="props.close"/>
        </div>
    </div>
</t>

<div class="o_preview_container d-flex flex-column justify-content-center fixed bottom-0 m-0 p-0">
    <div id="doc_viewer" class="o_doc_viewer flex-grow-1 overflow-auto">
        Loading preview...
    </div>
</div>
<t t-set-slot="footer">
.
</t>
</Dialog>
`;


document.addEventListener("click", async function (ev) {

    const card = ev.target.closest(".o-mail-AttachmentCard");
    if (!card) return;

    const nameEl = card.querySelector(".text-truncate");
    const filename = nameEl ? nameEl.innerText.trim() : "";

    console.log("file", filename)

    if (![".csv", ".xlsx", ".docx", ".pptx"]
        .some(ext => filename.toLowerCase().endsWith(ext))) {
        return;
    }

    ev.preventDefault();
    ev.stopPropagation();

    const btn = card.querySelector('button[title="Download"]');
    const downloadUrl = btn.getAttribute("data-download-url");

    const attachmentId = new URL(downloadUrl, window.location.origin)
        .pathname
        .split("/")
        .pop();

    const dialogService = owl.Component.env.services.dialog;
    dialogService.add(DocumentPreview, {
        title: filename,
        attachmentId: attachmentId,
    });


    setTimeout(async () => {

        const viewer = document.getElementById("doc_viewer");
        const fileUrl = `/web/content/${attachmentId}`;

        viewer.innerHTML = "Loading preview...";

        viewer.classList.remove("docx-mode", "other-mode");


        if (filename.endsWith(".xlsx") || filename.endsWith(".csv")) {
            viewer.classList.add("other-mode");
            const response = await fetch(`/csv/preview/${attachmentId}`);

            const data = await response.json();

            console.log("data", data);

            viewer.innerHTML = "";
            viewer.style.height = "100%";

            const spreadsheet = x_spreadsheet(viewer, {
                mode: "read",
                showToolbar: false,
                showGrid: true,
                showContextmenu: false
            });
            const sheets = [];

            data.sheets.forEach(sheet => {

                const sheetRows = {
                    len: sheet.rows.length
                };

                sheet.rows.forEach((row, r) => {

                    const cells = {};

                    row.forEach((cell, c) => {
                        cells[c] = {
                            text: cell ? String(cell) : ""
                        };
                    });

                    sheetRows[r] = {cells};

                });

                sheets.push({
                    name: sheet.name,
                    rows: sheetRows
                });

            });

            spreadsheet.loadData(sheets);

            console.log("spreadsheet", spreadsheet);

        } else if (filename.endsWith(".docx")) {

            const viewer = document.getElementById("doc_viewer");
            viewer.classList.add("docx-mode");

            viewer.innerHTML = `<div id="pdf_container"></div>`;

            try {

                const response = await fetch(`/docx/preview/${attachmentId}`);
                const pdfData = await response.arrayBuffer();

                const pdfModule = await import("/ts_office_files_preview/static/lib/pdf.mjs");
                const pdfjsLib = pdfModule;

                pdfjsLib.GlobalWorkerOptions.workerSrc =
                    "/ts_office_files_preview/static/lib/pdf.worker.mjs";

                const loadingTask = pdfjsLib.getDocument({data: pdfData});
                const pdf = await loadingTask.promise;

                const container = document.getElementById("pdf_container");

                for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {

                    const page = await pdf.getPage(pageNum);
                    const viewport = page.getViewport({scale: 1.2});

                    const wrapper = document.createElement("div");
                    wrapper.style.position = "relative";
                    wrapper.style.margin = "20px auto";
                    wrapper.style.width = viewport.width + "px";

                    const canvas = document.createElement("canvas");
                    const ctx = canvas.getContext("2d");

                    canvas.height = viewport.height;
                    canvas.width = viewport.width;

                    wrapper.appendChild(canvas);

                    const pageNumber = document.createElement("div");
                    pageNumber.innerText = `Page ${pageNum} of ${pdf.numPages}`;

                    pageNumber.style.position = "absolute";
                    pageNumber.style.bottom = "8px";
                    pageNumber.style.right = "12px";
                    pageNumber.style.fontSize = "11px";
                    pageNumber.style.color = "#555";
                    pageNumber.style.background = "rgba(255,255,255,0.7)";
                    pageNumber.style.padding = "2px 6px";
                    pageNumber.style.borderRadius = "4px";

                    wrapper.appendChild(pageNumber);

                    container.appendChild(wrapper);

                    await page.render({
                        canvasContext: ctx,
                        viewport: viewport
                    }).promise;
                }
            } catch (err) {

                console.error("DOCX preview failed:", err);

                viewer.innerHTML = `
            <div style="padding:10px">
                Preview not supported.<br>
                Please download the file.
            </div>
        `;
            }
        } else if (filename.endsWith(".pptx")) {

            const viewer = document.getElementById("doc_viewer");
            if (!viewer) return;
            viewer.classList.add("other-mode");
            viewer.innerHTML = `<div id="pdf_container"></div>`;

            try {

                const response = await fetch(`/ppt/preview/${attachmentId}`);
                const pdfData = await response.arrayBuffer();

                // load PDF.js
                const pdfModule = await import("/ts_office_files_preview/static/lib/pdf.mjs");
                const pdfjsLib = pdfModule;

                pdfjsLib.GlobalWorkerOptions.workerSrc =
                    "/ts_office_files_preview/static/lib/pdf.worker.mjs";

                const loadingTask = pdfjsLib.getDocument({data: pdfData});
                const pdf = await loadingTask.promise;

                const container = document.getElementById("pdf_container");

                for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {

                    const page = await pdf.getPage(pageNum);

                    const viewport = page.getViewport({scale: 1});

                    const canvas = document.createElement("canvas");
                    const ctx = canvas.getContext("2d");

                    canvas.height = viewport.height;
                    canvas.width = viewport.width;

                    canvas.style.display = "block";
                    canvas.style.margin = "10px auto";
                    container.appendChild(canvas);


                    const pageNumber = document.createElement("div");
                    pageNumber.innerText = `Page ${pageNum} of ${pdf.numPages}`;

                    pageNumber.style.textAlign = "right";
                    pageNumber.style.fontSize = "10px";
                    pageNumber.style.color = "#555";
                    pageNumber.style.marginBottom = "10px";

                    container.appendChild(pageNumber);

                    await page.render({
                        canvasContext: ctx,
                        viewport: viewport
                    }).promise;
                }

            } catch (err) {

                console.error("PDF preview failed:", err);

                viewer.innerHTML = `
                <div style="padding:10px">
                    Preview not supported.<br>
                    Please download the file.
                </div>
            `;
            }
        }
    }, 60);
});


document.addEventListener("mouseover", function (ev) {

    const card = ev.target.closest(".o-mail-AttachmentCard");
    if (!card) return;

    if (!card.dataset.cursorSet) {
        card.style.cursor = "zoom-in";
        card.dataset.cursorSet = "true";
    }

});
//
// async function splitDocxPages(container) {
//     const PAGE_HEIGHT = 939;
//     const article = container.querySelector("article");
//     if (!article) return;
//     const images = article.querySelectorAll('img');
//     await Promise.all([...images].map(img => {
//         if (img.complete) return Promise.resolve();
//         return new Promise(resolve => {
//             img.onload = resolve;
//             img.onerror = resolve;
//         });
//     }));
//
//     const nodes = [...article.children];
//     let page = container.querySelector("section.docx");
//     let currentHeight = 0;
//     const pages = [page];
//
//     nodes.forEach(node => {
//         const nodeHeight = node.getBoundingClientRect().height; // Use more precise measurement
//
//         if (currentHeight + nodeHeight > PAGE_HEIGHT && currentHeight > 0) {
//             const newPage = document.createElement("section");
//             newPage.className = "docx";
//             const newArticle = document.createElement("article");
//             newPage.appendChild(newArticle);
//             container.appendChild(newPage);
//
//             page = newPage;
//             pages.push(page);
//             currentHeight = 0;
//         }
//
//         page.querySelector("article").appendChild(node);
//         currentHeight += nodeHeight;
//     });
//
//     pages.forEach((p, index) => {
//         const existing = p.querySelector(".docx-page-number");
//         if (existing) existing.remove();
//
//         const footer = document.createElement("div");
//         footer.className = "docx-page-number";
//         footer.textContent = `Page ${index + 1} of ${pages.length}`;
//         p.appendChild(footer);
//     });
// }