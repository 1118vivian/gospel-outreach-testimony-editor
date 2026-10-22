(function () {
  "use strict";
  const fields = ["reportType", "team", "group", "name", "week", "startDate", "endDate", "content"];
  const form = document.getElementById("testimonyForm");
  const preview = document.getElementById("documentPreview");
  const viewport = document.getElementById("previewViewport");
  const storageKey = "gospel-outreach-testimony-editor-draft-v1";
  const values = () => Object.fromEntries(fields.map(key => [key, document.getElementById(key).value.trim()]));
  const dateText = value => value ? value.split("-").map((n, i) => i ? Number(n) : n).join("/") : "";
  const fileName = (data, ext) => [data.name || data.team || "未命名", data.week ? `第${data.week}週` : "", data.reportType === "team" ? "福音開展全隊蒙恩報導" : "福音開展蒙恩見證稿"].filter(Boolean).join("_").replace(/[<>:"/\\|?*\x00-\x1f]/g, "_") + "." + ext;

  function fitPage() {
    const content = preview.querySelector(".page-content");
    content.style.transform = "none";
    const ratio = Math.min(1, 1049 / content.scrollHeight);
    content.style.transform = `scale(${ratio})`;
    const scale = Math.min(1, viewport.parentElement.clientWidth / 794);
    preview.style.transform = `scale(${scale})`;
    viewport.style.width = `${794 * scale}px`;
    viewport.style.height = `${1123 * scale}px`;
  }

  function syncPreview() {
    const data = values();
    const display = { ...data, startDateText: dateText(data.startDate), endDateText: dateText(data.endDate) };
    document.querySelectorAll("[data-preview]").forEach(node => { node.textContent = display[node.dataset.preview] || ""; });
    document.getElementById("fileNamePreview").textContent = fileName(data, "docx");
    const count = Array.from(data.content.replace(/\s/g, "")).length;
    const isTeam = data.reportType === "team";
    document.getElementById("memberChoice").textContent = !isTeam && data.reportType === "member" ? "☑" : "□";
    document.getElementById("teamChoice").textContent = isTeam ? "☑" : "□";
    const target = isTeam ? 500 : 300;
    document.getElementById("wordCount").textContent = `正文共 ${count} 字（不含空白），${isTeam ? "全隊報導" : "隊員見證"}建議約 ${target} 字。`;
    fitPage();
  }

  function saveDraft() {
    try { localStorage.setItem(storageKey, JSON.stringify(values())); document.getElementById("draftStatus").textContent = "草稿已儲存在此瀏覽器。"; }
    catch { document.getElementById("draftStatus").textContent = "此瀏覽器無法儲存草稿，請在關閉前下載檔案。"; }
  }

  async function captureA4(scale) {
    if (!window.html2canvas) throw new Error("版面套件尚未載入，請重新整理後再試。");
    await document.fonts.ready;
    fitPage();
    const clone = preview.cloneNode(true);
    clone.id = "documentPreviewExport";
    clone.classList.add("pdf-export-page");
    clone.style.transform = "none";
    clone.setAttribute("aria-hidden", "true");
    document.body.appendChild(clone);
    try { return await window.html2canvas(clone, { scale, backgroundColor: "#ffffff", useCORS: true, width: 794, height: 1123 }); }
    finally { clone.remove(); }
  }

  async function createDocxBlob() {
    if (!window.docx) throw new Error("Word 套件尚未載入，請重新整理後再試。");
    const d = window.docx;
    const data = values();
    const font = { ascii: "Microsoft JhengHei", hAnsi: "Microsoft JhengHei", eastAsia: "Microsoft JhengHei", cs: "Microsoft JhengHei" };
    const paragraph = (text, options = {}) => new d.Paragraph({
      alignment: options.alignment || d.AlignmentType.LEFT,
      spacing: { before: 0, after: options.after ?? 0, line: options.line || 340, lineRule: d.LineRuleType.EXACT },
      keepNext: false,
      keepLines: false,
      children: text.split(/\r?\n/).map((line, index) => new d.TextRun({
        text: line, break: index ? 1 : undefined, font,
        size: options.size || 26, bold: options.bold || false, snapToGrid: false,
      })),
    });
    const tableWidth = 10706;
    const border = { style: d.BorderStyle.SINGLE, size: 6, color: "C9D1D9" };
    const contentBox = (text, height) => new d.Table({
      width: { size: tableWidth, type: d.WidthType.DXA }, columnWidths: [tableWidth], layout: d.TableLayoutType.FIXED,
      margins: { top: 100, bottom: 100, left: 140, right: 140 },
      borders: { top: border, bottom: border, left: border, right: border },
      rows: [new d.TableRow({ cantSplit: false, height: { value: height, rule: d.HeightRule.ATLEAST }, children: [new d.TableCell({ width: { size: tableWidth, type: d.WidthType.DXA }, verticalAlign: d.VerticalAlign.TOP, children: (text || "").split(/\r?\n/).map(line => paragraph(line)) })] })],
    });
    const doc = new d.Document({ title: "福音開展蒙恩見證稿",
      styles: { default: { document: { run: { font, size: 26, color: "000000" }, paragraph: { spacing: { after: 0 } } } } },
      sections: [{
        properties: { verticalAlign: d.VerticalAlign.TOP, page: { size: { width: 11906, height: 16838 }, margin: { top: 600, right: 600, bottom: 600, left: 600 } } },
        children: [
          paragraph("臺灣福音工作全時間訓練壯年成全班", { bold: true, size: 28, alignment: d.AlignmentType.CENTER, line: 360, after: 40 }),
          paragraph("福音開展蒙恩見證稿", { bold: true, size: 38, alignment: d.AlignmentType.CENTER, line: 460, after: 100 }),
          paragraph(`${data.reportType === "member" ? "☑" : "□"} 隊員一週開展蒙恩見證稿（約300字）`, { size: 22, alignment: d.AlignmentType.CENTER, line: 280 }),
          paragraph(`${data.reportType === "team" ? "☑" : "□"} 全隊一週開展全隊蒙恩報導（約500字）`, { size: 22, alignment: d.AlignmentType.CENTER, line: 280, after: 80 }),
          paragraph(`隊別：${data.team}    第 ${data.group} 組    姓名 ${data.name}    第 ${data.week} 週`, { size: 24, alignment: d.AlignmentType.CENTER, line: 300, after: 20 }),
          paragraph(`${dateText(data.startDate)} 至 ${dateText(data.endDate)}`, { size: 24, alignment: d.AlignmentType.CENTER, line: 300, after: 80 }),
          contentBox(data.content, 9000),
        ],
      }],
    });
    return d.Packer.toBlob(doc);
  }

  function saveBlob(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
  }

  async function buildPdf() {
    if (!window.jspdf?.jsPDF) throw new Error("PDF 套件尚未載入，請重新整理後再試。");
    const canvas = await captureA4(2);
    const pdf = new window.jspdf.jsPDF({ unit: "mm", format: "a4", orientation: "portrait" });
    pdf.addImage(canvas.toDataURL("image/png"), "PNG", 0, 0, 210, 297);
    pdf.save(fileName(values(), "pdf"));
  }

  async function shareDocx() {
    const name = fileName(values(), "docx");
    const blob = await createDocxBlob();
    const file = new File([blob], name, { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
    if (navigator.share && navigator.canShare?.({ files: [file] })) {
      await navigator.share({ title: "福音開展蒙恩見證稿", files: [file] });
    } else {
      saveBlob(blob, name);
      alert("已下載 Word，請到 LINE 或其他通訊軟體選擇檔案傳送。");
    }
  }

  async function runAction(button, action) {
    if (form.elements.endDate.value && form.elements.startDate.value && form.elements.endDate.value < form.elements.startDate.value) { alert("結束日期不可早於開始日期。"); form.elements.endDate.focus(); return; }
    const label = button.textContent;
    const buttons = document.querySelectorAll(".actions button, #clearForm");
    buttons.forEach(b => { b.disabled = true; });
    button.textContent = "產生中…";
    try { await action(); }
    catch (error) { if (error.name !== "AbortError") alert(error.message || "產生檔案時發生錯誤。"); }
    finally { buttons.forEach(b => { b.disabled = false; }); button.textContent = label; }
  }

  form.addEventListener("submit", event => event.preventDefault());
  form.addEventListener("input", () => { syncPreview(); saveDraft(); });
  document.getElementById("clearForm").addEventListener("click", () => {
    if (fields.some(key => document.getElementById(key).value) && !confirm("確定清空所有欄位與此瀏覽器草稿？")) return;
    form.reset();
    try { localStorage.removeItem(storageKey); } catch { /* 無儲存權限時仍可清空表單。 */ }
    document.getElementById("draftStatus").textContent = "已清空。";
    syncPreview();
  });
  document.getElementById("downloadDocx").addEventListener("click", event => runAction(event.currentTarget, async () => saveBlob(await createDocxBlob(), fileName(values(), "docx"))));
  document.getElementById("downloadPdf").addEventListener("click", event => runAction(event.currentTarget, buildPdf));
  document.getElementById("shareDocx").addEventListener("click", event => runAction(event.currentTarget, shareDocx));
  try {
    const draft = JSON.parse(localStorage.getItem(storageKey) || "null");
    if (draft && typeof draft === "object") { fields.forEach(key => { if (typeof draft[key] === "string") document.getElementById(key).value = draft[key]; }); document.getElementById("draftStatus").textContent = "已還原此瀏覽器的草稿。"; }
  } catch { document.getElementById("draftStatus").textContent = "無法讀取草稿，仍可編輯與下載。"; }
  new ResizeObserver(fitPage).observe(viewport.parentElement);
  document.fonts.ready.then(fitPage);
  syncPreview();
})();


