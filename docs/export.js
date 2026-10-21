// KorkedAgri — export de la saison sélectionnée en Excel (.xlsx) ou PDF.
// Fonctionne hors ligne : l'Excel est généré ici (zip + XML), le PDF via jsPDF embarqué.
(function () {
  const NAT = { charge: "Charge", benef: "Bénéfice" };

  function seasonRows() {
    return ops
      .filter(o => seasonOf(o.date) === season)
      .sort((a, b) => (a.date || "").localeCompare(b.date || "") || (a.at || 0) - (b.at || 0));
  }
  function totals(rows) {
    const ben = rows.filter(o => o.mode === "benef").reduce((s, o) => s + o.amount, 0);
    const chg = rows.filter(o => o.mode === "charge").reduce((s, o) => s + o.amount, 0);
    const byType = {};
    rows.forEach(o => {
      const k = o.mode + "|" + o.kind;
      byType[k] = (byType[k] || 0) + o.amount;
    });
    const types = Object.entries(byType)
      .map(([k, v]) => ({ mode: k.split("|")[0], kind: k.split("|").slice(1).join("|"), total: v }))
      .sort((a, b) => (a.mode === b.mode ? b.total - a.total : a.mode === "benef" ? -1 : 1));
    return { ben, chg, net: ben - chg, types };
  }
  const r3 = n => Math.round(n * 1000) / 1000;
  const frDate = d => (d ? d.slice(8, 10) + "/" + d.slice(5, 7) + "/" + d.slice(0, 4) : "");
  const fileBase = () => "KorkedAgri_Saison_" + season + "-" + (season + 1);

  // ---------- ZIP (stockage sans compression) ----------
  const CRC = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();
  function crc32(b) {
    let c = 0xffffffff;
    for (let i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  }
  function zip(files) {
    const enc = new TextEncoder();
    const parts = [], central = [];
    let offset = 0;
    for (const f of files) {
      const name = enc.encode(f.name), data = enc.encode(f.data), crc = crc32(data);
      const h = new DataView(new ArrayBuffer(30));
      h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true);
      h.setUint16(8, 0, true); h.setUint16(10, 0, true); h.setUint16(12, 0x5521, true);
      h.setUint32(14, crc, true); h.setUint32(18, data.length, true); h.setUint32(22, data.length, true);
      h.setUint16(26, name.length, true); h.setUint16(28, 0, true);
      parts.push(new Uint8Array(h.buffer), name, data);
      const c = new DataView(new ArrayBuffer(46));
      c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true);
      c.setUint16(8, 0x0800, true); c.setUint16(10, 0, true); c.setUint16(12, 0, true); c.setUint16(14, 0x5521, true);
      c.setUint32(16, crc, true); c.setUint32(20, data.length, true); c.setUint32(24, data.length, true);
      c.setUint16(28, name.length, true); c.setUint32(42, offset, true);
      central.push(new Uint8Array(c.buffer), name);
      offset += 30 + name.length + data.length;
    }
    const cSize = central.reduce((s, p) => s + p.length, 0);
    const e = new DataView(new ArrayBuffer(22));
    e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.length, true); e.setUint16(10, files.length, true);
    e.setUint32(12, cSize, true); e.setUint32(16, offset, true);
    return new Blob([...parts, ...central, new Uint8Array(e.buffer)], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
  }

  // ---------- XLSX ----------
  const xesc = s => String(s ?? "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const colL = i => String.fromCharCode(65 + i);
  const serial = d => (Date.UTC(+d.slice(0, 4), +d.slice(5, 7) - 1, +d.slice(8, 10)) - Date.UTC(1899, 11, 30)) / 86400000;
  // styles : 0 normal, 1 en-tête gras, 2 date, 3 montant, 4 montant gras, 5 titre
  function cell(r, c, v, s) {
    const ref = colL(c) + r, st = s ? ` s="${s}"` : "";
    if (v === null || v === undefined || v === "") return "";
    if (typeof v === "number") return `<c r="${ref}"${st}><v>${v}</v></c>`;
    return `<c r="${ref}"${st} t="inlineStr"><is><t xml:space="preserve">${xesc(v)}</t></is></c>`;
  }
  function sheet(rows, widths) {
    const cols = "<cols>" + widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join("") + "</cols>";
    const data = rows.map((row, i) => `<row r="${i + 1}">` + row.map((c, j) => (c ? cell(i + 1, j, c[0], c[1]) : "")).join("") + "</row>").join("");
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${cols}<sheetData>${data}</sheetData></worksheet>`;
  }
  function buildXlsx() {
    const rows = seasonRows(), t = totals(rows), label = seasonLabel(season);
    const s1 = [
      [["KorkedAgri — " + label, 5]],
      [["Du 01/10/" + season + " au 30/09/" + (season + 1)]],
      [],
      [["Date", 1], ["Nature", 1], ["Type", 1], ["Commentaire", 1], ["Montant (DT)", 1]],
      ...rows.map(o => [
        o.date ? [serial(o.date), 2] : null,
        [NAT[o.mode] || o.mode],
        [o.kind],
        [o.com || ""],
        [o.mode === "charge" ? -r3(o.amount) : r3(o.amount), 3],
      ]),
    ];
    const s2 = [
      [["KorkedAgri — Résumé " + label, 5]],
      [],
      [["Total bénéfices", 1], [r3(t.ben), 4]],
      [["Total charges", 1], [r3(t.chg), 4]],
      [["Résultat net", 1], [r3(t.net), 4]],
      [],
      [["Nature", 1], ["Type", 1], ["Total (DT)", 1]],
      ...t.types.map(x => [[NAT[x.mode]], [x.kind], [r3(x.total), 3]]),
    ];
    const files = [
      { name: "[Content_Types].xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>` },
      { name: "_rels/.rels", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>` },
      { name: "xl/workbook.xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Opérations" sheetId="1" r:id="rId1"/><sheet name="Résumé" sheetId="2" r:id="rId2"/></sheets></workbook>` },
      { name: "xl/_rels/workbook.xml.rels", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>` },
      { name: "xl/styles.xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.000"/></numFmts><fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="14"/><color rgb="FF2F5D3A"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFDEBD6"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="6"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="164" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/><xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>` },
      { name: "xl/worksheets/sheet1.xml", data: sheet(s1, [12, 11, 22, 34, 15]) },
      { name: "xl/worksheets/sheet2.xml", data: sheet(s2, [20, 24, 15]) },
    ];
    return { blob: zip(files), name: fileBase() + ".xlsx" };
  }

  // ---------- PDF ----------
  const pdfTxt = s => String(s ?? "").replace(/→/g, "-").replace(/[−]/g, "-");
  const money = n => (n < 0 ? "-" : "") + Math.abs(r3(n)).toLocaleString("fr-FR", { minimumFractionDigits: 3, maximumFractionDigits: 3 }).replace(/ | /g, " ") + " DT";
  function buildPdf() {
    const { jsPDF } = window.jspdf;
    const rows = seasonRows(), t = totals(rows), label = seasonLabel(season);
    const doc = new jsPDF({ unit: "mm", format: "a4" });
    const green = [47, 93, 58], orange = [232, 118, 15];
    doc.setFillColor(...orange); doc.roundedRect(14, 12, 10, 10, 2, 2, "F");
    doc.setFillColor(255, 255, 255); doc.circle(19, 17.4, 3.4, "F");
    doc.setFont("helvetica", "bold"); doc.setFontSize(18); doc.setTextColor(29, 42, 31);
    doc.text("KorkedAgri", 27, 19.5);
    doc.setFontSize(13); doc.setTextColor(...green); doc.text(label, 196, 16, { align: "right" });
    doc.setFont("helvetica", "normal"); doc.setFontSize(9); doc.setTextColor(95, 109, 95);
    doc.text("Du 01/10/" + season + " au 30/09/" + (season + 1), 196, 21, { align: "right" });

    doc.autoTable({
      startY: 28,
      theme: "plain",
      head: [["Bénéfices", "Charges", "Résultat net"]],
      body: [[money(t.ben), money(t.chg), money(t.net)]],
      headStyles: { fillColor: green, textColor: 255, fontStyle: "bold", fontSize: 9 },
      bodyStyles: { fillColor: [234, 241, 233], fontStyle: "bold", fontSize: 12, textColor: [29, 42, 31] },
      didParseCell: d => { if (d.section === "body" && d.column.index === 2 && t.net < 0) d.cell.styles.textColor = [194, 65, 45]; },
      margin: { left: 14, right: 14 },
    });

    doc.setFont("helvetica", "bold"); doc.setFontSize(11); doc.setTextColor(29, 42, 31);
    doc.text("Totaux par type", 14, doc.lastAutoTable.finalY + 9);
    doc.autoTable({
      startY: doc.lastAutoTable.finalY + 12,
      head: [["Nature", "Type", "Total"]],
      body: t.types.map(x => [NAT[x.mode], pdfTxt(x.kind), money(x.total)]),
      headStyles: { fillColor: [253, 235, 214], textColor: [29, 42, 31], fontStyle: "bold" },
      columnStyles: { 2: { halign: "right" } },
      styles: { fontSize: 9 },
      margin: { left: 14, right: 14 },
    });

    doc.text("Détail des opérations", 14, doc.lastAutoTable.finalY + 9);
    doc.autoTable({
      startY: doc.lastAutoTable.finalY + 12,
      head: [["Date", "Nature", "Type", "Commentaire", "Montant"]],
      body: rows.map(o => [frDate(o.date), NAT[o.mode], pdfTxt(o.kind), pdfTxt(o.com || ""), (o.mode === "charge" ? "-" : "+") + money(o.amount)]),
      headStyles: { fillColor: green, textColor: 255, fontStyle: "bold" },
      columnStyles: { 0: { cellWidth: 22 }, 1: { cellWidth: 20 }, 4: { halign: "right", cellWidth: 30 } },
      styles: { fontSize: 8.5, overflow: "linebreak" },
      didParseCell: d => {
        if (d.section === "body" && d.column.index === 4) d.cell.styles.textColor = d.cell.raw.startsWith("-") ? [194, 65, 45] : [47, 122, 69];
      },
      margin: { left: 14, right: 14 },
    });
    const n = doc.getNumberOfPages();
    for (let i = 1; i <= n; i++) {
      doc.setPage(i); doc.setFont("helvetica", "normal"); doc.setFontSize(8); doc.setTextColor(140);
      doc.text("KorkedAgri - " + label + " - exporté le " + frDate(today()), 14, 290);
      doc.text(i + " / " + n, 196, 290, { align: "right" });
    }
    return { blob: doc.output("blob"), name: fileBase() + ".pdf" };
  }

  // ---------- Livraison du fichier ----------
  function toB64(blob) {
    return new Promise((res, rej) => {
      const r = new FileReader();
      r.onload = () => res(String(r.result).split(",")[1]);
      r.onerror = rej;
      r.readAsDataURL(blob);
    });
  }
  async function deliver({ blob, name }) {
    // Appli Android (APK) : enregistrement dans Téléchargements/KorkedAgri
    if (window.Android && window.Android.saveFile) {
      const res = window.Android.saveFile(name, blob.type, await toB64(blob));
      return res === "ok" ? "Enregistré dans Téléchargements › KorkedAgri" : res;
    }
    // iPhone / navigateurs qui savent partager un fichier : feuille de partage
    try {
      const file = new File([blob], name, { type: blob.type });
      if (navigator.canShare && navigator.canShare({ files: [file] })) {
        await navigator.share({ files: [file], title: name });
        return "Fichier prêt : " + name;
      }
    } catch (e) {
      if (e && e.name === "AbortError") return "";
    }
    // Sinon : téléchargement classique
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
    return "Téléchargé : " + name;
  }

  async function run(kind, btn) {
    const msg = document.getElementById("xmsg");
    if (!seasonRows().length) { msg.textContent = "Aucune ligne dans cette saison."; return; }
    btn.disabled = true; msg.textContent = "Préparation…";
    try {
      const out = kind === "pdf" ? buildPdf() : buildXlsx();
      msg.textContent = await deliver(out);
    } catch (e) {
      msg.textContent = "Export impossible : " + (e && e.message ? e.message : e);
    } finally {
      btn.disabled = false;
    }
  }
  document.getElementById("exXlsx").onclick = e => run("xlsx", e.currentTarget);
  document.getElementById("exPdf").onclick = e => run("pdf", e.currentTarget);
  window.KorkedExport = { buildXlsx, buildPdf };
})();
