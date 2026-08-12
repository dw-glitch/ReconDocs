export const REPORT_COLORS = {
  navy: "153A5C",
  teal: "0C7657",
  amber: "A56812",
  red: "A64035",
  pale: "F4F7F9",
  white: "FFFFFF",
};

export const DEFAULT_REPORT_SUBTITLE = "Conferência entre SGP, Consulta Geral do SIGEM e Documentos Previstos";

export function createBrandLogoDataUrl() {
  const canvas = document.createElement("canvas");
  canvas.width = 820;
  canvas.height = 170;
  const context = canvas.getContext("2d");
  if (!context) throw new Error("Não foi possível preparar a identidade visual do relatório.");

  context.clearRect(0, 0, canvas.width, canvas.height);
  context.fillStyle = "#153A5C";
  context.beginPath();
  context.roundRect(16, 18, 132, 132, 26);
  context.fill();

  context.strokeStyle = "#FFFFFF";
  context.lineWidth = 10;
  context.lineCap = "round";
  context.lineJoin = "round";
  context.beginPath();
  context.moveTo(45, 66);
  context.lineTo(116, 66);
  context.moveTo(45, 66);
  context.lineTo(61, 50);
  context.moveTo(45, 66);
  context.lineTo(61, 82);
  context.stroke();

  context.beginPath();
  context.moveTo(116, 101);
  context.lineTo(45, 101);
  context.moveTo(116, 101);
  context.lineTo(100, 85);
  context.moveTo(116, 101);
  context.lineTo(100, 117);
  context.stroke();

  context.fillStyle = "#0C7657";
  context.beginPath();
  context.arc(122, 126, 25, 0, Math.PI * 2);
  context.fill();
  context.strokeStyle = "#FFFFFF";
  context.lineWidth = 6;
  context.beginPath();
  context.moveTo(110, 126);
  context.lineTo(119, 135);
  context.lineTo(136, 115);
  context.stroke();

  context.fillStyle = "#153A5C";
  context.font = "700 64px Arial, sans-serif";
  context.fillText("RECON", 178, 94);
  const reconWidth = context.measureText("RECON").width;
  context.fillStyle = "#0C7657";
  context.fillText("DOCS", 178 + reconWidth, 94);
  context.fillStyle = "#657B8D";
  context.font = "600 22px Arial, sans-serif";
  context.letterSpacing = "3px";
  context.fillText("CONFERÊNCIA SGP × SIGEM", 182, 132);
  return canvas.toDataURL("image/png");
}

export function addReportBranding(sheet, logoImageId, endColumn, title, subtitle = DEFAULT_REPORT_SUBTITLE) {
  sheet.addImage(logoImageId, {
    tl: { col: 0.15, row: 0.18 },
    ext: { width: 218, height: 45 },
  });
  sheet.mergeCells(`D1:${endColumn}1`);
  sheet.getCell("D1").value = title;
  sheet.getCell("D1").font = { bold: true, size: 16, color: { argb: "153A5C" } };
  sheet.getCell("D1").alignment = { vertical: "middle", horizontal: "left" };
  sheet.mergeCells(`D2:${endColumn}2`);
  sheet.getCell("D2").value = subtitle;
  sheet.getCell("D2").font = { size: 10, color: { argb: "52687B" } };
  sheet.mergeCells(`D3:${endColumn}3`);
  sheet.getCell("D3").value = `Relatório gerado em ${new Intl.DateTimeFormat("pt-BR", { dateStyle: "short", timeStyle: "short" }).format(new Date())}`;
  sheet.getCell("D3").font = { size: 9, italic: true, color: { argb: "7B8F9D" } };
  sheet.getRow(1).height = 29;
  sheet.getRow(2).height = 20;
  sheet.getRow(3).height = 18;
  sheet.getRow(4).height = 8;
}
