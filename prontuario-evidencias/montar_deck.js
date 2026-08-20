// Gera o PowerPoint de demonstração do Portal Prontuário (Evidências).
const pptxgen = require("pptxgenjs");
const fs = require("fs");
const sizeOf = (p) => {
  // lê dimensões PNG do header (largura/altura em bytes 16..24)
  const b = fs.readFileSync(p);
  return { w: b.readUInt32BE(16), h: b.readUInt32BE(20) };
};

const SHOTS = "shots";
const p = new pptxgen();
p.defineLayout({ name: "W", width: 13.333, height: 7.5 });
p.layout = "W";

// ---- paleta (a partir do azul do portal) ----
const NAVY = "1F4E78";      // dominante
const NAVY_DK = "14395C";   // fundo escuro
const INK = "1B2733";
const SLATE = "5A6B7B";
const CLOUD = "EEF2F6";     // fundo claro
const LINE = "D8E0E8";
const AMBER = "C8811A";     // acento quente
const GREEN = "2E7D32";
const RED = "B23A2E";
const WHITE = "FFFFFF";

const FS = "Calibri";
const FH = "Cambria";

const PW = 13.333, PH = 7.5, M = 0.6;

// helpers -------------------------------------------------
function bg(s, color) { s.background = { color }; }

function shot(slide, file, box) {
  // encaixa a imagem dentro de box mantendo proporção, alinhada ao topo
  const { w, h } = sizeOf(`${SHOTS}/${file}`);
  const r = w / h;
  let iw = box.w, ih = iw / r;
  if (ih > box.h) { ih = box.h; iw = ih * r; }
  const x = box.x + (box.w - iw) / 2;
  const y = box.align === "top" ? box.y : box.y + (box.h - ih) / 2;
  slide.addImage({
    path: `${SHOTS}/${file}`, x, y, w: iw, h: ih,
    rounding: false,
    shadow: { type: "outer", color: "8A99A8", opacity: 0.45, blur: 9, offset: 3, angle: 90 },
  });
  // moldura fina
  slide.addShape(p.ShapeType.rect, { x, y, w: iw, h: ih, fill: { type: "none" }, line: { color: LINE, width: 1 } });
  return { x, y, w: iw, h: ih };
}

function kicker(slide, txt, x, y, color = AMBER) {
  slide.addText(txt.toUpperCase(), {
    x, y, w: 6, h: 0.3, fontFace: FS, fontSize: 11, bold: true,
    color, charSpacing: 3, margin: 0, align: "left",
  });
}

function title(slide, txt, x, y, w, color = INK, size = 30) {
  slide.addText(txt, { x, y, w, h: 0.7, fontFace: FH, fontSize: size, bold: true, color, margin: 0, align: "left" });
}

// ============================================================
// SLIDE 1 — Capa
// ============================================================
let s = p.addSlide(); bg(s, NAVY_DK);
s.addShape(p.ShapeType.rect, { x: 0, y: 0, w: PW, h: PH, fill: { color: NAVY_DK } });
// bloco de marca
s.addText("SAMBAÍBA · PRONTUÁRIO", {
  x: M, y: 2.15, w: 11, h: 0.4, fontFace: FS, fontSize: 14, bold: true,
  color: "9FC0DE", charSpacing: 4, margin: 0,
});
s.addText("Portal de Evidências", {
  x: M, y: 2.55, w: 12, h: 1.4, fontFace: FH, fontSize: 60, bold: true, color: WHITE, margin: 0,
});
s.addText(
  "Registro, validação e envio das evidências de conduta operacional — do apontamento à ciência do colaborador.",
  { x: M, y: 4.05, w: 9.6, h: 0.9, fontFace: FS, fontSize: 17, color: "CFE0EF", margin: 0, lineSpacingMultiple: 1.15 }
);
// faixa de KPIs na capa
const capKPI = [["117", "enviadas"], ["81", "na fila Nimer"], ["47%", "taxa de ciência"], ["210", "evidências"]];
let cx = M;
capKPI.forEach(([n, l], i) => {
  s.addText(n, { x: cx, y: 5.25, w: 2.4, h: 0.75, fontFace: FH, fontSize: 40, bold: true, color: "6FB2E4", margin: 0 });
  s.addText(l.toUpperCase(), { x: cx, y: 6.0, w: 2.6, h: 0.35, fontFace: FS, fontSize: 11, bold: true, color: "9FC0DE", charSpacing: 2, margin: 0 });
  cx += 2.75;
});
s.addText("Demonstração do projeto · " + "Agosto/2026", {
  x: M, y: 6.85, w: 8, h: 0.3, fontFace: FS, fontSize: 12, color: "7C97AE", margin: 0,
});

// ============================================================
// SLIDE 2 — Problema & Solução
// ============================================================
s = p.addSlide(); bg(s, CLOUD);
kicker(s, "Por que o portal existe", M, M);
title(s, "Do apontamento solto ao prontuário rastreável", M, M + 0.35, 12, INK, 30);

const cards = [
  ["Antes", "Evidências de conduta espalhadas em prints, planilhas e mensagens — sem padrão, sem histórico e sem rastreio de quem viu o quê.", RED],
  ["Agora", "Um fluxo único: o CCO registra a evidência com foto, o ADH valida, e o sistema controla o envio à Nimer e a ciência do colaborador.", NAVY],
  ["Ganho", "Cada evento tem status, autor, data e documento assinado — auditável de ponta a ponta e pronto para o RH e o jurídico.", GREEN],
];
const cw = (PW - 2 * M - 2 * 0.4) / 3;
cards.forEach(([h, b, c], i) => {
  const x = M + i * (cw + 0.4);
  s.addShape(p.ShapeType.roundRect, { x, y: 2.15, w: cw, h: 3.9, rectRadius: 0.09, fill: { color: WHITE }, line: { color: LINE, width: 1 }, shadow: { type: "outer", color: "C3CEDA", opacity: 0.5, blur: 8, offset: 3, angle: 90 } });
  s.addShape(p.ShapeType.ellipse, { x: x + 0.35, y: 2.5, w: 0.3, h: 0.3, fill: { color: c } });
  s.addText(h, { x: x + 0.8, y: 2.42, w: cw - 1, h: 0.45, fontFace: FH, fontSize: 20, bold: true, color: INK, margin: 0, valign: "middle" });
  s.addText(b, { x: x + 0.35, y: 3.15, w: cw - 0.7, h: 2.6, fontFace: FS, fontSize: 15, color: SLATE, margin: 0, lineSpacingMultiple: 1.2, valign: "top" });
});
s.addText("Perfis: CCO · CAC · Coordenador registram e corrigem   |   ADH valida (aprova ou devolve com motivo)   |   ADMIN administra tudo",
  { x: M, y: 6.35, w: PW - 2 * M, h: 0.5, fontFace: FS, fontSize: 13, italic: true, color: NAVY, align: "center", margin: 0 });

// ============================================================
// SLIDE 3 — Visão geral em números (KPIs)
// ============================================================
s = p.addSlide(); bg(s, CLOUD);
kicker(s, "Início · Visão geral", M, M);
title(s, "O estado do fluxo em números", M, M + 0.35, 12, INK, 30);
const kpis = [
  ["6", "Aguardando evento", NAVY], ["6", "Na validação do ADH", NAVY],
  ["0", "Negadas (voltaram ao CCO)", GREEN], ["81", "Fila de envio Nimer", AMBER],
  ["117", "Enviadas", NAVY],
];
const kw = (PW - 2 * M - 4 * 0.35) / 5;
kpis.forEach(([n, l, c], i) => {
  const x = M + i * (kw + 0.35);
  s.addShape(p.ShapeType.roundRect, { x, y: 2.2, w: kw, h: 2.0, rectRadius: 0.08, fill: { color: WHITE }, line: { color: LINE, width: 1 }, shadow: { type: "outer", color: "C3CEDA", opacity: 0.5, blur: 7, offset: 3, angle: 90 } });
  s.addText(n, { x, y: 2.42, w: kw, h: 0.95, fontFace: FH, fontSize: 46, bold: true, color: c, align: "center", margin: 0 });
  s.addText(l, { x: x + 0.15, y: 3.42, w: kw - 0.3, h: 0.7, fontFace: FS, fontSize: 12, color: SLATE, align: "center", margin: 0, lineSpacingMultiple: 1.0 });
});
// banner de status
s.addShape(p.ShapeType.roundRect, { x: M, y: 4.65, w: PW - 2 * M, h: 0.95, rectRadius: 0.07, fill: { color: "FBF1D8" }, line: { color: "E7CF94", width: 1 } });
s.addText([
  { text: "⚠  Envios à Nimer bloqueados até segunda ordem.  ", options: { bold: true, color: AMBER } },
  { text: "A fila acumula os eventos aprovados pelo ADH; nada é transmitido sem liberação expressa.", options: { color: "7A6320" } },
], { x: M + 0.3, y: 4.65, w: PW - 2 * M - 0.6, h: 0.95, fontFace: FS, fontSize: 15, valign: "middle", margin: 0 });
s.addText("Números lidos da tela Início do portal (Agosto/2026).", { x: M, y: 6.75, w: 10, h: 0.3, fontFace: FS, fontSize: 11, italic: true, color: SLATE, margin: 0 });

// ============================================================
// SLIDE 4 — O fluxo (diagrama)
// ============================================================
s = p.addSlide(); bg(s, NAVY_DK);
kicker(s, "Como funciona", M, M, "8FB8DD");
title(s, "O ciclo de vida de uma evidência", M, M + 0.35, 12, WHITE, 30);
const steps = [
  ["1", "Nova evidência", "CCO registra RE, data, descrição e foto"],
  ["2", "Validação (ADH)", "Aprova ou devolve com motivo"],
  ["3", "Fila Nimer", "Aprovadas aguardam o envio"],
  ["4", "Enviada", "Transmitida com a foto de evidência"],
  ["5", "Retorno Nimer", "Colaborador dá ciência (doc. assinado)"],
];
const sw = (PW - 2 * M - 4 * 0.45) / 5;
steps.forEach(([n, h, b], i) => {
  const x = M + i * (sw + 0.45);
  s.addShape(p.ShapeType.roundRect, { x, y: 2.7, w: sw, h: 2.5, rectRadius: 0.08, fill: { color: "1C4A73" }, line: { color: "3D6E9C", width: 1 } });
  s.addShape(p.ShapeType.ellipse, { x: x + sw / 2 - 0.32, y: 2.35, w: 0.64, h: 0.64, fill: { color: AMBER } });
  s.addText(n, { x: x + sw / 2 - 0.32, y: 2.35, w: 0.64, h: 0.64, fontFace: FH, fontSize: 22, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0 });
  s.addText(h, { x: x + 0.12, y: 3.15, w: sw - 0.24, h: 0.7, fontFace: FS, fontSize: 15, bold: true, color: WHITE, align: "center", margin: 0, valign: "top" });
  s.addText(b, { x: x + 0.15, y: 3.85, w: sw - 0.3, h: 1.2, fontFace: FS, fontSize: 12, color: "BFD5E8", align: "center", margin: 0, lineSpacingMultiple: 1.1 });
  if (i < steps.length - 1) {
    s.addText("→", { x: x + sw - 0.05, y: 2.7, w: 0.55, h: 2.5, fontFace: FS, fontSize: 24, bold: true, color: "6FB2E4", align: "center", valign: "middle", margin: 0 });
  }
});
// desvio
s.addShape(p.ShapeType.roundRect, { x: M, y: 5.7, w: PW - 2 * M, h: 0.85, rectRadius: 0.07, fill: { color: "2A1F14" }, line: { color: AMBER, width: 1 } });
s.addText([
  { text: "Desvio  ", options: { bold: true, color: AMBER } },
  { text: "Se o ADH devolve, a evidência volta ao CCO para correção antes de reentrar no fluxo — nada avança sem validação.", options: { color: "E8D9C4" } },
], { x: M + 0.3, y: 5.7, w: PW - 2 * M - 0.6, h: 0.85, fontFace: FS, fontSize: 14, valign: "middle", margin: 0 });

// ============================================================
// SLIDES DE TELA — helper
// ============================================================
function screenSlide(kick, ttl, file, notes, cropTall) {
  const sl = p.addSlide(); bg(sl, CLOUD);
  const LW = 3.75; // largura da coluna esquerda (título + descrição)
  kicker(sl, kick, M, 0.5);
  sl.addText(ttl, { x: M, y: 0.85, w: LW, h: 1.05, fontFace: FH, fontSize: 24, bold: true, color: INK, margin: 0, align: "left", valign: "top" });
  // descrição à esquerda
  sl.addText(notes.map((t, i) => ({
    text: t, options: { bullet: { code: "2022", indent: 14 }, breakLine: true, color: SLATE, paraSpaceAfter: 8 },
  })), { x: M, y: 2.15, w: LW, h: 4.6, fontFace: FS, fontSize: 14, margin: 0, valign: "top", lineSpacingMultiple: 1.1 });
  // imagem à direita (começa após a coluna esquerda, com folga)
  shot(sl, file, { x: 4.65, y: 0.55, w: PW - 4.65 - M, h: PH - 0.55 - 0.55, align: "top" });
  return sl;
}

// SLIDE 5 — Nova evidência
screenSlide("Registro", "Nova evidência", "02_nova_evidencia.png", [
  "Ponto de entrada do fluxo, usado por CCO, CAC e Coordenador.",
  "Campos: RE do colaborador, data do evento e descrição da falha de conduta.",
  "Data limitada: de hoje até 6 meses atrás — não aceita data futura.",
  "Foto de evidência por arraste ou câmera: jpg, png, webp ou HEIC (iPhone), até 30 MB.",
]);

// SLIDE 6 — Evidências (lista)
screenSlide("Consulta", "Evidências", "03_evidencias_crop.png", [
  "Lista central de todas as evidências — 210 registros no período.",
  "Busca por RE ou nome e filtro por status.",
  "Status: aguardando evento, vinculada, enviada, aprovada.",
  "Exportação para Excel e PDF; clique na linha abre os detalhes.",
]);

// SLIDE 7 — Validação (ADH)
screenSlide("Validação", "Validação do ADH", "04_validacao_crop.png", [
  "Fila de aprovação do ADH — 6 evidências aguardando.",
  "Cada card traz a foto de evidência, o colaborador (nome e RE) e o evento do sistema.",
  "Dois caminhos: Aprovar envia à fila da Nimer; Devolver retorna ao CCO com motivo.",
  "As imagens vêm do MaxTrack e dos sistemas de operação (mapa, telemetria, escala).",
]);

// SLIDE 8 — Fila Nimer
screenSlide("Envio", "Fila de envio à Nimer", "05_fila_nimer.png", [
  "Acumula os eventos já aprovados pelo ADH, prontos para transmitir.",
  "Envios bloqueados até segunda ordem — nada sai sem liberação expressa.",
  "Fila vazia na captura: o gargalo é a liberação, não o acúmulo.",
  "Controle central de quando e o que é transmitido à Nimer.",
]);

// SLIDE 9 — Retorno Nimer
screenSlide("Ciência", "Retorno da Nimer", "06_retorno_nimer_crop.png", [
  "Fecha o ciclo: quantos colaboradores deram ciência do documento.",
  "117 enviados com evidência · 55 deram ciência · 47% de taxa · 32 h de espera média.",
  "Ciência chega por webhook (warning.signed) e casa por RE e data do evento.",
  "Considera só envios que levam a foto de evidência — o fluxo do portal.",
]);

// SLIDE 10 — Usuários
screenSlide("Administração", "Gestão de usuários", "07_usuarios.png", [
  "Área ADMIN: cria, edita, redefine senha e desativa usuários.",
  "Perfis: CCO · CAC · Coordenador (registram e corrigem).",
  "ADH valida — aprova ou devolve com motivo.",
  "ADMIN faz tudo, inclusive esta tela.",
]);

// ============================================================
// SLIDE 11 — Status atual & próximos passos
// ============================================================
s = p.addSlide(); bg(s, NAVY_DK);
kicker(s, "Onde estamos", M, M, "8FB8DD");
title(s, "Status atual e próximos passos", M, M + 0.35, 12, WHITE, 30);
const colL = [
  ["Em produção", "Portal completo em uso: registro, consulta, validação do ADH, fila e retorno com ciência.", GREEN],
  ["117 evidências enviadas", "Com 47% de taxa de ciência e documento assinado por punho próprio arquivado.", "6FB2E4"],
];
const colR = [
  ["Envio à Nimer liberar", "Fila com 81 eventos aprovados aguardando a liberação expressa para transmitir.", AMBER],
  ["Elevar a taxa de ciência", "Reduzir as 32 h de espera média e os 53% ainda sem ciência.", "6FB2E4"],
];
function bl(list, x) {
  list.forEach(([h, b, c], i) => {
    const y = 2.35 + i * 1.9;
    s.addShape(p.ShapeType.roundRect, { x, y, w: (PW - 2 * M - 0.6) / 2, h: 1.6, rectRadius: 0.08, fill: { color: "1C4A73" }, line: { color: "3D6E9C", width: 1 } });
    s.addShape(p.ShapeType.ellipse, { x: x + 0.3, y: y + 0.32, w: 0.26, h: 0.26, fill: { color: c } });
    s.addText(h, { x: x + 0.72, y: y + 0.22, w: (PW - 2 * M - 0.6) / 2 - 1, h: 0.5, fontFace: FH, fontSize: 18, bold: true, color: WHITE, margin: 0, valign: "middle" });
    s.addText(b, { x: x + 0.3, y: y + 0.78, w: (PW - 2 * M - 0.6) / 2 - 0.6, h: 0.7, fontFace: FS, fontSize: 13.5, color: "C7DBEC", margin: 0, lineSpacingMultiple: 1.1 });
  });
}
bl(colL, M);
bl(colR, M + (PW - 2 * M - 0.6) / 2 + 0.6);
s.addText("Sambaíba · Prontuário — Portal de Evidências", { x: M, y: 6.7, w: 10, h: 0.3, fontFace: FS, fontSize: 12, color: "7C97AE", margin: 0 });

// ------------------------------------------------------------
p.writeFile({ fileName: "Evidencias_Portal_Prontuario.pptx" }).then((f) => console.log("gerado:", f));
