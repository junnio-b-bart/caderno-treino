/**
 * Caderno de Treino — ponte entre o app e esta planilha.
 * Abas: "Treinos" (1 linha por treino), "Séries" (1 linha por série),
 * "Progresso" (resumo automático) e "_app" (oculta, guarda o estado do app).
 *
 * Depois de colar este código: rode a função "configurar" uma vez,
 * depois Implantar > Nova implantação > App da Web.
 */

const ABAS = {
  treinos: { nome: "Treinos", cab: ["Data", "Perfil", "Treino", "Fase", "Séries feitas", "Séries planejadas", "Duração (min)", "Como foi", "Observação", "Registrado em", "ID"] },
  series:  { nome: "Séries",  cab: ["Data", "Perfil", "Treino", "Fase", "Exercício", "Série", "Alvo", "Feito", "Unidade", "Elástico", "Bateu o alvo", "ID do treino"] },
  app:     { nome: "_app",    cab: ["Perfil", "Dados (não editar)", "Atualizado em"] }
};

function aba_(k) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(ABAS[k].nome);
  if (!sh) {
    sh = ss.insertSheet(ABAS[k].nome);
    sh.getRange(1, 1, 1, ABAS[k].cab.length).setValues([ABAS[k].cab]).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Rode uma vez pelo editor (botão ▶ Executar com "configurar" selecionado). */
function configurar() {
  const ss = SpreadsheetApp.getActive();
  aba_("treinos"); aba_("series");
  aba_("app").hideSheet();
  let pr = ss.getSheetByName("Progresso");
  if (!pr) pr = ss.insertSheet("Progresso", 0);
  pr.clear();
  pr.getRange("A1").setValue("Resumo por exercício (atualiza sozinho)").setFontWeight("bold");
  pr.getRange("A3").setFormula(
    "=IFERROR(QUERY('Séries'!A:L,\"select B, E, count(F), max(H), avg(H), max(J), max(A) where B is not null group by B, E order by B, E label count(F) 'Séries feitas', max(H) 'Melhor série', avg(H) 'Média por série', max(J) 'Maior elástico', max(A) 'Última vez'\",1),\"Ainda sem séries registradas\")"
  );
  ["Página1", "Sheet1", "Planilha1"].forEach(n => {
    const s = ss.getSheetByName(n);
    if (s && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
}

function json_(x) {
  return ContentService.createTextOutput(JSON.stringify(x)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const sh = aba_("app");
  const vals = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues() : [];
  const perfis = {};
  vals.forEach(([id, dados]) => { try { perfis[id] = JSON.parse(dados); } catch (err) {} });
  return json_({ ok: true, perfis });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const op = JSON.parse(e.postData.contents);
    if (op.acao === "perfil") salvarPerfil_(op.id, op.perfil);
    else if (op.acao === "treino") registrarTreino_(op.treino, op.series || []);
    else if (op.acao === "removerTreino") removerTreino_(op.id);
    else return json_({ ok: false, erro: "ação desconhecida" });
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, erro: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function salvarPerfil_(id, perfil) {
  const sh = aba_("app");
  const n = sh.getLastRow();
  const ids = n > 1 ? sh.getRange(2, 1, n - 1, 1).getValues().map(r => String(r[0])) : [];
  const linha = ids.indexOf(id);
  const valores = [[id, JSON.stringify(perfil), new Date()]];
  if (linha >= 0) sh.getRange(linha + 2, 1, 1, 3).setValues(valores);
  else sh.getRange(n + 1, 1, 1, 3).setValues(valores);
}

function registrarTreino_(t, series) {
  const tr = aba_("treinos");
  const n = tr.getLastRow();
  const ids = n > 1 ? tr.getRange(2, 11, n - 1, 1).getValues().map(r => String(r[0])) : [];
  if (ids.indexOf(t.id) >= 0) return; // já registrado (reenvio)
  tr.appendRow([t.data, t.perfil, t.treino, t.fase, t.series, t.de, t.duracaoMin, t.comoFoi || "", t.nota || "", new Date(), t.id]);
  if (series.length) {
    const se = aba_("series");
    const linhas = series.map(s => [s.data, s.perfil, s.treino, s.fase, s.exercicio, s.serie, s.alvo, s.feito, s.unidade, s.elastico, s.bateu ? "Sim" : "Não", t.id]);
    se.getRange(se.getLastRow() + 1, 1, linhas.length, linhas[0].length).setValues(linhas);
  }
}

function removerTreino_(id) {
  [["treinos", 11], ["series", 12]].forEach(([k, col]) => {
    const sh = aba_(k);
    const n = sh.getLastRow();
    if (n < 2) return;
    const vals = sh.getRange(2, col, n - 1, 1).getValues();
    for (let i = vals.length - 1; i >= 0; i--) if (String(vals[i][0]) === id) sh.deleteRow(i + 2);
  });
}
