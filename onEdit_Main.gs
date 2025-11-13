function onEdit(e) {
  const ss = e && e.source;
  ensureLegacyFormattingCleared_(ss);

  const sh = ss && ss.getActiveSheet();
  if (!sh || !e.range || e.range.getRow() === 1) return;
  const name = sh.getName();
  if (name === "Achats") return handleAchats(e);
  if (name === "Stock")  return handleStock(e);
}
