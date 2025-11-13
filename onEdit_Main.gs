function onEdit(e) {
  const sh = e && e.source && e.source.getActiveSheet();
  if (!sh || !e.range || e.range.getRow() === 1) return;
  const name = sh.getName();
  if (name === "Achats") return handleAchats(e);
  if (name === "Stock")  return handleStock(e);
}
