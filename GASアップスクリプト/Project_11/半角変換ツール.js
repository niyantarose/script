function TOHALFWIDTH(text) {
  if (text === null || text === "") return "";
  return text
    .toString()
    // 全角英数字を半角に
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) {
      return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
    })
    // 全角ハイフンやアンダーバーを半角に
    .replace(/[ー＿]/g, function(s) {
      if (s === "ー") return "-";
      if (s === "＿") return "_";
      return s;
    });
}
