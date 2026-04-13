function 中国価格_onEdit(e) {  // ← ここだけ変更、中身はそのまま

  const sheetName = "中国価格計算";
  const urlCell = "B2";
  const priceCell = "B3";

  const sheet = e.source.getActiveSheet();
  if(sheet.getName() !== sheetName) return;

  if(e.range.getA1Notation() !== urlCell) return;

  const url = e.range.getValue();
  if(!url) return;

  const idMatch = url.match(/id=(\d+)/);

  if(!idMatch) return;

  const itemId = idMatch[1];

  const api =
  "https://h5api.m.taobao.com/h5/mtop.taobao.detail.getdetail/6.0/?data=" +
  encodeURIComponent(JSON.stringify({
      itemNumId:itemId
  }));

  const res = UrlFetchApp.fetch(api);
  const json = JSON.parse(res.getContentText());

  try{

    const price =
    json.data.price.price.priceText;

    sheet.getRange(priceCell).setValue(price);

  }catch(err){

    sheet.getRange(priceCell).setValue("価格取得失敗");

  }

}