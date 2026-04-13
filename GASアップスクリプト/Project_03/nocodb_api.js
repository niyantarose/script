/*━━━━━━ NocoDB 共通設定 ━━━━━━━━━━━━━━━━━━━━━*/
/* ★ 以下 3 つだけ自分の値に置き換えてください */
const VPS_IP         = '133.167.89.53';        // ★VPS IP
const BASE_SLUG      = 'pqidczb7avc3o9z';      // ★ベース slug
const TABLE_SLUG     = 'm4poh99bzwolwd9';      // ★テーブル slug
/*━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━*/

const NOCO_ENDPOINT = `http://${VPS_IP}:8080/api/v1/db/data/v1/${BASE_SLUG}/${TABLE_SLUG}`;
const NOCO_TOKEN    = PropertiesService.getScriptProperties().getProperty('NOCO_TOKEN');

/** Insert 専用 */
function nocodbInsertStock(obj){
  return JSON.parse(UrlFetchApp.fetch(NOCO_ENDPOINT,{
    method : 'post',
    contentType:'application/json',
    headers : {'xc-auth':NOCO_TOKEN},
    payload : JSON.stringify(obj),
    muteHttpExceptions:true
  }).getContentText());
}

/** Upsert (存在すれば Patch) */
function nocodbUpsertStock(obj){
  const where = encodeURIComponent(`ProductID,eq,${obj.ProductID}&LocationCode,eq,${obj.LocationCode}`);
  const rows  = JSON.parse(
      UrlFetchApp.fetch(`${NOCO_ENDPOINT}?where=${where}`,{headers:{'xc-auth':NOCO_TOKEN}})
      .getContentText()).list || [];
  if (!rows.length) return nocodbInsertStock(obj);

  const id = rows[0].Id;
  return JSON.parse(UrlFetchApp.fetch(`${NOCO_ENDPOINT}/${id}`,{
    method:'patch',
    contentType:'application/json',
    headers:{'xc-auth':NOCO_TOKEN},
    payload: JSON.stringify(obj),
    muteHttpExceptions:true
  }).getContentText());
}
