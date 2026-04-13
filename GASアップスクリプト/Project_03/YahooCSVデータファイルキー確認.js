function testType() {
  [1,2,3,4,9,11].forEach(id => {
    try {
      const key = XmlService
                  .parse(UrlFetchApp.fetch(
                     'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/downloadRequest',
                     {method:'post',contentType:'application/x-www-form-urlencoded',
                      headers:{Authorization:'Bearer '+getValidAccessToken()},
                      payload:`seller_id=${SELLER_ID}&type=${id}`}
                  ).getContentText())
                  .getRootElement().getChildText('FileKey');
      Logger.log('type='+id+' → fileKey='+(key||'なし'));
    } catch(e){ Logger.log('type='+id+' → ERROR'); }
  });
}