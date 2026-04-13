function testUpsert(){
  const obj = {
    ProductID   : 'TEST-999',
    LocationCode: 'SHELF-X1',
    Quantity    : 7,
    LastUpdated : new Date().toISOString()
  };
  const res = nocodbUpsertStock(obj);
  Logger.log(res);      // ← 200/201 が返れば成功
}
