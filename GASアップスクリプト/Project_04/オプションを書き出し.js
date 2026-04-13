function オプション一覧に書き出し() {

  const ss   = SpreadsheetApp.getActive();
  const 入力 = ss.getSheetByName('①商品入力シート');
  const 出力 = ss.getSheetByName('②オプション入力シート');

  if (!入力 || !出力) {
    throw new Error('① or ②シートが見つかりません');
  }

  const lastRow = 入力.getLastRow();

  // クリア
  if (出力.getLastRow() > 1) {
    出力.getRange(2,1,出力.getLastRow(),4)
      .clearContent()
      .clearFormat();
  }

  if (lastRow < 3) return;

  const src = 入力.getRange(3,1,lastRow-2,4).getValues();

  let parents = [];

  src.forEach(r => {

    const code  = String(r[2]||'').trim(); // C列
    const count = Number(r[3]||0);         // D列

    if (!code || count <= 1) return;

    parents.push({code,count});
  });

  parents.sort((a,b)=>a.code.localeCompare(b.code));

  let row = 2;
  let prev = '';

  parents.forEach(p=>{

    if (prev && prev !== p.code) row++;

    for (let i=1;i<=p.count;i++){

      // a
      出力.getRange(row,1,1,4).setValues([[
        p.code,
        '即納（日本在庫）',
        `${p.code}-${i}a`,
        ''
      ]]);
      row++;

      // b
      出力.getRange(row,1,1,4).setValues([[
        p.code,
        'お取り寄せ（韓国から）',
        `${p.code}-${i}b`,
        ''
      ]]);

      // 罫線
      出力.getRange(row,1,1,4).setBorder(
        false,false,true,false,false,false
      );

      row++;
    }

    prev = p.code;
  });

}
function 作成Yahooオプション文字列_(code){

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('②オプション入力シート');

  const last = sh.getLastRow();
  if (last < 2) return '';

  const rows = sh.getRange(2,1,last-1,4)
    .getValues()
    .filter(r=>String(r[0]).trim()===code);

  if (rows.length===0) return '';

  let groups = {};

  rows.forEach(r=>{

    const sub = r[2];
    const name = r[3];

    const m = sub.match(/-(\d+)[ab]$/);
    const key = m ? m[1] : sub.slice(0,-1);

    if (!groups[key]) {
      groups[key]={name:name,items:[]};
    }

    groups[key].items.push(sub);
  });

  const keys = Object.keys(groups)
    .sort((a,b)=>Number(a)-Number(b));

  let aList=[];
  let bList=[];

  keys.forEach((k,i)=>{

    const no = i+1;
    const name = groups[k].name || `${code}-${k}`;

    groups[k].items.forEach(sub=>{

      if (sub.endsWith('a')){
        aList.push(
          `★在庫の設定:即納（日本在庫）#★種類の選択:${no}.${name}=${sub}`
        );
      }

      if (sub.endsWith('b')){
        bList.push(
          `★在庫の設定:お取り寄せ（韓国から）#★種類の選択:${no}.${name}=${sub}`
        );
      }

    });
  });

  return [...aList,...bList].join('&');
}
function 作成Yahooオプション項目名_(code){

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('②オプション入力シート');

  const last = sh.getLastRow();
  if (last<2){
    return '★在庫の選択 即納（日本在庫） お取り寄せ（韓国から）';
  }

  const rows = sh.getRange(2,1,last-1,4)
    .getValues()
    .filter(r=>String(r[0]).trim()===code);

  if (rows.length===0){
    return '★在庫の選択 即納（日本在庫） お取り寄せ（韓国から）';
  }

  let map={};

  rows.forEach(r=>{

    const sub=r[2];
    const name=r[3];

    const m=sub.match(/-(\d+)[ab]$/);
    const key=m?m[1]:sub.slice(0,-1);

    if(!map[key]){
      map[key]=name||`${code}-${key}`;
    }
  });

  const keys=Object.keys(map)
    .sort((a,b)=>Number(a)-Number(b));

  const list=keys.map((k,i)=>{
    return `${i+1}.${map[k]}`;
  });

  return (
    '★在庫の選択 即納（日本在庫） お取り寄せ（韓国から）\n'+
    '★種類の選択 '+list.join(' ')
  );
}
function Yahoo商品登録シートに反映(){

  const ss = SpreadsheetApp.getActive();

  const src = ss.getSheetByName('①商品入力シート');
  const dst = ss.getSheetByName('Yahoo商品登録シート');

  const last = src.getLastRow();
  if (last<3) return;

  const data = src.getRange(3,1,last-2,20).getValues();

  let out=[];

  data.forEach(r=>{

    const code = String(r[2]).trim();

    if(!code){
      out.push(r);
      return;
    }

    let detail = 作成Yahooオプション文字列_(code);
    let header = 作成Yahooオプション項目名_(code);

    // サブ1だけのとき
    if(!detail){

      const a=code+'a';
      const b=code+'b';

      detail =
        `★在庫の選択:即納（日本在庫）=${a}`+
        `&★在庫の選択:お取り寄せ（韓国から）=${b}`;

      header =
        '★在庫の選択 即納（日本在庫） お取り寄せ（韓国から）';
    }

    r[3]=detail; // D列
    r[5]=header; // F列

    out.push(r);
  });

  dst.getRange(3,1,out.length,out[0].length)
     .setValues(out);
}
