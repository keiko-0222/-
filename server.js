/**
 * 集計表（Excelマクロブック）検索サーバー
 * --------------------------------------------------
 * このフォルダ内にある「対象モール名を含むファイル名」の .xlsm/.xlsx を
 * 自動で探し、「集計表」シートを読み込んで、注文日 または お客様番号の
 * 範囲条件で絞り込み、進捗管理一覧（index.html）にそのまま渡します。
 *
 * 起動方法：start.bat をダブルクリック（初回のみ npm install が走ります）
 *
 * 列がずれている／シート名が違う場合は、下の COL / SHEET_NAME を
 * 実際のExcelに合わせて書き換えてください。
 */
const express = require('express');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');

const app = express();
const PORT = process.env.PORT || 3000;

// Excelファイル（複数のモール別マクロブック）を探すフォルダ。
// 既定ではこの server.js と同じフォルダを見ます。
const DATA_DIR = __dirname;

// 読み込み対象のシート名
const SHEET_NAME = '集計表';

// ラジオボタンに出すモール一覧（ファイル名にこの文字列が含まれるものを検索対象にします）
const CHANNELS = ['yahoo', 'オフィシャル', 'アマゾン', 'ギフトモール', 'うちのこいちばん', 'ふるさと納税', '社内', '4you'];

// 集計表シートの列構成（0始まり。A列=0, B列=1 ... のExcel列インデックス）
// ※サンプルファイル「新システム2025年７月_4you売上帳.xlsm」の集計表シートを基に設定しています。
//   他のモールのファイルで列がずれている場合はここを調整してください。
const COL = {
  id: 0,            // A列  お客様番号
  branch: 1,        // B列  枝番
  no: 2,            // C列  楽天№（注文番号）
  date: 3,           // D列  注文日
  name: 4,           // E列  ご注文者
  ordererZip: 5,     // F列  注文者郵便番号
  ordererAddr1: 6,   // G列  注文者住所
  ordererAddr2: 7,   // H列  注文者住所2
  ordererPhone: 8,   // I列  注文者電話番号
  shipName: 9,       // J列  キット送り先名前
  shipZip: 10,       // K列  送付先郵便番号
  shipAddr1: 11,      // L列  送付先住所
  shipPhone: 12,      // M列  送付先電話番号
  productName: 13,    // N列  商品名
  spec: 14,           // O列  規格
  price: 15,          // P列  単価
  qty: 16,            // Q列  個数
  amount: 17,         // R列  金額
  codFee: 20,         // U列  代引手数料
  postage: 21,        // V列  送料
  total: 22,           // W列  合計金額（注文単位）
  point: 23,           // X列  ポイント利用額
  couponStore: 24,     // Y列  店舗発行クーポン
  couponRakuten: 25,   // Z列  楽天発行クーポン
  billing: 26,         // AA列 請求金額
  laterFee: 27,        // AB列 後払い手数料
  payment: 28,         // AC列 決済方法
  note: 29,            // AD列 備考
  shipDate: 31,        // AF列 発送日
};

    const productMaster = {
  "チョイスギフト": "2", "そのまんまあんよ": "2", "そのまんまあんよプレミアム": "2",
  "そのまんまおててプレミアム": "2", "そのまんまあんよ  本革 タイプ": "2", "そのまんまおてて  本革 タイプ": "2",
  "手形アートキーホルダー": "2", "足形アートキーホルダー": "2", "天使のゆりかご": "2",
  "天使のゆりかごプレミアム": "2", "天使の誘惑": "2", "天使の誘惑×２個セット": "2",
  "天使の誘惑×３個セット": "2", "満天の輝き": "2", "満天の輝き×２個セット": "2",
  "満天の輝き×３個セット": "2", "満天の輝きミニ": "2", "エンジェルクラウド": "2",
  "エンジェル クラウド×２個セット": "2", "エンジェル クラウド×３個セット": "2",
  "エンジェル メモリーズ": "2", "エンジェルメモリーズ×２個セット": "2", "エンジェルメモリーズ×３個セット": "2",
  "未来への扉": "2", "和〜なごみ〜": "2", "天使の輝き": "2", "天使の輝き×２個セット": "2",
  "天使の輝き×３個セット": "2", "虹の散歩道": "2", "虹の散歩道×2個セット": "2",
  "虹の散歩道×３個セット": "2", "未来への架け橋": "2", "未来への架け橋×３個セット": "2",
  "未来への架け橋×２個セット": "2", "そのまんまオブジェ": "2", "そのまんまオブジェ フォトフレーム": "2",
  "そのまんまオブジェ ファミリーフレーム": "2", "森の天使": "2", "天使の贈り物": "2",
  "小さい手見つけた": "2", "小さい足見つけた": "2", "小さい手足みつけた専用　木製フレーム": "1",
  "飛翔プレミアム 御影石": "2", "飛翔プレミアム 大理石": "2", "飛翔A": "2", "飛翔Ｂ": "2",
  "星座物語": "2", "夢捺": "2", "絆": "2", "絆プレミアム": "2", "絆【差し替え用】": "2",
  "絆 フォトフレーム": "2", "思い出いっぱい": "2", "思い出いっぱい×２個セット": "2",
  "思い出いっぱい×３個セット": "2", "ブラザークラウド": "2", "ファミリークラウドA": "2",
  "エンジェルストーリー": "2", "ブラザーストーリー": "2", "ファミリーストーリー": "2",
  "エンジェル スター": "2", "なかよしこよし": "2", "小さな一歩": "2", "ハッピーメモリーズ": "2",
  "ネイチャーエンジェル": "2", "ネイチャーエンジェル プレミアム": "2", "新夢時計": "2",
  "ファースト ステップ クロック": "2", "ファミリーパーティー": "2", "おいたちの小箱": "2",
  "天使のぬくもり": "2", "天使のぬくもり プレミアム": "2", "天使の微笑み": "2",
  "天使の微笑み×2個セット": "2", "天使の微笑み×３個セット": "2", "ポップアップ　エンジェル": "2",
  "ポップアップブラザー": "2", "ポップアップファミリー": "2", "たんじょうものがたり": "2",
  "色彩の窓": "2", "色彩の窓 兄弟用": "2", "Hamatebako": "2", "はい ちーず": "2",
  "フォトホルダー": "2", "メモリアルカレンダー": "2", "クリスマス アート キーホルダー": "2",
  "アニマルアート１５ｃｍ角サイズ": "2", "アニマルアート２０ｃｍ角サイズ": "2",
  "アニマルアート２５ｃｍ角サイズ": "2", "アニマルアート兄弟用": "2",
  "ドリームオブ ピクチャー フォトフレーム": "2", "ドリーム オブ ピクチャーフレーム×2個セット": "2",
  "ドリーム オブ ピクチャーフレーム×３個セット": "2", "手形アート 足形アート フォトフレーム": "2",
  "ぞうさんのファミリーアート": "2", "手形アート 足形アート レザーキャンバス": "2",
  "初節句 フォトフレーム　男の子": "2", "アニマルピンボード": "2",
  "ぺったんてがたあーとフレーム": "2", "ぺったんてがたあーとフレームA4": "2",
  "ぺったんてがたあーとフレーム2L": "2", "Colorful Hand Print T-shirts": "2",
  "Cheers!! Foot Art Print T-Shirts": "2", "Hand Art T-Shirts": "2",
  "【オプション】手足型修正・デザイン料・特急便料金": "3", "クリスタルなわたし": "2",
  "クリスタルなわたし大": "2", "クリスタルなわたし　ペット": "2", "新夢箱": "2",
  "天使の窓": "2", "手形 足形キーホルダー クリスタルなわたし": "2",
  "手形のキーホルダー クリスタルなわたし": "2", "足形のキーホルダー クリスタルなわたし": "2",
  "ミッフィー手形足形キット": "1", "オリジナル手形足型キット": "1",
  "大サイズ発色紙２枚×発色液１パック": "1", "マイエンジェルフォトフレーム": "2",
  "思い出の足跡": "2", "小さい足見つけた　ペット用": "2", "そのまんまシルエット": "2",
  "壮健": "2", "年寿": "2", "その他": "3", "命名 フォトフレーム": "2",
  "手形アートギャラリーＡ4": "2", "手形アートギャラリーＡ3": "2", "パームカラーズ": "1",
  "はいポーズ": "2", "命名 フォトフレームMini": "2", "エンジェルスマイル": "2",
  "Mini leather key ring": "2", "そのまんまオブジェフォトフレーム兄弟タイプ": "2",
  "Monthly　Banner": "1", "命名Foot Print Photo Flame": "2", "Love Clock": "2",
  "New Bornバナー": "1", "Happy Birthday バナー": "1", "令和　年月日バナー": "1",
  "数字バナー　10個": "1", "名入れプレート": "1", "お名前レターバナー": "1",
  "Woody　ひな飾り": "1", "Woody　五月人形": "1", "バナーフルセット": "1",
  "おしゃれかわいい手足形アートキーホルダー": "2", "おしゃれかわいい足形アートキーホルダー": "2",
  "おしゃれかわいい手形アートキーホルダー": "2", "peaceful key ring": "2",
  "painter key ring": "2", "PICAKE（ピケーキ）": "1", "エンジェルウィング": "1",
  "WOODEN CHARM": "1", "天使のはね　木製バナー": "1", "まあるい天使の木時計": "2",
  "星に願いをこめて": "2", "百日祝い　お食い初めバナー": "1", "紙にこだわった命名書": "1",
  "ポスターハンガー": "1", "パームカラーズ【命名書】": "1", "PER TE": "1",
  "PER TE 手形": "1", "PER TE 足形": "1", "身長ベア　フォトフレーム": "2",
  "端午の節句　木製名前札": "1", "端午の節句　アクリル名前札": "1", "名入れキーホルダー": "1",
  "天然木の和柄命名プレート": "1", "革職人名前キーホルダー": "1", "やさしいっ手": "1",
  "ファーストアートタグ": "1", "egao ブレスレット": "1", "hug ネックレス": "1",
  "ファーストアートフルセット": "1", "formal　ブレスレット": "1", "吹き出し　バナー": "1",
  "マンスリーカードパームカラーズ": "1", "マンスリーカード": "1", "マンスリー　フレーム　パーム": "1",
  "マンスリーカード　やさしいっ手": "1", "カード　フレーム　やさしいっ手": "1",
  "パームカラーズ　メゾカラー": "1", "身長計　マンスリー　パーム　スタンド": "1",
  "はじめてのお友達": "2", "うちのレジェンド": "1", "彫刻面板のみ(未来への架け橋 )": "2",
  "ちいさな記念館": "1", "ファーストカット アート 台紙": "1"
};
app.use(express.static(__dirname));

// ---- ユーティリティ ----
function pad(n) { return String(n).padStart(2, '0'); }

function dateToISO(d) {
  if (!(d instanceof Date) || isNaN(d)) return '';
  return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
}

function dateToDisplay(d) {
  if (!(d instanceof Date) || isNaN(d)) return '';
  return `${d.getUTCFullYear()}/${pad(d.getUTCMonth() + 1)}/${pad(d.getUTCDate())}`;
}

// お客様番号は "2600001" のような数字だけのものと、"A123" のように
// アルファベットで始まるものが混在する。単純な文字列比較だと
// "A9" > "A10" のように桁数が違うと正しく並ばないため、
// 文字部分と数字部分に分けて比較する「自然順ソート」で範囲判定する。
function naturalCompare(a, b) {
  const split = s => String(s).match(/\d+|\D+/g) || [];
  const pa = split(a), pb = split(b);
  const len = Math.max(pa.length, pb.length);
  for (let i = 0; i < len; i++) {
    const x = pa[i] || '', y = pb[i] || '';
    if (x === y) continue;
    const isNumX = /^\d+$/.test(x), isNumY = /^\d+$/.test(y);
    if (isNumX && isNumY) {
      const diff = Number(x) - Number(y);
      if (diff !== 0) return diff < 0 ? -1 : 1;
    } else {
      return x < y ? -1 : 1;
    }
  }
  return 0;
}

// 指定モール名を含む .xlsm / .xlsx ファイルをこのフォルダから探す
// （Excelを開いた時にできるロックファイル "~$..." は除外）
function findChannelFiles(channel) {
  const all = fs.readdirSync(DATA_DIR).filter(f =>
    /\.(xlsm|xlsx)$/i.test(f) && !f.startsWith('~$')
  );
  const key = String(channel).toLowerCase();
  return all.filter(f => f.toLowerCase().includes(key));
}

// 集計表シートを読み込み、データ行（配列の配列）を返す
// ヘッダー／合計行の行数がファイルによって違っても良いように、
// 「D列が日付になっている最初の行」を自動でデータ開始行とみなす
function readSheetRows(filePath) {
  const wb = XLSX.readFile(filePath, { cellDates: true });
  const sheetName = wb.SheetNames.includes(SHEET_NAME)
    ? SHEET_NAME
    : wb.SheetNames.find(n => n.includes('集計'));
  if (!sheetName) return null;

  const sheet = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true });

  let startRow = -1;
  for (let i = 0; i < rows.length; i++) {
    if (rows[i] && rows[i][COL.date] instanceof Date) { startRow = i; break; }
  }
  if (startRow === -1) return [];

  const dataRows = [];
  let blankStreak = 0;
  for (let i = startRow; i < rows.length; i++) {
    const r = rows[i];
    if (!r || r[COL.id] === null || r[COL.id] === undefined || r[COL.id] === '') {
      blankStreak++;
      if (blankStreak > 30) break;
      continue;
    }
    blankStreak = 0;
    dataRows.push(r);
  }
  return dataRows;
}

// 明細行（同じ楽天№の複数行）を1注文にまとめる
function buildOrders(rows) {
  const map = new Map();
  const groupSets = new Map();

  for (const r of rows) {
    const no = r[COL.no] != null ? String(r[COL.no]).trim() : '';
    if (!no) continue;

    if (!map.has(no)) {
      map.set(no, {
        id: r[COL.id] != null ? String(r[COL.id]) : '',
        no,
        date: dateToDisplay(r[COL.date]),
        name: r[COL.name] || '',
        status: '受注',
        qty: 0,
        group: '',
        total: r[COL.total] != null ? r[COL.total] : '',
        note: r[COL.note] || '',
        items: [],
        ordererZip: r[COL.ordererZip] || '',
        ordererPref: '',
        ordererAddr1: r[COL.ordererAddr1] || '',
        ordererAddr2: r[COL.ordererAddr2] || '',
        ordererPhone: r[COL.ordererPhone] || '',
        ordererEmail: '',
        shipName: r[COL.shipName] || '',
        shipZip: r[COL.shipZip] || '',
        shipPref: '',
        shipAddr1: r[COL.shipAddr1] || '',
        shipAddr2: '',
        shipPhone: r[COL.shipPhone] || '',
        billing: r[COL.billing] != null ? r[COL.billing] : '',
        payment: r[COL.payment] || '',
        codFee: r[COL.codFee] != null ? r[COL.codFee] : '',
        postage: r[COL.postage] != null ? r[COL.postage] : '',
        point: r[COL.point] != null ? r[COL.point] : '',
        coupon: (Number(r[COL.couponStore] || 0) + Number(r[COL.couponRakuten] || 0)) || '',
        shipDate: r[COL.shipDate] ? dateToDisplay(r[COL.shipDate]) : '',
      });
      groupSets.set(no, new Set());
    }

    const order = map.get(no);
    order.qty += Number(r[COL.qty]) || 0;

    const productName = r[COL.productName] ? String(r[COL.productName]).trim() : '';
    if (productName) {
      if (productMaster[productName]) groupSets.get(no).add(productMaster[productName]);
      order.items.push({
        name: productName,
        spec: r[COL.spec] || '',
        price: r[COL.price] != null ? r[COL.price] : '',
        qty: r[COL.qty] != null ? r[COL.qty] : '',
        amount: r[COL.amount] != null ? r[COL.amount] : '',
        // この商品の行に直接入力されていた送り先（J～M列）。
        // 空のことが多い（＝注文の送り先と同じ＝1行目の値を使う）が、
        // 値が入っている場合はこの商品だけ別の送り先という意味。
        shipName: r[COL.shipName] || '',
        shipZip: r[COL.shipZip] || '',
        shipAddr1: r[COL.shipAddr1] || '',
        shipPhone: r[COL.shipPhone] || '',
      });
    }
  }

  for (const [no, order] of map.entries()) {
    const gs = groupSets.get(no);
    order.group = gs.has('2') ? '2' : gs.has('1') ? '1' : '';
  }

  return Array.from(map.values());
}

// ---- API ----
app.get('/api/channels', (req, res) => {
  res.json({ channels: CHANNELS });
});

app.get('/api/orders', (req, res) => {
  try {
    const { channel, mode, from, to } = req.query;
    if (!channel) return res.status(400).json({ error: '対象モールを指定してください' });
    if (!from) return res.status(400).json({ error: '検索条件を入力してください' });
    const toVal = to || from;

    const files = findChannelFiles(channel);
    if (files.length === 0) {
      return res.status(404).json({
        error: `「${channel}」を含むExcelファイルが見つかりません。${DATA_DIR} 内にファイルがあるか確認してください。`
      });
    }

    let allRows = [];
    const usedFiles = [];
    for (const f of files) {
      const rows = readSheetRows(path.join(DATA_DIR, f));
      if (rows === null) continue; // 「集計表」シートが無いファイルはスキップ
      usedFiles.push(f);
      allRows = allRows.concat(rows);
    }

    if (usedFiles.length === 0) {
      return res.status(404).json({
        error: `「${channel}」に該当するファイルは見つかりましたが、「${SHEET_NAME}」シートが見つかりませんでした（対象: ${files.join('、')}）`
      });
    }

    let filtered;
    if (mode === 'id') {
      filtered = allRows.filter(r => {
        const v = r[COL.id] != null ? String(r[COL.id]).trim() : '';
        if (!v) return false;
        return naturalCompare(v, from) >= 0 && naturalCompare(v, toVal) <= 0;
      });
    } else {
      filtered = allRows.filter(r => {
        const s = dateToISO(r[COL.date]);
        return s && s >= from && s <= toVal;
      });
    }

    const orders = buildOrders(filtered);
    orders.sort((a, b) => (a.date + a.no).localeCompare(b.date + b.no));

    res.json({ orders, matchedFiles: usedFiles, totalLineRows: filtered.length });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'サーバーエラーが発生しました: ' + err.message });
  }
});

app.listen(PORT, () => {
  console.log(`進捗管理一覧サーバー起動: http://localhost:${PORT}`);
  console.log(`Excel検索フォルダ: ${DATA_DIR}`);
});

// ---- データ永続化 ----
const DATA_FILE = path.join(__dirname, 'orders_data.json');

// 保存済みデータを読み込む
app.get('/api/saved-orders', (req, res) => {
  try {
    if (!fs.existsSync(DATA_FILE)) return res.json({ orders: [] });
    const raw = fs.readFileSync(DATA_FILE, 'utf8');
    res.json(JSON.parse(raw));
  } catch (err) {
    res.json({ orders: [] });
  }
});

// データを保存する
app.post('/api/saved-orders', express.json({ limit: '50mb' }), (req, res) => {
  try {
    fs.writeFileSync(DATA_FILE, JSON.stringify(req.body, null, 2), 'utf8');
    res.json({ ok: true });
  } catch (err) {
    res.status(500).json({ error: '保存に失敗しました: ' + err.message });
  }
});
