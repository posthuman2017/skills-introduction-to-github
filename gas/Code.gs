/**
 * 生活介護 送迎表 自動生成 (Google Apps Script)
 *
 * シート構成:
 *   設定          : モード/出発時刻/事業所名/実施日
 *   利用者マスタ  : 氏名 住所 電話 車椅子 医療的ケア 備考 乗車所要分
 *   車両マスタ    : 車両名 定員 リフト 運転手 添乗員
 *   割当          : 車両名 乗車順 氏名 (手動入力)
 *   送迎表        : 自動生成
 */

const SHEETS = {
  CONFIG: '設定',
  USERS: '利用者マスタ',
  VEHICLES: '車両マスタ',
  ASSIGNMENT: '割当',
  RESULT: '送迎表'
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('送迎表')
    .addItem('送迎表を生成', 'generateSchedule')
    .addSeparator()
    .addItem('シート初期化（サンプルデータ）', 'initSheets')
    .addToUi();
}

function initSheets() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert(
    '初期化の確認',
    '既存の「設定」「利用者マスタ」「車両マスタ」「割当」シートを上書きしてサンプルデータを作成します。よろしいですか？',
    ui.ButtonSet.YES_NO
  );
  if (res !== ui.Button.YES) return;

  setupConfigSheet_(ss);
  setupUsersSheet_(ss);
  setupVehiclesSheet_(ss);
  setupAssignmentSheet_(ss);
  ui.alert('初期化が完了しました。各シートの内容を編集してから「送迎表を生成」を実行してください。');
}

function setupConfigSheet_(ss) {
  const sh = ss.getSheetByName(SHEETS.CONFIG) || ss.insertSheet(SHEETS.CONFIG);
  sh.clear();
  const data = [
    ['設定項目', '値', '説明'],
    ['モード', '送り', '「送り」=自宅→事業所 / 「迎え」=事業所→自宅'],
    ['出発時刻', '08:30', 'HH:mm 形式（最初の停車地の予定時刻）'],
    ['事業所名', 'サンプル生活介護センター', '送迎表のヘッダーに表示'],
    ['実施日', '', '空欄なら本日。例: 2026/04/30']
  ];
  sh.getRange(1, 1, data.length, 3).setValues(data);
  sh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#cfe2f3');
  sh.setColumnWidth(1, 140);
  sh.setColumnWidth(2, 220);
  sh.setColumnWidth(3, 360);
  sh.setFrozenRows(1);
}

function setupUsersSheet_(ss) {
  const sh = ss.getSheetByName(SHEETS.USERS) || ss.insertSheet(SHEETS.USERS);
  sh.clear();
  const headers = ['氏名', '住所', '電話', '車椅子', '医療的ケア', '備考', '乗車所要分'];
  const sample = [
    ['山田 太郎', '東京都新宿区西新宿1-1-1', '03-0000-0001', '○', '-', '', 5],
    ['佐藤 花子', '東京都新宿区西新宿2-2-2', '03-0000-0002', '-', '-', '玄関先で待機', 4],
    ['鈴木 一郎', '東京都中野区中野3-3-3', '03-0000-0003', '○', '○', '吸引必要', 6],
    ['田中 次郎', '東京都中野区中野4-4-4', '03-0000-0004', '-', '-', '', 3],
    ['伊藤 三郎', '東京都杉並区高円寺5-5-5', '03-0000-0005', '-', '-', '', 5]
  ];
  const all = [headers].concat(sample);
  sh.getRange(1, 1, all.length, headers.length).setValues(all);
  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#cfe2f3');
  sh.setColumnWidth(1, 110);
  sh.setColumnWidth(2, 280);
  sh.setColumnWidth(3, 130);
  sh.setColumnWidth(4, 70);
  sh.setColumnWidth(5, 90);
  sh.setColumnWidth(6, 200);
  sh.setColumnWidth(7, 100);
  sh.setFrozenRows(1);
}

function setupVehiclesSheet_(ss) {
  const sh = ss.getSheetByName(SHEETS.VEHICLES) || ss.insertSheet(SHEETS.VEHICLES);
  sh.clear();
  const headers = ['車両名', '定員', 'リフト', '運転手', '添乗員'];
  const sample = [
    ['1号車', 8, '○', '高橋', '渡辺'],
    ['2号車', 6, '-', '中村', '小林']
  ];
  const all = [headers].concat(sample);
  sh.getRange(1, 1, all.length, headers.length).setValues(all);
  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#cfe2f3');
  sh.setColumnWidth(1, 100);
  sh.setColumnWidth(2, 60);
  sh.setColumnWidth(3, 70);
  sh.setColumnWidth(4, 100);
  sh.setColumnWidth(5, 100);
  sh.setFrozenRows(1);
}

function setupAssignmentSheet_(ss) {
  const sh = ss.getSheetByName(SHEETS.ASSIGNMENT) || ss.insertSheet(SHEETS.ASSIGNMENT);
  sh.clear();
  const headers = ['車両名', '乗車順', '氏名'];
  const sample = [
    ['1号車', 1, '山田 太郎'],
    ['1号車', 2, '佐藤 花子'],
    ['1号車', 3, '鈴木 一郎'],
    ['2号車', 1, '田中 次郎'],
    ['2号車', 2, '伊藤 三郎']
  ];
  const all = [headers].concat(sample);
  sh.getRange(1, 1, all.length, headers.length).setValues(all);
  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#cfe2f3');
  sh.setColumnWidth(1, 100);
  sh.setColumnWidth(2, 80);
  sh.setColumnWidth(3, 140);
  sh.setFrozenRows(1);
}

function generateSchedule() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  try {
    const config = readConfig_(ss);
    const users = readTable_(ss, SHEETS.USERS);
    const vehicles = readTable_(ss, SHEETS.VEHICLES);
    const assignment = readTable_(ss, SHEETS.ASSIGNMENT);

    const usersByName = {};
    users.forEach(function (u) { usersByName[String(u['氏名']).trim()] = u; });

    const byVehicle = {};
    assignment.forEach(function (a) {
      const v = String(a['車両名']).trim();
      if (!v) return;
      if (!byVehicle[v]) byVehicle[v] = [];
      byVehicle[v].push(a);
    });
    Object.keys(byVehicle).forEach(function (v) {
      byVehicle[v].sort(function (a, b) { return Number(a['乗車順']) - Number(b['乗車順']); });
    });

    validateAssignment_(vehicles, byVehicle, usersByName);
    writeResult_(ss, config, vehicles, byVehicle, usersByName);
    ui.alert('送迎表を生成しました。');
  } catch (e) {
    ui.alert('エラー: ' + e.message);
    throw e;
  }
}

function validateAssignment_(vehicles, byVehicle, usersByName) {
  const vehiclesByName = {};
  vehicles.forEach(function (v) { vehiclesByName[String(v['車両名']).trim()] = v; });

  const messages = [];
  Object.keys(byVehicle).forEach(function (vname) {
    const veh = vehiclesByName[vname];
    if (!veh) {
      messages.push('割当の車両「' + vname + '」が車両マスタに存在しません。');
      return;
    }
    const list = byVehicle[vname];
    if (list.length > Number(veh['定員'])) {
      messages.push('「' + vname + '」の割当人数(' + list.length + ')が定員(' + veh['定員'] + ')を超えています。');
    }
    list.forEach(function (a) {
      const name = String(a['氏名']).trim();
      if (!usersByName[name]) {
        messages.push('割当の氏名「' + name + '」が利用者マスタに存在しません。');
      }
    });
    if (String(veh['リフト']).trim() !== '○') {
      list.forEach(function (a) {
        const u = usersByName[String(a['氏名']).trim()];
        if (u && String(u['車椅子']).trim() === '○') {
          messages.push('「' + vname + '」(リフトなし)に車椅子利用者「' + u['氏名'] + '」が割り当てられています。');
        }
      });
    }
  });
  if (messages.length) throw new Error('\n' + messages.join('\n'));
}

function readConfig_(ss) {
  const sh = ss.getSheetByName(SHEETS.CONFIG);
  if (!sh) throw new Error('「設定」シートがありません。先に「シート初期化」を実行してください。');
  const values = sh.getDataRange().getValues();
  const cfg = {};
  for (let i = 1; i < values.length; i++) cfg[values[i][0]] = values[i][1];
  if (['送り', '迎え'].indexOf(cfg['モード']) < 0) {
    throw new Error('設定の「モード」は「送り」または「迎え」にしてください。');
  }
  if (!cfg['出発時刻']) throw new Error('設定の「出発時刻」が空です。');
  return cfg;
}

function readTable_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('「' + name + '」シートがありません。');
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    if (values[i].every(function (v) { return v === '' || v === null; })) continue;
    const row = {};
    headers.forEach(function (h, j) { row[h] = values[i][j]; });
    rows.push(row);
  }
  return rows;
}

function writeResult_(ss, config, vehicles, byVehicle, usersByName) {
  const sh = ss.getSheetByName(SHEETS.RESULT) || ss.insertSheet(SHEETS.RESULT);
  sh.clear();

  const mode = config['モード'];
  const startTime = parseTime_(config['出発時刻']);
  const facilityName = config['事業所名'] || '';
  const dateStr = formatDate_(config['実施日']);

  const COLS = 8;
  const out = [];
  out.push([facilityName + '  送迎表（' + mode + '）  ' + dateStr, '', '', '', '', '', '', '']);
  out.push(['', '', '', '', '', '', '', '']);

  const mergeRows = [];
  const vehicleHeaderRows = [];
  const tableHeaderRows = [];

  vehicles.forEach(function (v) {
    const vname = String(v['車両名']).trim();
    const list = byVehicle[vname] || [];

    vehicleHeaderRows.push(out.length);
    out.push([
      '【' + vname + '】 定員:' + v['定員'] + '  リフト:' + v['リフト'] +
        '  運転手:' + v['運転手'] + '  添乗員:' + v['添乗員'],
      '', '', '', '', '', '', ''
    ]);

    tableHeaderRows.push(out.length);
    out.push(['順', '時刻', '氏名', '住所', '電話', '車椅子', '医療的ケア', '備考']);

    let cumMin = 0;
    list.forEach(function (a, idx) {
      const u = usersByName[String(a['氏名']).trim()];
      const dur = u ? (Number(u['乗車所要分']) || 0) : 0;
      cumMin += dur;
      const t = addMinutes_(startTime, cumMin);
      out.push([
        idx + 1,
        formatTime_(t),
        u ? u['氏名'] : a['氏名'],
        u ? u['住所'] : '',
        u ? u['電話'] : '',
        u ? u['車椅子'] : '',
        u ? u['医療的ケア'] : '',
        u ? u['備考'] : ''
      ]);
    });

    out.push(['', '', '', '', '', '', '', '']);
  });

  mergeRows.push(0);

  sh.getRange(1, 1, out.length, COLS).setValues(out);

  sh.getRange(1, 1, 1, COLS).merge()
    .setFontSize(14).setFontWeight('bold').setBackground('#fff2cc')
    .setHorizontalAlignment('center');

  vehicleHeaderRows.forEach(function (r) {
    sh.getRange(r + 1, 1, 1, COLS).merge()
      .setFontWeight('bold').setBackground('#d9ead3');
  });
  tableHeaderRows.forEach(function (r) {
    sh.getRange(r + 1, 1, 1, COLS)
      .setFontWeight('bold').setBackground('#cfe2f3').setHorizontalAlignment('center');
  });

  sh.setColumnWidth(1, 40);
  sh.setColumnWidth(2, 70);
  sh.setColumnWidth(3, 110);
  sh.setColumnWidth(4, 260);
  sh.setColumnWidth(5, 120);
  sh.setColumnWidth(6, 60);
  sh.setColumnWidth(7, 80);
  sh.setColumnWidth(8, 180);

  sh.getRange(1, 1, out.length, COLS).setVerticalAlignment('middle');
  sh.setHiddenGridlines(false);

  ss.setActiveSheet(sh);
}

function parseTime_(s) {
  if (s instanceof Date) return { h: s.getHours(), m: s.getMinutes() };
  const m = String(s).match(/^(\d{1,2}):(\d{2})$/);
  if (!m) throw new Error('時刻形式が不正です（例 08:30）: ' + s);
  return { h: Number(m[1]), m: Number(m[2]) };
}

function addMinutes_(t, min) {
  const total = t.h * 60 + t.m + min;
  return { h: ((Math.floor(total / 60) % 24) + 24) % 24, m: ((total % 60) + 60) % 60 };
}

function formatTime_(t) {
  const pad = function (n) { return String(n).padStart(2, '0'); };
  return pad(t.h) + ':' + pad(t.m);
}

function formatDate_(v) {
  let d;
  if (v instanceof Date) d = v;
  else if (v) d = new Date(v);
  else d = new Date();
  if (isNaN(d.getTime())) d = new Date();
  return Utilities.formatDate(d, Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyy年M月d日');
}
