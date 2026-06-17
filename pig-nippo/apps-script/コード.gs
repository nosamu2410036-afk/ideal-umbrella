/******************************************************************
 * 養豚場 日報 自動化（Google スプレッドシート版 / Apps Script）
 * ----------------------------------------------------------------
 * 導入方法:
 *   1. スプレッドシートを開く
 *   2. 上部メニュー [拡張機能]→[Apps Script]
 *   3. 既定の Code.gs の中身を全部消し、このファイルの中身を貼り付け
 *   4. 保存（フロッピーアイコン / Ctrl+S）
 *   5. スプレッドシートに戻ってページを再読み込み（F5）
 *   6. 上部に増えた [日報] メニューから各操作を実行
 *      （初回は「承認」を求められるので、自分のアカウントで許可）
 *
 * 機能（[日報] メニュー）:
 *   翌日シート作成 … 最新シートを複製し翌日シートを作成
 *   本日入力       … 死亡/出荷/移動を入力し本日・週・月の累計を更新
 *   月次集計       … 「月次集計」シートに月別集計を出力
 *   チェック       … 在庫マイナス等を「チェック結果」シートに出力
 *
 * セル対応は Excel(VBA)版・Python版と同一。値はすべて結合セルの
 * 左上に入るため、結合セルがあっても問題なく動作します。
 ******************************************************************/

var REIWA_BASE = 2018;        // 令和N年 = 2018 + N
var RATE_THRESHOLD = 8.0;     // 事故率の警告閾値(%)

/** スプレッドシートを開いたときにメニューを追加 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('日報')
    .addItem('翌日シート作成', 'nextDay')
    .addItem('本日入力', 'recordToday')
    .addSeparator()
    .addItem('月次集計', 'monthlyReport')
    .addItem('チェック', 'runCheck')
    .addToUi();
}

/* ===================== 公開操作 ===================== */

function nextDay() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  try {
    var src = newestSheet_(ss);
    if (!src) { ui.alert('日付シートが見つかりません。'); return; }
    var d = nameToDate_(src.getName());
    var nd = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1);
    var nm = dateToName_(nd);
    if (ss.getSheetByName(nm)) { ui.alert('シート ' + nm + ' は既に存在します。'); return; }

    // シート複製（結合セル・書式・メモ・数式をすべて保持）
    var ns = src.copyTo(ss);
    ns.setName(nm);
    ss.setActiveSheet(ns);
    ss.moveActiveSheet(1);   // 先頭へ

    ns.getRange('A1').setValue(nd);
    ns.getRange('C20').setValue(0);
    ns.getRange('C23').setValue(0);
    ns.getRange('C27').setValue(0);

    var msg = 'シート ' + nm + ' を作成しました（元: ' + src.getName() + '）';
    if (weekdayMon_(nd) === 1) {  // 月曜 → 週累計リセット
      ns.getRange('E20').setFormula('=0');
      ns.getRange('E23').setFormula('=0');
      ns.getRange('E27').setFormula('=0');
      msg += '\n・週累計(E列)をリセット（月曜）';
    }
    if (nd.getDate() === 1) {      // 1日 → 月累計リセット
      ns.getRange('G20').setFormula('=0');
      ns.getRange('G23').setFormula('=0');
      ns.getRange('G27').setFormula('=0');
      msg += '\n・月累計(G列)をリセット（月初）';
    }
    ui.alert(msg);
  } catch (e) {
    ui.alert('翌日シート作成でエラー: ' + e.message);
  }
}

function recordToday() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var ws = ss.getActiveSheet();
  if (!nameToDate_(ws.getName())) {
    ui.alert('日付シートを表示してから実行してください（現在: ' + ws.getName() + '）');
    return;
  }
  var ds = promptOrCancel_(ui, '死亡（例: 中1=2,肉4=3）／無ければ空欄'); if (ds === null) return;
  var ss2 = promptOrCancel_(ui, '出荷（例: 肉6=40,肉7=22）／無ければ空欄'); if (ss2 === null) return;
  var ms = promptOrCancel_(ui, '移動（例: 肉6=99）／無ければ空欄'); if (ms === null) return;

  var cd, cs, cm;
  try {
    cd = parseSpec_(ds); cs = parseSpec_(ss2); cm = parseSpec_(ms);
  } catch (e) {
    ui.alert('入力エラー: ' + e.message);
    return;
  }

  var dTot = applySpec_(ws, cd, 'death');
  var sTot = applySpec_(ws, cs, 'ship');
  applySpec_(ws, cm, 'move');

  ws.getRange('C23').setValue(dTot);
  ws.getRange('C20').setValue(sTot);
  if (dTot > 0) { appendTerm_(ws, 'E23', dTot); appendTerm_(ws, 'G23', dTot); }
  if (sTot > 0) { appendTerm_(ws, 'E20', sTot); appendTerm_(ws, 'G20', sTot); }

  ui.alert('記録しました（' + ws.getName() + '）\n本日合計  死亡=' + dTot + '  出荷=' + sTot);
}

function monthlyReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var latest = {};   // "yyyy-mm" -> {name, d}
  ss.getSheets().forEach(function (sh) {
    var d = nameToDate_(sh.getName());
    if (!d) return;
    var key = ym_(d);
    var cur = latest[key];
    if (!cur || d > cur.d ||
        (d.getTime() === cur.d.getTime() && isInventory_(cur.name) && !isInventory_(sh.getName()))) {
      latest[key] = { name: sh.getName(), d: d };
    }
  });
  var keys = Object.keys(latest).sort();
  if (keys.length === 0) { ui.alert('日付シートがありません。'); return; }

  var rep = ss.getSheetByName('月次集計') || ss.insertSheet('月次集計');
  rep.clear();
  rep.getRange(1, 1, rep.getMaxRows(), 14).breakApart();   // 前回の見出し結合を解除

  // 2段見出し：中豚舎 / 肉豚舎 / 累計 をそれぞれ 導入・死亡・出荷・事故率
  var rows = [];
  rows.push(['月', '基準シート', '中豚舎', '', '', '', '肉豚舎', '', '', '', '累計(中+肉)', '', '', '']);
  rows.push(['', '', '導入', '死亡', '出荷', '事故率%', '導入', '死亡', '出荷', '事故率%', '導入', '死亡', '出荷', '事故率%']);
  keys.forEach(function (k) {
    var sh = ss.getSheetByName(latest[k].name);
    var n = houseFigures_(sh, '中');
    var m = houseFigures_(sh, '肉');
    var row = [k, sh.getName()];
    pushSet_(row, n.intro, n.death, n.ship);
    pushSet_(row, m.intro, m.death, m.ship);
    pushSet_(row, n.intro + m.intro, n.death + m.death, n.ship + m.ship);
    rows.push(row);
  });
  rep.getRange(1, 1, rows.length, 14).setValues(rows);
  rep.getRange('A1:A2').merge();
  rep.getRange('B1:B2').merge();
  rep.getRange('C1:F1').merge();
  rep.getRange('G1:J1').merge();
  rep.getRange('K1:N1').merge();
  rep.getRange(1, 1, 2, 14).setFontWeight('bold').setHorizontalAlignment('center');
  rep.autoResizeColumns(1, 14);
  ss.setActiveSheet(rep);
  ui.alert('月次集計（中豚/肉豚/累計）を更新しました（' + keys.length + ' か月）。');
}

/** 各群(中/肉)の導入からの累計：導入頭数・死亡・出荷を8房合計で返す */
function houseFigures_(sh, house) {
  var block = sh.getRange(105, 1, 15, 20).getValues();    // 行105-119, 列A-T
  var fblock = sh.getRange(105, 1, 15, 20).getFormulas();
  var headIdx = (house === '中') ? 2 : 13;   // C / N
  var deathIdx = (house === '中') ? 4 : 15;  // E / P
  var shipIdx = (house === '中') ? 8 : 19;   // I / T
  var intro = 0, death = 0, ship = 0;
  for (var pen = 1; pen <= 8; pen++) {
    var rr = (pen - 1) * 2;
    intro += leadingNum_(fblock[rr][headIdx] || block[rr][headIdx]);  // 頭数式先頭の定数=導入頭数
    death += num_(block[rr][deathIdx]);
    ship += num_(block[rr][shipIdx]);
  }
  return { intro: intro, death: death, ship: ship };
}

function leadingNum_(s) {
  s = String(s);
  if (s.charAt(0) === '=') s = s.substring(1);
  var m = s.match(/^\s*(\d+(?:\.\d+)?)/);
  return m ? parseFloat(m[1]) : 0;
}

function pushSet_(row, intro, death, ship) {
  row.push(intro, death, ship, intro > 0 ? Math.round(death / intro * 10000) / 100 : '');
}

function runCheck() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var dated = [];
  ss.getSheets().forEach(function (sh) {
    var d = nameToDate_(sh.getName());
    if (d) dated.push({ name: sh.getName(), d: d });
  });
  if (dated.length === 0) { ui.alert('日付シートがありません。'); return; }
  dated.sort(function (a, b) { return a.d - b.d; });

  var warns = [];

  // 1) 日付の抜け
  for (var i = 1; i < dated.length; i++) {
    if (dated[i].d.getTime() !== dated[i - 1].d.getTime()) {
      var gap = Math.round((dated[i].d - dated[i - 1].d) / 86400000);
      if (gap > 1) warns.push('日付抜け: ' + dateToName_(dated[i - 1].d) + ' の後 ' + (gap - 1) + ' 日分が欠落');
    }
  }

  // 2) 在庫マイナス（転落/悪化のみ） & 3) 月事故率ピーク
  var prevHead = {};
  var rate = {};   // ym -> {rate, name}
  dated.forEach(function (o) {
    var sh = ss.getSheetByName(o.name);
    var block = sh.getRange(105, 1, 15, 20).getValues();  // 行105-119, 列A-T
    var g = sh.getRange(20, 7, 8, 1).getValues();         // G20-G27
    var deaths = num_(g[3][0]);  // G23
    var intro = num_(g[7][0]);   // G27
    if (!isInventory_(o.name)) {
      [['中', 2], ['肉', 13]].forEach(function (hc) {   // 中=列C(idx2), 肉=列N(idx13)
        var house = hc[0], colIdx = hc[1];
        for (var pen = 1; pen <= 8; pen++) {
          var head = num_(block[(pen - 1) * 2][colIdx]);
          var hk = house + pen;
          if (head < 0 && (!(hk in prevHead) || prevHead[hk] !== head)) {
            warns.push(o.name + ': ' + house + pen + ' の在庫がマイナス (' + head + ')');
          }
          prevHead[hk] = head;
        }
      });
    }
    if (intro > 0) {
      var r = deaths / intro * 100;
      var key = ym_(o.d);
      if (!(key in rate) || r > rate[key].rate) rate[key] = { rate: r, name: o.name };
    }
  });
  Object.keys(rate).sort().forEach(function (key) {
    if (rate[key].rate > RATE_THRESHOLD) {
      warns.push(key + ' 月の事故率 ' + rate[key].rate.toFixed(1) + '% （' +
        rate[key].name + '、閾値 ' + RATE_THRESHOLD + '%超）');
    }
  });

  if (warns.length === 0) {
    ui.alert('問題は検出されませんでした（' + dated.length + ' シート確認）。');
    return;
  }
  var chk = ss.getSheetByName('チェック結果') || ss.insertSheet('チェック結果');
  chk.clear();
  var out = [['チェック結果 ' + new Date().toLocaleString() + '（' + warns.length + ' 件）']];
  warns.forEach(function (w) { out.push([w]); });
  chk.getRange(1, 1, out.length, 1).setValues(out);
  chk.autoResizeColumns(1, 1);
  ss.setActiveSheet(chk);
  ui.alert(warns.length + ' 件の注意があります（「チェック結果」シート参照）。');
}

/* ===================== 補助関数 ===================== */

function penRow_(pen) { return 105 + (pen - 1) * 2; }   // 1→105 … 8→119

function colFor_(house, kind) {
  var map = {
    head: { '中': 'C', '肉': 'N' },
    death: { '中': 'E', '肉': 'P' },
    move: { '中': 'G', '肉': 'R' },
    ship: { '中': 'I', '肉': 'T' }
  };
  return map[kind][house];
}

function appendTerm_(ws, a1, addend) {
  var rng = ws.getRange(a1);
  var f = rng.getFormula();          // 数式なら "=..."、それ以外は ""
  var base;
  if (f) {
    base = (f.charAt(0) === '=') ? f.substring(1) : f;
  } else {
    var v = rng.getValue();
    base = (v === '' || v === null) ? '0' : String(v);
  }
  rng.setFormula('=' + base + '+' + addend);
}

function parseSpec_(spec) {
  var out = [];
  spec = (spec || '').trim();
  if (!spec) return out;
  spec.split(',').forEach(function (item) {
    item = item.trim();
    if (!item) return;
    if (item.indexOf('=') < 0) throw new Error('『房=頭数』形式で入力してください: ' + item);
    var p = item.split('=');
    var key = p[0].trim();
    var m = key.match(/^([中肉])\s*([1-8])$/);
    if (!m) throw new Error('房は 中1〜中8 / 肉1〜肉8 で指定してください: ' + key);
    var n = parseInt((p[1] || '').trim(), 10);
    if (isNaN(n) || n < 0) throw new Error('頭数は0以上の数値で入力してください: ' + item);
    out.push({ house: m[1], pen: parseInt(m[2], 10), n: n });
  });
  return out;
}

function applySpec_(ws, list, kind) {
  var total = 0;
  list.forEach(function (e) {
    if (e.n !== 0) {
      appendTerm_(ws, colFor_(e.house, kind) + penRow_(e.pen), e.n);
      total += e.n;
    }
  });
  return total;
}

function nameToDate_(name) {
  var m = String(name).trim().match(/^R(\d+)\.(\d+)\.(\d+)/);
  if (!m) return null;
  var y = REIWA_BASE + parseInt(m[1], 10);
  var mo = parseInt(m[2], 10), d = parseInt(m[3], 10);
  var dt = new Date(y, mo - 1, d);
  if (dt.getFullYear() !== y || dt.getMonth() !== mo - 1 || dt.getDate() !== d) return null;
  return dt;
}

function dateToName_(dt) {
  return 'R' + (dt.getFullYear() - REIWA_BASE) + '.' + (dt.getMonth() + 1) + '.' + dt.getDate();
}

function isInventory_(name) { return String(name).indexOf('棚卸') >= 0; }

function weekdayMon_(dt) { var w = dt.getDay(); return w === 0 ? 7 : w; }  // 月=1 … 日=7

function ym_(d) { return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2); }

function num_(v) { return (typeof v === 'number') ? v : (parseFloat(v) || 0); }

function newestSheet_(ss) {
  var best = null, bestD = null;
  ss.getSheets().forEach(function (sh) {
    var d = nameToDate_(sh.getName());
    if (d && (!best || d > bestD ||
        (d.getTime() === bestD.getTime() && !isInventory_(sh.getName())))) {
      best = sh; bestD = d;
    }
  });
  return best;
}

function promptOrCancel_(ui, text) {
  var r = ui.prompt(text, ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return null;
  return r.getResponseText();
}
