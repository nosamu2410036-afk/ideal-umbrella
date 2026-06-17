#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""養豚場 日報（中豚舎・肉豚舎）Excel 自動化ツール.

1つの xlsx に「1日1シート」で記録される養豚場の日報を自動化する。

サブコマンド:
  next-day  最新シートを複製して翌日シートを作成（日付更新・本日欄リセット・週/月累計リセット）
  record    本日の死亡・出荷・移動頭数を各房に追記し、本日/週/月の累計を更新
  report    全シートを横断して月次集計（死亡・出荷・導入・事故率）を出力
  check     入力漏れ・在庫マイナス・事故率異常・日付抜けを検出

設計方針:
  Excel が付与するコメント枠(VML)・印刷設定・書式を壊さないよう、openpyxl での
  全体再保存は行わず、ZIP/XML を最小限だけ書き換える。読み取り(report/check)は
  数式を自前評価するため、Excel で開き直す前でも正しい値を集計できる。
"""
from __future__ import annotations

import argparse
import datetime as _dt
import os
import re
import shutil
import sys
import zipfile
from typing import Dict, List, Optional, Tuple

# ---------------------------------------------------------------------------
# 房 → セル対応表
# ---------------------------------------------------------------------------
# 中豚舎 1..8 / 肉豚舎 1..8 はそれぞれ行 105,107,...,119 に並ぶ。
PEN_ROWS = {i: 105 + (i - 1) * 2 for i in range(1, 9)}  # {1:105, 2:107, ... 8:119}

# 列: 頭数 / 死亡 / 移動 / 出荷
HOUSE_COLS = {
    "中": {"head": "C", "death": "E", "move": "G", "ship": "I"},
    "肉": {"head": "N", "death": "P", "move": "R", "ship": "T"},
}

# 上部サマリー欄
TODAY = {"ship": "C20", "death": "C23", "intro": "C27"}     # 本日
WEEK = {"ship": "E20", "death": "E23", "intro": "E27"}      # 週累計
MONTH = {"ship": "G20", "death": "G23", "intro": "G27"}     # 月累計

DATE_CELL = "A1"

EPOCH = _dt.date(1899, 12, 30)  # Excel 1900 日付システムの基準日


def to_serial(d: _dt.date) -> int:
    return (d - EPOCH).days


def from_serial(n: int) -> _dt.date:
    return EPOCH + _dt.timedelta(days=int(n))


# ---------------------------------------------------------------------------
# シート名 <-> 日付 (令和)
# ---------------------------------------------------------------------------
REIWA_BASE = 2018  # 令和N年 = 2018 + N
_NAME_RE = re.compile(r"^R(\d+)\.(\d+)\.(\d+)")


def name_to_date(name: str) -> Optional[_dt.date]:
    m = _NAME_RE.match(name.strip())
    if not m:
        return None
    e, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    try:
        return _dt.date(REIWA_BASE + e, mo, d)
    except ValueError:
        return None


def date_to_name(d: _dt.date) -> str:
    return "R%d.%d.%d" % (d.year - REIWA_BASE, d.month, d.day)


def is_inventory(name: str) -> bool:
    """棚卸シートか（全角/半角括弧どちらにも対応）。"""
    return "棚卸" in name


# ---------------------------------------------------------------------------
# 房指定文字列のパース  例: "中1=2,肉4=3"
# ---------------------------------------------------------------------------
def parse_pen_spec(spec: str) -> Dict[Tuple[str, int], int]:
    out: Dict[Tuple[str, int], int] = {}
    if not spec:
        return out
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "=" not in part:
            raise ValueError("『房=頭数』の形式で指定してください: %r" % part)
        key, val = part.split("=", 1)
        key = key.strip()
        m = re.match(r"^([中肉])\s*([1-8])$", key)
        if not m:
            raise ValueError("房は 中1〜中8 / 肉1〜肉8 で指定してください: %r" % key)
        house, pen = m.group(1), int(m.group(2))
        n = int(val.strip())
        if n < 0:
            raise ValueError("頭数は0以上で指定してください: %r" % part)
        out[(house, pen)] = out.get((house, pen), 0) + n
    return out


def pen_cell(house: str, pen: int, kind: str) -> str:
    return HOUSE_COLS[house][kind] + str(PEN_ROWS[pen])


# ---------------------------------------------------------------------------
# 簡易数式エバリュエータ（report / check 用）
#   対応: 数値, + - * /, 括弧, セル参照(A1形式), SUM(範囲[,範囲...])
# ---------------------------------------------------------------------------
_REF_RE = re.compile(r"\$?([A-Z]{1,3})\$?(\d+)")
_RANGE_RE = re.compile(r"\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)")


def col_to_num(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n


def num_to_col(n: int) -> str:
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(ord("A") + r) + s
    return s


class FormulaSheet:
    """1シートの全セル値/数式を保持し、参照解決つきで評価する。"""

    def __init__(self, cells: Dict[str, Tuple[Optional[str], Optional[str]]]):
        # cells[ref] = (formula_text_or_None, raw_value_or_None)
        self.cells = cells
        self._cache: Dict[str, float] = {}
        self._stack: set = set()

    def eval_ref(self, ref: str) -> float:
        if ref in self._cache:
            return self._cache[ref]
        if ref in self._stack:  # 循環参照ガード
            return 0.0
        f, v = self.cells.get(ref, (None, None))
        self._stack.add(ref)
        try:
            if f is not None:
                val = self.eval_formula(f)
            elif v is not None and v != "":
                try:
                    val = float(v)
                except ValueError:
                    val = 0.0  # 文字列セル
            else:
                val = 0.0
        finally:
            self._stack.discard(ref)
        self._cache[ref] = val
        return val

    def eval_formula(self, formula: str) -> float:
        expr = formula

        # SUM(range,...) を展開
        def repl_sum(m: "re.Match") -> str:
            inner = m.group(1)
            total = 0.0
            for token in inner.split(","):
                token = token.strip()
                rm = _RANGE_RE.fullmatch(token)
                if rm:
                    c1, r1, c2, r2 = rm.groups()
                    a1, a2 = col_to_num(c1), col_to_num(c2)
                    b1, b2 = int(r1), int(r2)
                    for col in range(min(a1, a2), max(a1, a2) + 1):
                        for row in range(min(b1, b2), max(b1, b2) + 1):
                            total += self.eval_ref(num_to_col(col) + str(row))
                elif _REF_RE.fullmatch(token):
                    total += self.eval_ref(token.replace("$", ""))
                elif token:
                    try:
                        total += float(token)
                    except ValueError:
                        pass
            return repr(total)

        expr = re.sub(r"SUM\(([^)]*)\)", repl_sum, expr, flags=re.I)
        # 単独セル参照を値に置換
        expr = _REF_RE.sub(lambda m: repr(self.eval_ref(m.group(1) + m.group(2))), expr)

        if not re.fullmatch(r"[0-9+\-*/.()eE\s]*", expr):
            return 0.0  # 未対応関数等は0扱い
        if expr.strip() == "":
            return 0.0
        try:
            return float(eval(expr, {"__builtins__": {}}, {}))  # noqa: S307 (式は数値のみに限定済み)
        except Exception:
            return 0.0


# ---------------------------------------------------------------------------
# Xlsx: ZIP/XML を最小編集で扱う
# ---------------------------------------------------------------------------
WORKSHEET_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
COMMENTS_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"


class Xlsx:
    def __init__(self, path: str):
        self.path = path
        self.parts: Dict[str, bytes] = {}
        self.order: List[str] = []
        with zipfile.ZipFile(path) as z:
            for info in z.infolist():
                self.order.append(info.filename)
                self.parts[info.filename] = z.read(info.filename)
        self.dirty = False

    # --- テキストパート読み書き -------------------------------------------
    def _text(self, name: str) -> str:
        return self.parts[name].decode("utf-8")

    def _set_text(self, name: str, text: str) -> None:
        self.parts[name] = text.encode("utf-8")
        self.dirty = True

    # --- ワークブック構造 --------------------------------------------------
    def sheet_list(self) -> List[Tuple[str, str, str]]:
        """[(name, sheetId, rId)] を workbook.xml の並び順で返す。"""
        wb = self._text("xl/workbook.xml")
        out = []
        for m in re.finditer(r"<sheet\b[^>]*?/>", wb):
            tag = m.group(0)
            name = re.search(r'name="([^"]*)"', tag)
            sid = re.search(r'sheetId="(\d+)"', tag)
            rid = re.search(r'r:id="(rId\d+)"', tag)
            if name and rid:
                out.append((name.group(1), sid.group(1) if sid else "", rid.group(1)))
        return out

    def _rid_to_target(self) -> Dict[str, str]:
        rels = self._text("xl/_rels/workbook.xml.rels")
        out = {}
        for m in re.finditer(r"<Relationship\b[^>]*/>", rels):
            tag = m.group(0)
            rid = re.search(r'Id="(rId\d+)"', tag)
            tgt = re.search(r'Target="([^"]*)"', tag)
            if rid and tgt:
                out[rid.group(1)] = tgt.group(1)
        return out

    def sheet_part(self, sheet_name: str) -> str:
        for name, _sid, rid in self.sheet_list():
            if name == sheet_name:
                target = self._rid_to_target()[rid]
                target = target.lstrip("/")
                if not target.startswith("xl/"):
                    target = "xl/" + target
                return target
        raise KeyError("シートが見つかりません: %s" % sheet_name)

    def newest_sheet(self) -> str:
        """日付が最新（=並びの先頭側で最大日付）のシート名。"""
        best: Optional[Tuple[_dt.date, str]] = None
        for name, _sid, _rid in self.sheet_list():
            d = name_to_date(name)
            if d is None:
                continue
            if best is None or d > best[0] or (d == best[0] and not is_inventory(name)):
                best = (d, name)
        if best is None:
            raise RuntimeError("日付シートが見つかりません")
        return best[1]

    # --- セル操作 ----------------------------------------------------------
    @staticmethod
    def _cell_re(ref: str) -> "re.Pattern":
        return re.compile(
            r'<c r="%s"(?P<attr>[^>]*?)\s*(?:/>|>(?P<body>.*?)</c>)' % re.escape(ref),
            re.S,
        )

    @staticmethod
    def _style_attr(attr: str) -> str:
        m = re.search(r'\ss="\d+"', attr)
        return m.group(0) if m else ""

    def read_cell(self, part: str, ref: str) -> Tuple[Optional[str], Optional[str]]:
        """(formula_text, raw_value) を返す。formula は先頭 '=' なし。"""
        xml = self._text(part)
        m = self._cell_re(ref).search(xml)
        if not m:
            return (None, None)
        body = m.group("body") or ""
        f = re.search(r"<f[^>]*>(.*?)</f>", body, re.S)
        v = re.search(r"<v[^>]*>(.*?)</v>", body, re.S)
        return (f.group(1) if f else None, v.group(1) if v else None)

    def _write_cell(self, part: str, ref: str, inner: str) -> None:
        xml = self._text(part)
        rx = self._cell_re(ref)
        m = rx.search(xml)
        if not m:
            raise KeyError("セル %s が %s に存在しません" % (ref, part))
        style = self._style_attr(m.group("attr") or "")
        new = '<c r="%s"%s>%s</c>' % (ref, style, inner)
        self._set_text(part, xml[: m.start()] + new + xml[m.end():])

    def set_number(self, part: str, ref: str, value) -> None:
        self._write_cell(part, ref, "<v>%s</v>" % _fmt_num(value))

    def set_formula(self, part: str, ref: str, formula: str) -> None:
        # キャッシュ値(<v>)は付けない → Excel が開いたとき再計算する
        self._write_cell(part, ref, "<f>%s</f>" % formula)

    def append_term(self, part: str, ref: str, addend) -> None:
        """累計セルに『+addend』を追記。数式でなければ数式化する。"""
        f, v = self.read_cell(part, ref)
        base = f if f is not None else (v if (v not in (None, "")) else "0")
        self.set_formula(part, ref, "%s+%s" % (base, _fmt_num(addend)))

    def reset_formula(self, part: str, ref: str) -> None:
        self.set_formula(part, ref, "0")

    # --- 再計算の強制 ------------------------------------------------------
    def force_recalc(self) -> None:
        wb = self._text("xl/workbook.xml")
        if "fullCalcOnLoad" in wb:
            return
        if "<calcPr" in wb:
            wb = re.sub(r"<calcPr\b([^>]*?)/>", r'<calcPr\1 fullCalcOnLoad="1"/>', wb, count=1)
        else:
            wb = wb.replace("</workbook>", '<calcPr calcId="0" fullCalcOnLoad="1"/></workbook>')
        self._set_text("xl/workbook.xml", wb)
        # calcChain は再計算で再生成されるため削除
        self.parts.pop("xl/calcChain.xml", None)
        ct = self._text("[Content_Types].xml")
        ct = ct.replace('<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>', "")
        self._set_text("[Content_Types].xml", ct)

    # --- シート複製 --------------------------------------------------------
    def _max_index(self, prefix: str, suffix: str) -> int:
        rx = re.compile(re.escape(prefix) + r"(\d+)" + re.escape(suffix))
        mx = 0
        for name in self.parts:
            m = rx.fullmatch(name)
            if m:
                mx = max(mx, int(m.group(1)))
        return mx

    def clone_sheet(self, src_name: str, new_name: str) -> str:
        src_part = self.sheet_part(src_name)             # xl/worksheets/sheetN.xml
        new_sheet_idx = self._max_index("xl/worksheets/sheet", ".xml") + 1
        new_part = "xl/worksheets/sheet%d.xml" % new_sheet_idx

        # シート本体をコピー
        self.parts[new_part] = self.parts[src_part]
        self.order.append(new_part)

        # 関連パート（コメント / VML / 印刷設定）をコピー
        src_rels_name = re.sub(r"(sheet\d+\.xml)$", r"_rels/\1.rels", src_part)
        new_rels = None
        if src_rels_name in self.parts:
            rels = self._text(src_rels_name)
            new_rels = rels
            for m in re.finditer(r"<Relationship\b[^>]*/>", rels):
                tag = m.group(0)
                typ = re.search(r'Type="[^"]*/(\w+)"', tag).group(1)
                tgt = re.search(r'Target="([^"]*)"', tag).group(1)
                src_rel_part = _resolve_rel(src_part, tgt)
                if typ == "comments":
                    idx = self._max_index("xl/comments", ".xml") + 1
                    dst = "xl/comments%d.xml" % idx
                    new_tgt = "../comments%d.xml" % idx
                    self._add_override(dst, COMMENTS_CT)
                elif typ == "vmlDrawing":
                    idx = self._max_index("xl/drawings/vmlDrawing", ".vml") + 1
                    dst = "xl/drawings/vmlDrawing%d.vml" % idx
                    new_tgt = "../drawings/vmlDrawing%d.vml" % idx
                elif typ == "printerSettings":
                    idx = self._max_index("xl/printerSettings/printerSettings", ".bin") + 1
                    dst = "xl/printerSettings/printerSettings%d.bin" % idx
                    new_tgt = "../printerSettings/printerSettings%d.bin" % idx
                else:
                    continue
                self.parts[dst] = self.parts[src_rel_part]
                self.order.append(dst)
                new_rels = new_rels.replace('Target="%s"' % tgt, 'Target="%s"' % new_tgt)
            new_rels_name = "xl/worksheets/_rels/sheet%d.xml.rels" % new_sheet_idx
            self.parts[new_rels_name] = new_rels.encode("utf-8")
            self.order.append(new_rels_name)

        # Content_Types に新シートを登録
        self._add_override(new_part, WORKSHEET_CT)

        # workbook.xml.rels に新リレーション
        rels = self._text("xl/_rels/workbook.xml.rels")
        max_rid = max(int(r[3:]) for r in re.findall(r'Id="(rId\d+)"', rels))
        new_rid = "rId%d" % (max_rid + 1)
        new_rel = ('<Relationship Id="%s" '
                   'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                   'Target="worksheets/sheet%d.xml"/>') % (new_rid, new_sheet_idx)
        rels = rels.replace("</Relationships>", new_rel + "</Relationships>")
        self._set_text("xl/_rels/workbook.xml.rels", rels)

        # workbook.xml に <sheet> を先頭挿入（最新を先頭に）
        wb = self._text("xl/workbook.xml")
        max_sid = max(int(s) for _n, s, _r in self.sheet_list() if s)
        new_sheet_tag = '<sheet name="%s" sheetId="%d" r:id="%s"/>' % (
            new_name, max_sid + 1, new_rid)
        wb = re.sub(r"(<sheets>)", r"\1" + new_sheet_tag, wb, count=1)
        self._set_text("xl/workbook.xml", wb)

        self.dirty = True
        return new_part

    def _add_override(self, part_name: str, content_type: str) -> None:
        ct = self._text("[Content_Types].xml")
        ov = '<Override PartName="/%s" ContentType="%s"/>' % (part_name, content_type)
        ct = ct.replace("</Types>", ov + "</Types>")
        self._set_text("[Content_Types].xml", ct)

    # --- 保存 --------------------------------------------------------------
    def save(self, out_path: str) -> None:
        tmp = out_path + ".tmp"
        written = set()
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            for name in self.order:
                if name in written or name not in self.parts:
                    continue
                z.writestr(name, self.parts[name])
                written.add(name)
        os.replace(tmp, out_path)


def _fmt_num(value) -> str:
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _resolve_rel(src_part: str, target: str) -> str:
    """シートパート基準の相対 Target を ZIP 内パス(xl/...)へ。"""
    base = os.path.dirname(src_part)  # xl/worksheets
    joined = os.path.normpath(os.path.join(base, target))
    return joined.replace(os.sep, "/")


# ---------------------------------------------------------------------------
# サブコマンド実装
# ---------------------------------------------------------------------------
def load_sheet_cells(x: Xlsx, sheet_name: str) -> FormulaSheet:
    part = x.sheet_part(sheet_name)
    xml = x._text(part)
    cells: Dict[str, Tuple[Optional[str], Optional[str]]] = {}
    for m in re.finditer(r'<c r="([A-Z]+\d+)"(?P<attr>[^>]*?)\s*(?:/>|>(?P<body>.*?)</c>)', xml, re.S):
        ref = m.group(1)
        body = m.group("body") or ""
        f = re.search(r"<f[^>]*>(.*?)</f>", body, re.S)
        v = re.search(r"<v[^>]*>(.*?)</v>", body, re.S)
        cells[ref] = (f.group(1) if f else None, v.group(1) if v else None)
    return FormulaSheet(cells)


def cmd_next_day(args) -> int:
    x = Xlsx(args.file)
    src = args.from_sheet or x.newest_sheet()
    src_date = name_to_date(src)
    if src_date is None:
        print("元シートの日付を解釈できません: %s" % src, file=sys.stderr)
        return 2
    new_date = _dt.date.fromisoformat(args.date) if args.date else src_date + _dt.timedelta(days=1)
    new_name = date_to_name(new_date)
    if any(n == new_name for n, _s, _r in x.sheet_list()):
        print("シート %s は既に存在します" % new_name, file=sys.stderr)
        return 2

    new_part = x.clone_sheet(src, new_name)

    # 日付
    x.set_number(new_part, DATE_CELL, to_serial(new_date))
    # 本日欄リセット
    for ref in TODAY.values():
        x.set_number(new_part, ref, 0)
    # 週初(月曜)に週累計リセット
    week_reset = (new_date.weekday() == 0) and not args.no_week_reset
    if week_reset:
        for ref in WEEK.values():
            x.reset_formula(new_part, ref)
    # 月初(1日)に月累計リセット
    month_reset = (new_date.day == 1) and not args.no_month_reset
    if month_reset:
        for ref in MONTH.values():
            x.reset_formula(new_part, ref)

    x.force_recalc()
    out = args.out or args.file
    x.save(out)
    print("作成: シート『%s』を『%s』から複製しました → %s" % (new_name, src, out))
    if week_reset:
        print("  週累計(E列)をリセットしました（月曜のため）")
    if month_reset:
        print("  月累計(G列)をリセットしました（月初のため）")
    return 0


def cmd_record(args) -> int:
    x = Xlsx(args.file)
    sheet = args.sheet or x.newest_sheet()
    part = x.sheet_part(sheet)

    deaths = parse_pen_spec(args.death or "")
    ships = parse_pen_spec(args.ship or "")
    moves = parse_pen_spec(args.move or "")

    log: List[str] = []

    # 各房へ追記
    for (house, pen), n in deaths.items():
        if n:
            x.append_term(part, pen_cell(house, pen, "death"), n)
            log.append("死亡 %s%d +%d" % (house, pen, n))
    for (house, pen), n in ships.items():
        if n:
            x.append_term(part, pen_cell(house, pen, "ship"), n)
            log.append("出荷 %s%d +%d" % (house, pen, n))
    for (house, pen), n in moves.items():
        if n:
            x.append_term(part, pen_cell(house, pen, "move"), n)
            log.append("移動 %s%d +%d" % (house, pen, n))

    death_total = sum(deaths.values())
    ship_total = sum(ships.values())
    intro_total = args.intro or 0

    # 本日欄
    x.set_number(part, TODAY["death"], death_total)
    x.set_number(part, TODAY["ship"], ship_total)
    if args.intro is not None:
        x.set_number(part, TODAY["intro"], intro_total)

    # 週・月累計へ本日合計を追記
    if death_total:
        x.append_term(part, WEEK["death"], death_total)
        x.append_term(part, MONTH["death"], death_total)
    if ship_total:
        x.append_term(part, WEEK["ship"], ship_total)
        x.append_term(part, MONTH["ship"], ship_total)
    if intro_total:
        x.append_term(part, WEEK["intro"], intro_total)
        x.append_term(part, MONTH["intro"], intro_total)

    x.force_recalc()
    out = args.out or args.file
    x.save(out)
    print("記録: シート『%s』" % sheet)
    for line in log:
        print("  " + line)
    print("  本日合計  死亡=%d 出荷=%d 導入=%d" % (death_total, ship_total, intro_total))
    print("  → %s" % out)
    return 0


def _month_key(d: _dt.date) -> str:
    return "%04d-%02d" % (d.year, d.month)


def house_figures(fs: "FormulaSheet", house: str) -> Tuple[float, float, float]:
    """各群(中/肉)の導入からの累計（導入頭数・死亡・出荷）を 8 房合計で返す。

    導入頭数は頭数セルの数式先頭の定数（例 ``430-E105-...`` の 430）。
    死亡・出荷は各房の累計値。
    """
    intro = death = ship = 0.0
    for pen in range(1, 9):
        f, v = fs.cells.get(pen_cell(house, pen, "head"), (None, None))
        if f is not None:
            m = re.match(r"\s*(\d+(?:\.\d+)?)", f)
            if m:
                intro += float(m.group(1))
        elif v not in (None, ""):
            try:
                intro += float(v)
            except ValueError:
                pass
        death += fs.eval_ref(pen_cell(house, pen, "death"))
        ship += fs.eval_ref(pen_cell(house, pen, "ship"))
    return intro, death, ship


def cmd_report(args) -> int:
    x = Xlsx(args.file)
    # 月ごとに最新の通常シート（棚卸でない最終日）を採用
    latest_per_month: Dict[str, Tuple[_dt.date, str]] = {}
    for name, _s, _r in x.sheet_list():
        d = name_to_date(name)
        if d is None:
            continue
        key = _month_key(d)
        cur = latest_per_month.get(key)
        # 同日なら棚卸でない方を優先、より新しい日付を優先
        if cur is None or d > cur[0] or (d == cur[0] and is_inventory(cur[1]) and not is_inventory(name)):
            latest_per_month[key] = (d, name)

    def rate(death: float, intro: float) -> float:
        return (death / intro * 100) if intro else 0.0

    rows = []
    for key in sorted(latest_per_month):
        _d, name = latest_per_month[key]
        fs = load_sheet_cells(x, name)
        ni, nd, ns = house_figures(fs, "中")
        mi, md, ms = house_figures(fs, "肉")
        rows.append((key, name, (ni, nd, ns), (mi, md, ms), (ni + mi, nd + md, ns + ms)))

    # 表示（中豚舎 / 肉豚舎 / 累計、それぞれ 導入・死亡・出荷・事故率）
    print("月次集計（各月末シート基準・導入からの累計）")
    header = "%-9s %-14s | %-22s | %-22s | %-22s" % (
        "月", "基準シート", "中豚舎(導入/死亡/出荷/率%)",
        "肉豚舎(導入/死亡/出荷/率%)", "累計(導入/死亡/出荷/率%)")
    print(header)
    print("-" * len(header))

    def fmt(group: Tuple[float, float, float]) -> str:
        i, dth, sh = group
        return "%5d %5d %5d %6.2f" % (i, dth, sh, rate(dth, i))

    for key, name, n, m, t in rows:
        print("%-9s %-14s | %s | %s | %s" % (key, name, fmt(n), fmt(m), fmt(t)))

    if args.csv:
        import csv
        with open(args.csv, "w", encoding="utf-8-sig", newline="") as fp:
            w = csv.writer(fp)
            w.writerow(["月", "基準シート",
                        "中豚_導入", "中豚_死亡", "中豚_出荷", "中豚_事故率%",
                        "肉豚_導入", "肉豚_死亡", "肉豚_出荷", "肉豚_事故率%",
                        "累計_導入", "累計_死亡", "累計_出荷", "累計_事故率%"])
            for key, name, n, m, t in rows:
                w.writerow([key, name,
                            int(n[0]), int(n[1]), int(n[2]), round(rate(n[1], n[0]), 2),
                            int(m[0]), int(m[1]), int(m[2]), round(rate(m[1], m[0]), 2),
                            int(t[0]), int(t[1]), int(t[2]), round(rate(t[1], t[0]), 2)])
        print("\nCSV出力: %s" % args.csv)
    return 0


NINE_DEATH = "N124"  # 9豚舎の死亡累計（本日死亡C23にも算入される。月初リセット）


def _death_cells() -> List[str]:
    cells = [pen_cell(h, p, "death") for h in ("中", "肉") for p in range(1, 9)]
    cells.append(NINE_DEATH)
    return cells


def cmd_check(args) -> int:
    x = Xlsx(args.file)
    dated = []
    for name, _s, _r in x.sheet_list():
        d = name_to_date(name)
        if d is not None:
            dated.append((d, name))
    dated.sort()

    warnings: List[str] = []

    # 1) 日付の抜け（棚卸の重複日は無視）
    seen_dates = sorted({d for d, _n in dated})
    for prev, nxt in zip(seen_dates, seen_dates[1:]):
        gap = (nxt - prev).days
        if gap > 1:
            warnings.append("日付抜け: %s の後 %d 日分が欠落（次は %s）"
                            % (date_to_name(prev), gap - 1, date_to_name(nxt)))

    # 2) 在庫マイナス（マイナスへ転落／悪化した日のみ報告）
    prev_head: Dict[Tuple[str, int], float] = {}
    for d, name in dated:
        if is_inventory(name):
            continue
        fs = load_sheet_cells(x, name)
        for house in ("中", "肉"):
            for pen in range(1, 9):
                head = fs.eval_ref(pen_cell(house, pen, "head"))
                key = (house, pen)
                was = prev_head.get(key)
                if head < 0 and (was is None or head != was):
                    warnings.append("%s: %s%d の在庫がマイナス (%d)" % (name, house, pen, int(head)))
                prev_head[key] = head

    # 3) 月事故率異常（月ごとに最大値を1回だけ報告）
    threshold = args.rate_threshold
    month_peak: Dict[str, Tuple[float, str]] = {}
    for d, name in dated:
        fs = load_sheet_cells(x, name)
        intro = fs.eval_ref(MONTH["intro"])
        deaths = fs.eval_ref(MONTH["death"])
        if intro:
            rate = deaths / intro * 100
            key = _month_key(d)
            if key not in month_peak or rate > month_peak[key][0]:
                month_peak[key] = (rate, name)
    for key in sorted(month_peak):
        rate, name = month_peak[key]
        if rate > threshold:
            warnings.append("%s 月の事故率が %.1f%% （%s 時点、閾値 %.1f%% 超）"
                            % (key, rate, name, threshold))

    # 4) 本日死亡(C23) と 各房+9豚舎 死亡の前日差の整合
    #    バッチ入替や月初リセットで死亡累計が減少した日は判定不能のため除外
    by_date = {}
    for d, name in dated:
        if not is_inventory(name):
            by_date.setdefault(d, name)
    ordered = sorted(by_date)
    death_cells = _death_cells()
    for prev, nxt in zip(ordered, ordered[1:]):
        if (nxt - prev).days != 1:
            continue
        fp = load_sheet_cells(x, by_date[prev])
        fn = load_sheet_cells(x, by_date[nxt])
        delta = 0.0
        had_reset = False
        for cell in death_cells:
            inc = fn.eval_ref(cell) - fp.eval_ref(cell)
            if inc < 0:
                had_reset = True
                break
            delta += inc
        if had_reset:
            continue
        c23 = fn.eval_ref(TODAY["death"])
        if abs(delta - c23) > 0.5:
            warnings.append("%s: 本日死亡(C23=%d) と 各房+9豚舎の死亡増分(%d) が不一致"
                            % (by_date[nxt], int(c23), int(delta)))

    if warnings:
        print("⚠ %d 件の注意:" % len(warnings))
        for w in warnings:
            print("  - " + w)
        return 1
    print("✓ 問題は検出されませんでした（%d シート確認）" % len(dated))
    return 0


# ---------------------------------------------------------------------------
def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="養豚場 日報 Excel 自動化ツール")
    p.add_argument("--file", "-f", required=True, help="対象の xlsx ファイル")
    sub = p.add_subparsers(dest="cmd", required=True)

    pn = sub.add_parser("next-day", help="翌日シートを作成")
    pn.add_argument("--date", help="作成する日付 YYYY-MM-DD（省略時は最新+1日）")
    pn.add_argument("--from-sheet", help="複製元シート名（省略時は最新シート）")
    pn.add_argument("--no-week-reset", action="store_true", help="月曜でも週累計をリセットしない")
    pn.add_argument("--no-month-reset", action="store_true", help="月初でも月累計をリセットしない")
    pn.add_argument("--out", help="出力先（省略時は上書き）")
    pn.set_defaults(func=cmd_next_day)

    pr = sub.add_parser("record", help="本日の死亡/出荷/移動を記録")
    pr.add_argument("--sheet", help="対象シート名（省略時は最新）")
    pr.add_argument("--death", help='死亡 例: "中1=2,肉4=3"')
    pr.add_argument("--ship", help='出荷 例: "肉6=40,肉7=22"')
    pr.add_argument("--move", help='移動 例: "肉6=99"')
    pr.add_argument("--intro", type=int, help="本日導入頭数（合計）")
    pr.add_argument("--out", help="出力先（省略時は上書き）")
    pr.set_defaults(func=cmd_record)

    prep = sub.add_parser("report", help="月次集計")
    prep.add_argument("--csv", help="CSV 出力先")
    prep.set_defaults(func=cmd_report)

    pc = sub.add_parser("check", help="入力チェック")
    pc.add_argument("--rate-threshold", type=float, default=8.0, help="事故率の警告閾値%%（既定8.0）")
    pc.set_defaults(func=cmd_check)
    return p


def main(argv=None) -> int:
    args = build_parser().parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
