#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""nippo.py のユニットテスト（xlsx 不要の純粋ロジック部分）。

実行: python3 test_nippo.py
"""
import datetime
import unittest

import nippo


class TestDateName(unittest.TestCase):
    def test_name_to_date(self):
        self.assertEqual(nippo.name_to_date("R8.6.13"), datetime.date(2026, 6, 13))
        self.assertEqual(nippo.name_to_date("R8.4.1"), datetime.date(2026, 4, 1))
        # 棚卸サフィックス（全角/半角どちらも）
        self.assertEqual(nippo.name_to_date("R8.5.31 (棚卸)"), datetime.date(2026, 5, 31))
        self.assertEqual(nippo.name_to_date("R8.4.30（棚卸）"), datetime.date(2026, 4, 30))
        self.assertIsNone(nippo.name_to_date("合計"))

    def test_date_to_name(self):
        self.assertEqual(nippo.date_to_name(datetime.date(2026, 6, 13)), "R8.6.13")
        self.assertEqual(nippo.date_to_name(datetime.date(2026, 7, 1)), "R8.7.1")

    def test_serial_roundtrip(self):
        # A1=46186 は 2026-06-13（実ファイルで確認済み）
        self.assertEqual(nippo.to_serial(datetime.date(2026, 6, 13)), 46186)
        self.assertEqual(nippo.from_serial(46186), datetime.date(2026, 6, 13))

    def test_inventory(self):
        self.assertTrue(nippo.is_inventory("R8.5.31 (棚卸)"))
        self.assertTrue(nippo.is_inventory("R8.4.30（棚卸）"))
        self.assertFalse(nippo.is_inventory("R8.6.13"))


class TestPenSpec(unittest.TestCase):
    def test_parse(self):
        self.assertEqual(
            nippo.parse_pen_spec("中1=2,肉4=3"),
            {("中", 1): 2, ("肉", 4): 3},
        )

    def test_whitespace_and_accumulate(self):
        self.assertEqual(
            nippo.parse_pen_spec(" 中1 = 2 , 中1=1 "),
            {("中", 1): 3},
        )

    def test_empty(self):
        self.assertEqual(nippo.parse_pen_spec(""), {})

    def test_errors(self):
        with self.assertRaises(ValueError):
            nippo.parse_pen_spec("中9=1")   # 房は1..8
        with self.assertRaises(ValueError):
            nippo.parse_pen_spec("豚1=1")   # 房種別が不正
        with self.assertRaises(ValueError):
            nippo.parse_pen_spec("中1")     # = なし
        with self.assertRaises(ValueError):
            nippo.parse_pen_spec("中1=-2")  # 負数

    def test_pen_cell(self):
        self.assertEqual(nippo.pen_cell("中", 1, "death"), "E105")
        self.assertEqual(nippo.pen_cell("中", 8, "head"), "C119")
        self.assertEqual(nippo.pen_cell("肉", 4, "ship"), "T111")
        self.assertEqual(nippo.pen_cell("肉", 1, "move"), "R105")


class TestFormulaEval(unittest.TestCase):
    def _sheet(self, mapping):
        cells = {}
        for ref, val in mapping.items():
            if isinstance(val, str) and val.startswith("="):
                cells[ref] = (val[1:], None)
            else:
                cells[ref] = (None, str(val))
        return nippo.FormulaSheet(cells)

    def test_pure_arithmetic(self):
        fs = self._sheet({"A1": "=0+1+2+5+7+4+6"})
        self.assertEqual(fs.eval_ref("A1"), 25)

    def test_reference_chain(self):
        fs = self._sheet({
            "E105": "=0+1+1",
            "G105": "=0",
            "I105": "0",
            "C105": "=430-E105-G105-I105",
        })
        self.assertEqual(fs.eval_ref("C105"), 428)

    def test_sum_range(self):
        fs = self._sheet({
            "N13": "=5", "N14": "=3", "AC13": "0", "AC14": "0",
            "C9": "=SUM(N13:AC14)",
        })
        self.assertEqual(fs.eval_ref("C9"), 8)

    def test_rate(self):
        fs = self._sheet({"G23": "=0+39", "G27": "=832", "I23": "=G23/G27*100"})
        self.assertAlmostEqual(fs.eval_ref("I23"), 39 / 832 * 100, places=6)

    def test_cycle_guard(self):
        fs = self._sheet({"A1": "=B1", "B1": "=A1"})
        self.assertEqual(fs.eval_ref("A1"), 0)  # 無限ループせず0


class TestColumns(unittest.TestCase):
    def test_col_num_roundtrip(self):
        for col in ("A", "Z", "AA", "AC", "BJ"):
            self.assertEqual(nippo.num_to_col(nippo.col_to_num(col)), col)


if __name__ == "__main__":
    unittest.main(verbosity=2)
