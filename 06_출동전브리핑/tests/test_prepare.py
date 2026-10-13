import json
import sys
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import prepare  # noqa: E402

ALARM = ROOT.parent / "alarm_log_20261008.csv"
SITE = ROOT.parent / "site_master.xlsx"
EQUIP = ROOT / "장비이력.xlsx"
ACTION = ROOT / "조치이력.xlsx"


class ParseTest(unittest.TestCase):
    def test_all_six_date_formats(self):  # B2
        expect = datetime(2026, 10, 8, 14, 51, 53)
        self.assertEqual(prepare.parse_dt("2026-10-08 14:51:53"), expect)
        self.assertEqual(prepare.parse_dt("2026-10-08T14:51:53"), expect)
        self.assertEqual(prepare.parse_dt("20261008 145153"), expect)
        self.assertEqual(prepare.parse_dt("10/08/2026 02:51:53 PM"), expect)
        self.assertEqual(prepare.parse_dt("2026/10/08 14:51"), datetime(2026, 10, 8, 14, 51))
        self.assertIsNone(prepare.parse_dt(""))
        self.assertIsNone(prepare.parse_dt("어제"))

    def test_open_and_needs_check(self):  # B1
        self.assertTrue(prepare.is_open({"상태": "미해제", "해제시각": ""}))
        self.assertTrue(prepare.is_open({"상태": "해제", "해제시각": ""}))
        self.assertFalse(prepare.is_open({"상태": "해제", "해제시각": "2026-10-08 01:00:00"}))
        self.assertTrue(prepare.needs_check({"상태": "해제", "해제시각": ""}))
        self.assertFalse(prepare.needs_check({"상태": "미해제", "해제시각": ""}))

    def test_mask_name(self):
        self.assertEqual(prepare.mask_name("박현우"), "박*우")
        self.assertEqual(prepare.mask_name("남궁민수"), "남**수")
        self.assertEqual(prepare.mask_name("이준"), "이*")
        self.assertEqual(prepare.mask_name(None), "")

    def test_mask_phone(self):  # T7
        self.assertEqual(prepare.mask_phone("010-4811-9279"), "010-****-9279")
        self.assertEqual(prepare.mask_phone(None), "")


class ScoreTest(unittest.TestCase):
    cfg = prepare.load_config()["score"]

    def test_spec_example_is_77(self):  # T4: 5G, Critical 4시간, 당일 5건
        now = datetime(2026, 10, 8, 12, 0)
        alarms = [{"등급": "Critical", "_raised": datetime(2026, 10, 8, 8, 0)}]
        s = prepare.score_site(alarms, 5, "5G", now, self.cfg)
        self.assertEqual(s["total"], 77)
        self.assertEqual(prepare.level_of(s["total"], self.cfg), "긴급")

    def test_caps(self):
        now = datetime(2026, 10, 9, 12, 0)
        alarms = [{"등급": "Critical", "_raised": datetime(2026, 10, 8, 0, 0)},
                  {"등급": "Minor", "_raised": datetime(2026, 10, 8, 1, 0)}]
        s = prepare.score_site(alarms, 20, "5G", now, self.cfg)
        self.assertEqual(s["parts"], {"grade": 50, "elapsed": 20, "repeat": 15, "multi": 5, "tech": 10})
        self.assertEqual(s["total"], 100)


class CauseTest(unittest.TestCase):
    cfg = prepare.load_config()
    now = datetime(2026, 10, 8, 12, 0)

    def act(self, row, site, symptom, cause, date="2026-09-01", recur="N"):
        return {"_row": row, "기지국ID": site, "증상": symptom, "실제원인": cause, "조치내용": "조치",
                "재발여부": recur, "조치일": date, "_date": prepare.parse_dt(date)}

    def test_site_history_wins_over_global(self):  # C1
        acts = [self.act(2, "A", "X", "원인1"), self.act(3, "B", "X", "원인2"), self.act(4, "C", "X", "원인2")]
        c = prepare.infer_cause("A", "X", acts, self.now, self.cfg)
        self.assertEqual((c["basis"], c["cause"], c["confidence"]), ("site", "원인1", "1/1"))
        self.assertEqual(c["evidence"][0]["row"], 2)

    def test_global_fallback_is_marked_estimate(self):
        acts = [self.act(2, "B", "X", "원인2"), self.act(3, "C", "X", "원인2"), self.act(4, "C", "X", "원인3")]
        c = prepare.infer_cause("A", "X", acts, self.now, self.cfg)
        self.assertEqual((c["basis"], c["cause"], c["confidence"]), ("global", "원인2", "2/3"))

    def test_natural_recovery_excluded(self):  # B7
        acts = [self.act(2, "A", "X", "원인 미상(자연복구)"), self.act(3, "B", "X", "원인2")]
        c = prepare.infer_cause("A", "X", acts, self.now, self.cfg)
        self.assertEqual((c["basis"], c["cause"], c["naturalRecovery"]), ("global", "원인2", 1))

    def test_no_history_means_no_guess(self):
        c = prepare.infer_cause("A", "X", [], self.now, self.cfg)
        self.assertEqual((c["basis"], c["cause"], c["evidence"]), ("none", None, []))

    def test_dispatch_class(self):  # C7
        r = lambda cause: [{"cause": cause}]
        self.assertEqual(prepare.dispatch_class("Major", r("일시적 SW 오류"), self.cfg), "원격 조치 먼저")
        self.assertEqual(prepare.dispatch_class("Major", r("보드 불량"), self.cfg), "지금 바로 출동")
        self.assertEqual(prepare.dispatch_class("Minor", r("보드 불량"), self.cfg), "다음 정기점검")
        self.assertEqual(prepare.dispatch_class("Critical", [{"cause": None}], self.cfg), "지금 바로 출동")


@unittest.skipUnless(SITE.exists() and ACTION.exists(), "개인정보 원본(site_master, 조치이력)이 없는 환경")
class RealDataTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.data = prepare.build(ALARM, SITE, EQUIP, ACTION)

    def test_counts(self):  # T1, T2
        c = self.data["meta"]["counts"]
        self.assertEqual((c["alarms"], c["open"], c["needsCheck"], c["openSites"], c["alarmSites"]),
                         (480, 31, 15, 31, 225))
        self.assertEqual(self.data["meta"]["warnings"], [])

    def test_branch_totals(self):  # T3
        got = {b["name"]: b["open"] for b in self.data["branches"]}
        self.assertEqual(got, {"서울지사": 7, "부산지사": 7, "대구지사": 4, "대전지사": 3, "제주지사": 3,
                               "경기지사": 3, "강원지사": 2, "광주지사": 2})

    def test_ranking_sorted(self):  # T5
        open_sites = [s for s in self.data["sites"] if s["isOpen"]]
        self.assertEqual([s["rank"] for s in open_sites], list(range(1, 32)))
        scores = [s["score"] for s in open_sites]
        self.assertEqual(scores, sorted(scores, reverse=True))

    def test_high_work_has_notice_and_kit(self):  # T6
        for s in self.data["sites"]:
            if s["highWork"]:
                self.assertIn("고소작업", s["tags"])
                self.assertTrue(s["notices"][0].startswith("고소작업"))
                self.assertIn("안전대", s["kit"])

    def test_every_cause_has_evidence_rows(self):  # README 4: 근거 없는 추측 금지
        for s in self.data["sites"]:
            for c in s["causes"]:
                if c["cause"]:
                    self.assertIn(c["basis"], ("site", "global"))
                    self.assertTrue(c["evidence"])
                    if c["basis"] == "site":
                        self.assertTrue(all(e["site"] == s["id"] for e in c["evidence"]))

    def test_masked_by_default(self):
        phones = [s["contact"]["phone"] for s in self.data["sites"] if s["contact"]["phone"]]
        self.assertTrue(phones and all("****" in p for p in phones))

    def test_firmware_suspect_found(self):  # A4
        top = self.data["firmwareSuspects"][0]
        self.assertEqual((top["kind"], top["model"], top["firmware"], top["units"]),
                         ("DU", "SMD-vDU-2100", "v23.2.4-p1", 4))
        self.assertEqual(top["commonCause"], "펌웨어 버그(메모리 누수)로 비정상 재기동")
        for s in self.data["firmwareSuspects"]:
            self.assertGreaterEqual(s["units"], 3)
            self.assertGreaterEqual(s["avg"], 2 * s["kindAvg"])
        flagged = {x["id"] for x in self.data["sites"] if "펌웨어의심" in x["tags"]}
        self.assertIn("SEL-0058", flagged)

    def test_age_compare_buckets(self):
        du = [r["bucket"] for r in self.data["ageCompare"] if r["kind"] == "DU"]
        self.assertEqual(du, ["~2016", "2017~2020", "2021~"])

    def test_outputs_written(self):  # T9 인코딩 포함
        with tempfile.TemporaryDirectory() as tmp:
            prepare.write_outputs(self.data, Path(tmp))
            html = (Path(tmp) / "index.html").read_text(encoding="utf-8")
            self.assertNotIn("/*__BRIEFING__*/null", html)
            self.assertIn("기지국", html)
            self.assertNotIn("<!--__STATIC__-->", html)
            self.assertIn("1. BSN-0008", html)  # 스크립트 없이도 보이는 목록
            head = html.split("</head>")[0]
            self.assertNotIn("<script src=", head)  # 외부 스크립트가 첫 화면을 막지 않음
            back = json.loads((Path(tmp) / "briefing.json").read_text(encoding="utf-8"))
            self.assertEqual(back["meta"]["counts"]["open"], 31)
            self.assertEqual(len(list((Path(tmp) / "cards").glob("*.md"))), 31)


if __name__ == "__main__":
    unittest.main()
