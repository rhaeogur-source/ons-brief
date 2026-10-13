"""출동 전 원인 브리핑 카드 생성기.

알람 로그 + 기지국 마스터 + 장비이력 + 조치이력을 기지국ID로 연결해
dist/briefing.json, dist/index.html(데이터 내장), dist/cards/*.md 를 만든다.

    python prepare.py [--alarm PATH] [--site PATH] [--equip PATH] [--action PATH]
                      [--out dist] [--now 2026-10-08T21:00] [--no-mask]
"""
from __future__ import annotations

import argparse
import csv
import json
import re
from collections import Counter, defaultdict
from datetime import datetime, timedelta
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
CONFIG = HERE / "config"
GRADE_ORDER = {"Critical": 0, "Major": 1, "Minor": 2, "Warning": 3}

DATE_FORMATS = [
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%dT%H:%M:%S",
    "%Y/%m/%d %H:%M",
    "%Y/%m/%d %H:%M:%S",
    "%Y%m%d %H%M%S",
    "%m/%d/%Y %I:%M:%S %p",
    "%Y-%m-%d",
]


# ---------------------------------------------------------------- 읽기

def parse_dt(value) -> datetime | None:
    """혼재된 날짜 형식(6종)을 datetime으로. 빈 값·해석 불가는 None."""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value
    text = str(value).strip()
    if not text:
        return None
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def read_xlsx(path: Path) -> list[dict]:
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True).active
    rows = list(ws.values)
    header = [str(h).strip() for h in rows[0]]
    out = []
    for i, r in enumerate(rows[1:], start=2):  # 엑셀 행 번호(헤더=1)
        rec = dict(zip(header, r))
        rec["_row"] = i
        out.append(rec)
    return out


def read_alarms(path: Path) -> list[dict]:
    with open(path, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))
    for i, r in enumerate(rows, start=2):
        r["_row"] = i
    return rows


def load_config() -> dict:
    with open(CONFIG / "alarm_map.csv", encoding="utf-8") as f:
        alarm_map = {r["알람명"]: r["표준유형"] for r in csv.DictReader(f)}
    return {
        "score": json.loads((CONFIG / "config.json").read_text(encoding="utf-8")),
        "causes": json.loads((CONFIG / "causes.json").read_text(encoding="utf-8")),
        "checklist": json.loads((CONFIG / "checklist.json").read_text(encoding="utf-8")),
        "alarm_map": alarm_map,
    }


# ---------------------------------------------------------------- 판정·계산

def is_open(alarm: dict) -> bool:
    return alarm["상태"].strip() == "미해제" or not alarm["해제시각"].strip()


def needs_check(alarm: dict) -> bool:
    return alarm["상태"].strip() == "해제" and not alarm["해제시각"].strip()


def mask_phone(phone: str | None) -> str:
    if not phone:
        return ""
    return re.sub(r"^(\d{2,3})-(\d{3,4})-(\d{4})$", r"\1-****-\3", str(phone))


def mask_name(name: str | None) -> str:
    if not name:
        return ""
    name = str(name)
    if len(name) <= 2:
        return name[0] + "*"
    return name[0] + "*" * (len(name) - 2) + name[-1]


def level_of(score: int, cfg: dict) -> str:
    for threshold, name in cfg["levels"]:
        if score >= threshold:
            return name
    return cfg["levels"][-1][1]


def score_site(open_alarms: list[dict], day_count: int, tech: str, now: datetime, cfg: dict) -> dict:
    top = min(open_alarms, key=lambda a: GRADE_ORDER.get(a["등급"], 9))
    grade = cfg["gradeScore"].get(top["등급"], 0)
    oldest = min(a["_raised"] for a in open_alarms if a["_raised"])
    hours = max(0.0, (now - oldest).total_seconds() / 3600)
    elapsed = min(cfg["elapsedMax"], int(hours * cfg["elapsedPerHour"]))
    repeat = 0
    if day_count >= cfg["repeatFrom"]:
        repeat = min(cfg["repeatMax"], (day_count - (cfg["repeatFrom"] - 1)) * cfg["repeatPer"])
    multi = cfg["multiOpen"] if len(open_alarms) >= 2 else 0
    tech5g = cfg["tech5g"] if tech == "5G" else 0
    total = min(100, grade + elapsed + repeat + multi + tech5g)
    return {
        "total": total,
        "parts": {"grade": grade, "elapsed": elapsed, "repeat": repeat, "multi": multi, "tech": tech5g},
        "topGrade": top["등급"],
        "hours": round(hours, 1),
    }


def infer_cause(site_id: str, symptom: str, actions: list[dict], now: datetime, cfg: dict) -> dict:
    """유력 원인: 같은 기지국·같은 증상 이력 → 없으면 전 기지국 같은 증상(추정)."""
    excluded = set(cfg["score"]["excludedCauses"])
    recent_from = now - timedelta(days=cfg["score"]["recentDays"])
    same_site = [a for a in actions if a["기지국ID"] == site_id and a["증상"] == symptom]
    natural = sum(1 for a in same_site if a["실제원인"] in excluded)
    result = {
        "symptom": symptom,
        "siteHistory": len(same_site),
        "siteRecent": sum(1 for a in same_site if a["_date"] and a["_date"] >= recent_from),
        "naturalRecovery": natural,
        "recurred": sum(1 for a in same_site if a["재발여부"] == "Y"),
        "basis": "none",
        "cause": None,
        "action": None,
        "confidence": None,
        "evidence": [],
    }
    pool = [a for a in same_site if a["실제원인"] not in excluded]
    basis = "site"
    if not pool:
        pool = [a for a in actions if a["증상"] == symptom and a["실제원인"] not in excluded]
        basis = "global"
    if not pool:
        return result
    counts = Counter(a["실제원인"] for a in pool)
    cause, n = sorted(counts.items(), key=lambda kv: (-kv[1], kv[0]))[0]
    hits = sorted((a for a in pool if a["실제원인"] == cause), key=lambda a: a["조치일"], reverse=True)
    action = Counter(a["조치내용"] for a in hits).most_common(1)[0][0]
    result.update(
        basis=basis,
        cause=cause,
        action=action,
        confidence=f"{n}/{len(pool)}",
        evidence=[
            {"file": "조치이력.xlsx", "row": a["_row"], "date": a["조치일"], "site": a["기지국ID"],
             "cause": a["실제원인"], "action": a["조치내용"], "recurred": a["재발여부"] == "Y"}
            for a in hits[:5]
        ],
    )
    return result


def dispatch_class(top_grade: str, causes: list[dict], cfg: dict) -> str:
    kinds = [cfg["causes"].get(c["cause"], {}).get("kind") for c in causes if c["cause"]]
    if kinds and all(k == "remote" for k in kinds):
        return "원격 조치 먼저"
    if top_grade in ("Critical", "Major"):
        return "지금 바로 출동"
    return "다음 정기점검"


def aging(eq: dict, now: datetime, years: int) -> bool:
    base = parse_dt(eq.get("최근교체일")) or parse_dt(eq.get("설치일"))
    return bool(base and (now - base).days >= years * 365)


def age_bucket(eq: dict) -> str:
    d = parse_dt(eq.get("설치일"))
    if not d:
        return "미상"
    return "~2016" if d.year <= 2016 else "2017~2020" if d.year <= 2020 else "2021~"


def peer_stats(equips: list[dict], actions: list[dict], cfg: dict) -> dict:
    """같은 구분·모델·펌웨어 장비끼리 누적고장 비교 (도전 과제: 공통 원인 확인)."""
    min_units, ratio = cfg["peerMinUnits"], cfg["peerRatio"]
    by_kind: dict[str, list[int]] = defaultdict(list)
    groups: dict[tuple, list[dict]] = defaultdict(list)
    ages: dict[tuple, list[int]] = defaultdict(list)
    for e in equips:
        f = e["누적고장횟수"] or 0
        by_kind[e["장비구분"]].append(f)
        groups[(e["장비구분"], e["모델"], e["펌웨어"])].append(e)
        ages[(e["장비구분"], age_bucket(e))].append(f)
    kind_avg = {k: sum(v) / len(v) for k, v in by_kind.items()}
    stats, suspects = {}, []
    for key, rows in groups.items():
        avg = sum((r["누적고장횟수"] or 0) for r in rows) / len(rows)
        suspect = len(rows) >= min_units and avg >= ratio * kind_avg[key[0]]
        stats[key] = {"units": len(rows), "avg": round(avg, 2), "kindAvg": round(kind_avg[key[0]], 2),
                      "suspect": suspect}
        if suspect:
            site_ids = sorted({r["기지국ID"] for r in rows})
            causes = Counter(a["실제원인"] for a in actions
                             if a["기지국ID"] in site_ids and a["실제원인"] not in cfg["excludedCauses"])
            total = sum(causes.values())
            top = causes.most_common(1)
            common = top[0][0] if top and top[0][1] / total >= 0.5 else None
            notes = sorted({r["비고"] for r in rows if r["비고"]})
            suspects.append({"kind": key[0], "model": key[1], "firmware": key[2], **stats[key],
                             "sites": site_ids, "topCauses": causes.most_common(3), "commonCause": common,
                             "notes": notes,
                             "rows": [r["_row"] for r in rows]})
    suspects.sort(key=lambda s: -s["avg"] / max(s["kindAvg"], 0.01))
    age_list = sorted(({"kind": k, "bucket": b, "units": len(v), "avg": round(sum(v) / len(v), 2)}
                       for (k, b), v in ages.items()),
                      key=lambda r: (r["kind"], ["~2016", "2017~2020", "2021~", "미상"].index(r["bucket"])))
    return {"stats": stats, "suspects": suspects, "ages": age_list}


# ---------------------------------------------------------------- 조립

def build(alarm_path: Path, site_path: Path, equip_path: Path, action_path: Path,
          now: datetime | None = None, mask: bool = True, mask_names: bool = False) -> dict:
    cfg = load_config()
    sc = cfg["score"]
    alarms = read_alarms(alarm_path)
    sites = {s["기지국ID"]: s for s in read_xlsx(site_path)}
    equips = read_xlsx(equip_path)
    actions = read_xlsx(action_path)

    warnings: list[str] = []
    for a in alarms:
        a["_raised"] = parse_dt(a["발생시각"])
        a["_cleared"] = parse_dt(a["해제시각"])
        if a["_raised"] is None:
            warnings.append(f"발생시각 해석 불가: {a['알람ID']} '{a['발생시각']}'")
        if a["해제시각"].strip() and a["_cleared"] is None:
            warnings.append(f"해제시각 해석 불가: {a['알람ID']} '{a['해제시각']}'")
        a["_type"] = cfg["alarm_map"].get(a["알람명"])
        if a["_type"] is None:
            a["_type"] = f"{a['장비구분']}-기타"
            warnings.append(f"매핑 없는 알람명: {a['알람명']}")
        if a["기지국ID"] not in sites:
            warnings.append(f"마스터에 없는 기지국: {a['기지국ID']} ({a['알람ID']})")
    for a in actions:
        a["조치일"] = str(a["조치일"])[:10]
        a["_date"] = parse_dt(a["조치일"])

    if now is None:
        now = max(t for a in alarms for t in (a["_raised"], a["_cleared"]) if t)

    equip_by_site: dict[str, list[dict]] = defaultdict(list)
    for e in equips:
        equip_by_site[e["기지국ID"]].append(e)
    action_count = Counter(a["기지국ID"] for a in actions)
    peers = peer_stats(equips, actions, sc)
    by_site: dict[str, list[dict]] = defaultdict(list)
    for a in alarms:
        by_site[a["기지국ID"]].append(a)

    out_sites = []
    for site_id, site_alarms in by_site.items():
        s = sites.get(site_id, {})
        open_alarms = [a for a in site_alarms if is_open(a)]
        target = open_alarms or site_alarms
        sym_done: dict[str, dict] = {}
        for a in target:
            if a["알람명"] not in sym_done:
                sym_done[a["알람명"]] = infer_cause(site_id, a["알람명"], actions, now, cfg)
        causes = list(sym_done.values())
        equip_kinds = {a["장비구분"] for a in target}
        equip_rows = sorted(
            (e for e in equip_by_site.get(site_id, []) if e["장비구분"] in equip_kinds),
            key=lambda e: (-(e["누적고장횟수"] or 0), e["장비시리얼"]),
        )
        score = score_site(open_alarms, len(site_alarms), s.get("기술", ""), now, sc) if open_alarms else None
        high = s.get("고소작업여부") == "Y"
        kit: list[str] = []
        for c in causes:
            kit += cfg["causes"].get(c["cause"], {}).get("kit", []) if c["cause"] else []
        if not kit:
            for a in target:
                kit += cfg["checklist"].get(a["_type"], [])
        if high:
            kit += sc["highWorkKit"]
        notices = []
        if high:
            notices.append(sc["highWorkNotice"])
        if s.get("설치형태") in ("철탑", "옥탑", "폴"):
            notices.append(sc["towerNotice"])
        tags = []
        if high:
            tags.append("고소작업")
        if any(needs_check(a) for a in open_alarms):
            tags.append("확인필요")
        if site_id not in sites:
            tags.append("마스터없음")
        if not equip_by_site.get(site_id):
            tags.append("장비이력없음")
        if any(c["basis"] == "global" for c in causes):
            tags.append("추정포함")
        if any(peers["stats"][(e["장비구분"], e["모델"], e["펌웨어"])]["suspect"] for e in equip_rows):
            tags.append("펌웨어의심")
        phone = s.get("담당자연락처") or ""
        out_sites.append({
            "id": site_id,
            "name": s.get("기지국명") or site_alarms[0]["기지국명"],
            "branch": s.get("지사") or site_alarms[0]["지사"],
            "region": s.get("권역"),
            "address": s.get("주소"),
            "lat": s.get("위도"),
            "lng": s.get("경도"),
            "vendor": s.get("제조사"),
            "model": s.get("장비모델"),
            "tech": s.get("기술"),
            "installYear": s.get("설치연도"),
            "installType": s.get("설치형태"),
            "highWork": high,
            "contact": {"name": mask_name(s.get("담당자")) if mask_names else s.get("담당자"),
                        "phone": mask_phone(phone) if mask else phone},
            "isOpen": bool(open_alarms),
            "score": score["total"] if score else 0,
            "scoreParts": score["parts"] if score else None,
            "level": level_of(score["total"], sc) if score else "해제",
            "topGrade": (score or {}).get("topGrade") or min(site_alarms, key=lambda a: GRADE_ORDER.get(a["등급"], 9))["등급"],
            "openHours": score["hours"] if score else 0,
            "dispatch": dispatch_class(score["topGrade"], causes, cfg) if score else "다음 정기점검",
            "dayCount": len(site_alarms),
            "historyCount": action_count.get(site_id, 0),
            "openAlarms": [_alarm_json(a) for a in open_alarms],
            "history": [_alarm_json(a) for a in sorted(site_alarms, key=lambda a: a["_raised"] or now)],
            "causes": causes,
            "equipment": [
                {"serial": e["장비시리얼"], "kind": e["장비구분"], "vendor": e["제조사"], "model": e["모델"],
                 "firmware": e["펌웨어"], "installed": str(e["설치일"] or "")[:10],
                 "replaced": str(e["최근교체일"] or "")[:10], "failures": e["누적고장횟수"] or 0,
                 "note": e["비고"], "aging": aging(e, now, sc["agingYears"]), "row": e["_row"],
                 "peer": peers["stats"][(e["장비구분"], e["모델"], e["펌웨어"])]}
                for e in equip_rows
            ],
            "kit": list(dict.fromkeys(kit)),
            "notices": notices,
            "tags": tags,
        })

    out_sites.sort(key=lambda x: (not x["isOpen"], -x["score"], GRADE_ORDER.get(x["topGrade"], 9),
                                  -x["openHours"], x["id"]))
    rank = 0
    for x in out_sites:
        if x["isOpen"]:
            rank += 1
            x["rank"] = rank
        else:
            x["rank"] = None

    branches: dict[str, dict] = {}
    for x in out_sites:
        if not x["isOpen"]:
            continue
        b = branches.setdefault(x["branch"], {"name": x["branch"], "open": 0, "byGrade": Counter(), "needsCheck": 0})
        b["open"] += len(x["openAlarms"])
        for a in x["openAlarms"]:
            b["byGrade"][a["grade"]] += 1
            b["needsCheck"] += a["needsCheck"]
    branch_list = sorted(
        ({**b, "byGrade": dict(b["byGrade"])} for b in branches.values()),
        key=lambda b: (-b["open"], b["name"]),
    )
    repeat_rank = sorted(
        ({"id": x["id"], "name": x["name"], "branch": x["branch"], "historyCount": x["historyCount"],
          "dayCount": x["dayCount"]} for x in out_sites),
        key=lambda r: (-r["historyCount"], -r["dayCount"], r["id"]),
    )[:10]

    open_all = [a for a in alarms if is_open(a)]
    return {
        "meta": {
            "generatedAt": datetime.now().replace(microsecond=0).isoformat(),
            "baseTime": now.replace(microsecond=0).isoformat(),
            "source": {"alarm": alarm_path.name, "site": site_path.name, "equip": equip_path.name,
                       "action": action_path.name},
            "counts": {
                "alarms": len(alarms),
                "open": len(open_all),
                "needsCheck": sum(needs_check(a) for a in alarms),
                "openSites": sum(1 for x in out_sites if x["isOpen"]),
                "alarmSites": len(out_sites),
            },
            "masked": mask,
            "namesMasked": mask_names,
            "warnings": warnings,
        },
        "branches": branch_list,
        "repeatRank": repeat_rank,
        "firmwareSuspects": peers["suspects"],
        "ageCompare": peers["ages"],
        "sites": out_sites,
    }


def _alarm_json(a: dict) -> dict:
    return {
        "id": a["알람ID"],
        "raisedAt": a["_raised"].isoformat() if a["_raised"] else a["발생시각"],
        "clearedAt": a["_cleared"].isoformat() if a["_cleared"] else None,
        "grade": a["등급"],
        "equip": a["장비구분"],
        "code": a["알람코드"],
        "name": a["알람명"],
        "type": a["_type"],
        "status": a["상태"],
        "needsCheck": needs_check(a),
        "row": a["_row"],
    }


# ---------------------------------------------------------------- 출력

def card_markdown(x: dict) -> str:
    lines = [f"# {x['id']} {x['name']}", ""]
    if x["isOpen"]:
        lines.append(f"- 순위 {x['rank']} · {x['score']}점 {x['level']} · **{x['dispatch']}**")
    lines += [
        f"- {x['branch']} · {x['address'] or ''}",
        f"- 장비 {x['vendor'] or ''} {x['model'] or ''} ({x['tech'] or ''}, {x['installYear'] or ''}년) · {x['installType'] or ''}",
        f"- 담당 {x['contact']['name'] or ''} {x['contact']['phone']}",
        "",
    ]
    for n in x["notices"]:
        lines.append(f"> ⚠️ {n}")
    lines += ["", "## 오늘 알람", ""]
    for a in x["openAlarms"] or x["history"]:
        lines.append(f"- {a['raisedAt'][11:16]} {a['grade']} {a['name']} ({a['type']})"
                     + (" · 확인필요" if a["needsCheck"] else ""))
    lines += ["", "## 예상 원인", ""]
    for c in x["causes"]:
        head = f"- **{c['symptom']}** — 이 기지국 과거 {c['siteHistory']}회 (최근 90일 {c['siteRecent']}회)"
        lines.append(head)
        if not c["cause"]:
            lines.append("    - 근거 없음")
            continue
        tag = "근거" if c["basis"] == "site" else "추정 (다른 기지국 같은 증상)"
        lines.append(f"    - {tag}: {c['cause']} → {c['action']} ({c['confidence']})")
        for ev in c["evidence"]:
            lines.append(f"    - {ev['file']} {ev['row']}행 · {ev['date']} · {ev['site']}"
                         + (" · 재발" if ev["recurred"] else ""))
    if x["equipment"]:
        lines += ["", "## 관련 장비", ""]
        for e in x["equipment"]:
            lines.append(f"- {e['kind']} {e['serial']} {e['vendor']} {e['model']} FW {e['firmware']} · "
                         f"고장 {e['failures']}회" + (" · 노후" if e["aging"] else "")
                         + (f" · {e['note']}" if e["note"] else "")
                         + (f" · ⚠️ 같은 모델·펌웨어 {e['peer']['units']}대 평균 {e['peer']['avg']}회"
                            f" (구분 평균 {e['peer']['kindAvg']}회)" if e["peer"]["suspect"] else ""))
    lines += ["", "## 챙길 자재", ""] + [f"- [ ] {k}" for k in x["kit"]]
    return "\n".join(lines) + "\n"


def static_list(data: dict) -> str:
    """스크립트 없이도 보이는 출동 목록 (스크립트가 돌면 덮어씀)."""
    from html import escape as e
    items = []
    for x in data["sites"]:
        if not x["isOpen"]:
            continue
        cause = next((c for c in x["causes"] if c["cause"]), None)
        why = (("추정: " if cause["basis"] == "global" else "") + f"{cause['cause']} → {cause['action']}") if cause else "근거 없음"
        items.append(
            f'<div class="card {e(x["level"])}"><div class="row"><span class="name">{x["rank"]}. {e(x["id"])} · {e(x["name"])}</span>'
            f'<span class="score">{x["score"]} {e(x["level"])}</span></div>'
            f'<div class="meta">{e(x["branch"])} · {e(x["topGrade"])} · {e(x["dispatch"])}'
            f'{" · 고소작업" if x["highWork"] else ""}</div><div class="meta">{e(why)}</div></div>')
    return "\n".join(items)


def write_outputs(data: dict, out: Path) -> None:
    out.mkdir(parents=True, exist_ok=True)
    payload = json.dumps(data, ensure_ascii=False, indent=1, default=str)
    (out / "briefing.json").write_text(payload, encoding="utf-8")
    template = (HERE / "templates" / "index.html").read_text(encoding="utf-8")
    embedded = json.dumps(data, ensure_ascii=False, default=str).replace("</", "<\\/")
    html = template.replace("/*__BRIEFING__*/null", embedded).replace("<!--__STATIC__-->", static_list(data))
    (out / "index.html").write_text(html, encoding="utf-8")
    cards = out / "cards"
    cards.mkdir(exist_ok=True)
    for old in cards.glob("*.md"):
        old.unlink()
    for x in data["sites"]:
        if x["isOpen"]:
            (cards / f"{x['rank']:02d}_{x['id']}.md").write_text(card_markdown(x), encoding="utf-8")


def main(argv: list[str] | None = None) -> dict:
    p = argparse.ArgumentParser(description="출동 전 원인 브리핑 카드 생성")
    p.add_argument("--alarm", default=str(HERE.parent / "alarm_log_20261008.csv"))
    p.add_argument("--site", default=str(HERE.parent / "site_master.xlsx"))
    p.add_argument("--equip", default=str(HERE / "장비이력.xlsx"))
    p.add_argument("--action", default=str(HERE / "조치이력.xlsx"))
    p.add_argument("--out", default=str(HERE / "dist"))
    p.add_argument("--now", help="기준시각 (기본: 로그의 마지막 시각)")
    p.add_argument("--no-mask", action="store_true", help="담당자 연락처를 가리지 않음")
    p.add_argument("--mask-names", action="store_true", help="담당자 이름도 가림 (공개 배포용)")
    args = p.parse_args(argv)
    now = parse_dt(args.now.replace("T", " ") + (":00" if args.now.count(":") == 1 else "")) if args.now else None
    data = build(Path(args.alarm), Path(args.site), Path(args.equip), Path(args.action),
                 now=now, mask=not args.no_mask, mask_names=args.mask_names)
    write_outputs(data, Path(args.out))
    c = data["meta"]["counts"]
    print(f"알람 {c['alarms']}건 · 미해제 {c['open']}건(확인필요 {c['needsCheck']}) · "
          f"출동 대상 {c['openSites']}곳 · 경고 {len(data['meta']['warnings'])}건 → {args.out}")
    return data


if __name__ == "__main__":
    main()
