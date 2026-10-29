#!/usr/bin/env python3
"""자금일보 대시보드 빌더 — 표준 라이브러리만 사용 (pip 설치 불필요).

실행:  python build_dashboard.py
읽는 것: 4_자금일보 폴더의 은행 거래내역 xlsx (원본은 수정하지 않음)
  - 월 파일  : 국민_*거래내역*.xlsx / 신한_*거래내역*.xlsx   (폴더 루트)
  - 일 파일  : 추가자료*/거래내역_YYYY-MM-DD.xlsx            (시트 2개: 국민·신한)
  - 대사     : 은행잔액대사_*.xlsx  (월말 잔액 교차검증용, 있으면 사용)
만드는 것: 같은 폴더의 자금일보_대시보드.html (데이터가 안에 박혀 있어 더블클릭으로 열림)

매일 업데이트 = 새 일 파일을 추가자료* 폴더에 넣고 이 스크립트를 다시 실행.
"""
import glob
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", ".."))  # 4_자금일보
COMPANY = "(주)다온생활"
ACCT_KB, ACCT_SH = "국민 업무용", "신한 급여"
ACCOUNTS = [ACCT_KB, ACCT_SH]

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}


# ───────────────────────── xlsx 읽기 (값만) ─────────────────────────
def _col(ref):
    n = 0
    for c in re.match(r"[A-Z]+", ref).group():
        n = n * 26 + ord(c) - 64
    return n - 1


def read_xlsx(path):
    """{시트이름: [행(list)...]} — 수식 셀은 캐시된 값."""
    z = zipfile.ZipFile(path)
    ss = []
    if "xl/sharedStrings.xml" in z.namelist():
        for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall("m:si", NS):
            ss.append("".join(t.text or "" for t in si.iter("{%s}t" % NS["m"])))
    rels = {r.get("Id"): r.get("Target")
            for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
    out = {}
    for s in ET.fromstring(z.read("xl/workbook.xml")).find("m:sheets", NS):
        t = rels[s.get("{%s}id" % NS["r"])].lstrip("/")
        t = t if t.startswith("xl/") else "xl/" + t
        rows = []
        for row in ET.fromstring(z.read(t)).iter("{%s}row" % NS["m"]):
            d = {}
            for c in row.findall("m:c", NS):
                v, ty = c.find("m:v", NS), c.get("t")
                if ty == "s" and v is not None:
                    val = ss[int(v.text)]
                elif ty == "inlineStr":
                    val = "".join(x.text or "" for x in c.iter("{%s}t" % NS["m"]))
                elif v is not None:
                    val = v.text
                    if ty not in ("str", "b", "e"):
                        f = float(val)
                        val = int(f) if f == int(f) else f
                else:
                    val = None
                d[_col(c.get("r"))] = val
            if d:
                rows.append([d.get(i) for i in range(max(d) + 1)])
        out[s.get("name")] = rows
    return out


# ───────────────────────── 은행별 정규화 ─────────────────────────
def _num(v):
    return int(v) if isinstance(v, (int, float)) and v else 0


def parse_kb(rows, src):
    """국민: 거래일시 / 적요 / 의뢰인·수취인 / 입금 / 출금 / 거래후잔액 / 구분."""
    hi = next(i for i, r in enumerate(rows) if r and r[0] == "거래일시")
    h = {n: i for i, n in enumerate(rows[hi])}
    out = []
    for r in rows[hi + 1:]:
        if not r or not r[0]:
            continue
        g = lambda k: r[h[k]] if h[k] < len(r) else None
        dt = datetime.strptime(str(g("거래일시")), "%Y-%m-%d %H:%M:%S")
        out.append(dict(acct=ACCT_KB, date=dt.strftime("%Y-%m-%d"), time=dt.strftime("%H:%M:%S"),
                        type=g("구분") or "", desc=g("의뢰인/수취인") or g("적요") or "",
                        inn=_num(g("입금")), out=_num(g("출금")), bal=_num(g("거래후잔액")), src=src))
    return out


def parse_sh(rows, src):
    """신한: 거래일자 / 거래시간 / 적요 / 기재내용 / 찾으신금액(=출금) / 맡기신금액(=입금) / 거래후잔액."""
    hi = next(i for i, r in enumerate(rows) if r and r[0] == "거래일자")
    h = {n: i for i, n in enumerate(rows[hi])}
    out = []
    for r in rows[hi + 1:]:
        if not r or not r[0]:
            continue
        g = lambda k: r[h[k]] if h[k] < len(r) else None
        d = str(g("거래일자")).replace(".", "-")
        out.append(dict(acct=ACCT_SH, date=d, time=str(g("거래시간")),
                        type=g("적요") or "", desc=g("기재내용") or "",
                        inn=_num(g("맡기신금액")), out=_num(g("찾으신금액")), bal=_num(g("거래후잔액")), src=src))
    return out


def load_all():
    txns, files = [], []
    for p in sorted(glob.glob(os.path.join(ROOT, "*거래내역*.xlsx"))):
        name = os.path.basename(p)
        sheets = read_xlsx(p)
        rows = next(iter(sheets.values()))
        if name.startswith("국민"):
            txns += parse_kb(rows, name)
        elif name.startswith("신한"):
            txns += parse_sh(rows, name)
        else:
            continue
        files.append(name)
    for p in sorted(glob.glob(os.path.join(ROOT, "추가자료*", "거래내역_*.xlsx"))):
        name = os.path.relpath(p, ROOT).replace("\\", "/")
        for sname, rows in read_xlsx(p).items():
            if "국민" in sname:
                txns += parse_kb(rows, name)
            elif "신한" in sname:
                txns += parse_sh(rows, name)
        files.append(name)
    # 같은 거래가 월 파일·일 파일에 중복되면 한 번만
    seen, uniq, dup = set(), [], 0
    for t in txns:
        k = (t["acct"], t["date"], t["time"], t["inn"], t["out"], t["bal"])
        if k in seen:
            dup += 1
            continue
        seen.add(k)
        uniq.append(t)
    return uniq, files, dup


# ───────────────────────── 계좌간이체 매칭 ─────────────────────────
def match_transfers(txns):
    """회사명이 찍힌 거래 중, 같은 날·같은 금액·반대 방향·다른 계좌를 짝짓는다.
    (같은 금액이어도 회사명이 없는 거래는 이체로 보지 않는다.)"""
    for t in txns:
        t["xfer"] = False
    own = [t for t in txns if COMPANY in t["desc"]]
    pairs, used = [], set()
    for a in own:
        if a["out"] == 0 or id(a) in used:
            continue
        cands = [b for b in own if id(b) not in used and b["acct"] != a["acct"]
                 and b["date"] == a["date"] and b["inn"] == a["out"] and b["out"] == 0]
        if not cands:
            continue
        sec = lambda t: sum(int(x) * m for x, m in zip(t["time"].split(":"), (3600, 60, 1)))
        b = min(cands, key=lambda b: abs(sec(b) - sec(a)))
        used.update((id(a), id(b)))
        a["xfer"] = b["xfer"] = True
        pairs.append(dict(date=a["date"], amount=a["out"], fromAcct=a["acct"], toAcct=b["acct"],
                          outDesc=a["desc"], inDesc=b["type"] or b["desc"]))
    unmatched = [t for t in own if id(t) not in used]
    return pairs, unmatched


# ───────────────────────── 일별 집계 ─────────────────────────
def build_days(txns, xfers):
    by_acct = {a: [t for t in txns if t["acct"] == a] for a in ACCOUNTS}
    # 계좌별 최초 전일 잔액 = 첫 거래 후잔액 − 입금 + 출금
    carry = {}
    for a, rows in by_acct.items():
        f = rows[0]
        carry[a] = f["bal"] - f["inn"] + f["out"]
    dates = sorted({t["date"] for t in txns})
    days, mtd = [], {a: [0, 0] for a in ACCOUNTS}
    prev_month = None
    for d in dates:
        if d[:7] != prev_month:
            mtd = {a: [0, 0] for a in ACCOUNTS}
            prev_month = d[:7]
        accts = {}
        for a in ACCOUNTS:
            rows = [t for t in by_acct[a] if t["date"] == d]
            i, o = sum(t["inn"] for t in rows), sum(t["out"] for t in rows)
            close = rows[-1]["bal"] if rows else carry[a]
            mtd[a][0] += i
            mtd[a][1] += o
            accts[a] = dict(open=carry[a], inn=i, out=o, close=close, mtdIn=mtd[a][0], mtdOut=mtd[a][1])
            carry[a] = close
        x = sum(p["amount"] for p in xfers if p["date"] == d)
        tin = sum(v["inn"] for v in accts.values())
        tout = sum(v["out"] for v in accts.values())
        days.append(dict(date=d, accts=accts, xfer=x, realIn=tin - x, realOut=tout - x,
                         totalOpen=sum(v["open"] for v in accts.values()),
                         totalClose=sum(v["close"] for v in accts.values())))
    return days


# ───────────────────────── 검산 ─────────────────────────
def run_checks(txns, days, pairs, unmatched):
    checks = []
    # 1) 거래 행마다 직전 잔액 ± 금액 = 거래후잔액
    bad = []
    for a in ACCOUNTS:
        rows = [t for t in txns if t["acct"] == a]
        for p, c in zip(rows, rows[1:]):
            if p["bal"] + c["inn"] - c["out"] != c["bal"]:
                bad.append(f'{a} {c["date"]} {c["time"]}')
    checks.append(dict(label="거래 행별 잔액 연속성 (직전 잔액 + 입금 − 출금 = 거래후잔액)",
                       ok=not bad, detail=f"{len(txns)}행 확인" if not bad else "불일치: " + ", ".join(bad[:5])))
    # 2) 일별 전일잔액 + 입금 − 출금 = 금일잔액
    bad = [f'{d["date"]} {a}' for d in days for a, v in d["accts"].items()
           if v["open"] + v["inn"] - v["out"] != v["close"]]
    checks.append(dict(label="일별 전일 잔액 + 입금 − 출금 = 금일 잔액", ok=not bad,
                       detail=f"{len(days)}일 × 2계좌" if not bad else "불일치: " + ", ".join(bad[:5])))
    # 3) 계좌간이체: 매칭 안 된 회사명 거래 없음
    checks.append(dict(label="계좌간이체 짝 맞춤 (회사명 거래가 전부 입금·출금 한 쌍)",
                       ok=not unmatched,
                       detail=f"{len(pairs)}쌍 매칭" if not unmatched else
                       "짝 없음: " + ", ".join(f'{t["date"]} {t["acct"]} {t["inn"] or t["out"]:,}' for t in unmatched)))
    # 4) 이체는 입금·출금 양쪽에서 같은 금액
    tin = sum(t["inn"] for t in txns if t["xfer"])
    tout = sum(t["out"] for t in txns if t["xfer"])
    checks.append(dict(label="이체 입금 합계 = 이체 출금 합계", ok=tin == tout, detail=f"{tin:,}원"))
    # 5) 월말 잔액 vs 은행잔액대사
    for p in glob.glob(os.path.join(ROOT, "은행잔액대사_*.xlsx")):
        rows = next(iter(read_xlsx(p).values()))
        m = re.search(r"(\d{4})-(\d{2})", os.path.basename(p))
        ym = f"{m.group(1)}-{m.group(2)}"
        last = [d for d in days if d["date"][:7] == ym]
        if not last:
            continue
        want = {r[0]: r[1] for r in rows[1:] if r and r[0] in ACCOUNTS}
        got = {a: last[-1]["accts"][a]["close"] for a in ACCOUNTS}
        ok = all(want.get(a) == got[a] for a in ACCOUNTS)
        checks.append(dict(label=f"{ym} 월말 잔액 = 은행잔액대사 파일", ok=ok,
                           detail=" / ".join(f"{a} {got[a]:,}" for a in ACCOUNTS)))
    return checks


def main():
    txns, files, dup = load_all()
    txns.sort(key=lambda t: (t["acct"], t["date"], t["time"]))  # 은행 파일 순서 유지(안정 정렬)
    pairs, unmatched = match_transfers(txns)
    days = build_days(txns, pairs)
    checks = run_checks(txns, days, pairs, unmatched)
    disp = sorted(txns, key=lambda t: (t["date"], t["time"]))
    data = dict(
        company=COMPANY, accounts=ACCOUNTS, asOf=days[-1]["date"],
        builtAt=datetime.now().strftime("%Y-%m-%d %H:%M"),
        files=files, dupDropped=dup, days=days, transfers=pairs, checks=checks,
        txns=[dict(date=t["date"], time=t["time"], acct=t["acct"], type=t["type"], desc=t["desc"],
                   inn=t["inn"], out=t["out"], xfer=t["xfer"]) for t in disp])
    with open(os.path.join(HERE, "template.html"), encoding="utf-8") as f:
        html = f.read()
    payload = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")
    html = html.replace("/*__DATA__*/null", payload)
    dest = os.path.join(HERE, "자금일보_대시보드.html")
    with open(dest, "w", encoding="utf-8") as f:
        f.write(html)
    last = days[-1]
    print(f"파일 {len(files)}개, 거래 {len(txns)}행(중복 제거 {dup}), 기준일 {data['asOf']}")
    print(f"  합계 잔액 {last['totalClose']:,}  실제 입금 {last['realIn']:,}  실제 출금 {last['realOut']:,}")
    for c in checks:
        print(("  PASS " if c["ok"] else "  FAIL ") + c["label"] + " — " + c["detail"])
    print("→", dest)
    return 0 if all(c["ok"] for c in checks) else 1


if __name__ == "__main__":
    sys.exit(main())
