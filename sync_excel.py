#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
과제별 엑셀(data/*.xlsx) → index.html 동기화 도구.

사용법:
  python sync_excel.py --init   # data/ 엑셀 4개를 현재 데이터로 생성 (최초 1회)
  python sync_excel.py          # data/ 엑셀을 읽어 index.html을 갱신

엑셀에 업무를 쓰고 `python sync_excel.py`를 실행하면, index.html의
  /* GEN:CATS:START */ ... /* GEN:CATS:END */          (총괄)
  /* GEN:SUBPROJECTS:START */ ... /* GEN:SUBPROJECTS:END */ (세부1·2·3)
영역이 다시 만들어집니다. 이후 git push(gtfc156)하면 GitHub Pages가 갱신됩니다.

엑셀 표 열: 구분키 | 구분명 | 업무명 | 시작일 | 종료일 | 비고 | 출처 | 주관
  - 같은 '구분키'끼리 한 카테고리로 묶이며, 막대 색은 키 등장 순서대로 자동(파랑·보라·초록·주황).
  - 출처: '계획서' 또는 '행정' (비우면 배지 없음).
  - 주관: 총괄/운영위/DLR/전담 중 쉼표로 구분 (세부과제는 보통 비움).
과제 제목·아이콘 등 표시 메타데이터는 아래 PROJECTS 설정에서 바꿉니다.
"""
import os, re, sys, json, datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment

BASE = os.path.dirname(os.path.abspath(__file__))
INDEX = os.path.join(BASE, "index.html")
DATA = os.path.join(BASE, "data")

HEADERS = ["구분키", "구분명", "업무명", "시작일", "종료일", "비고", "출처", "주관"]
CLS = ["a", "b", "c", "d"]
SRC_TO_KO = {"doc": "계획서", "std": "행정"}
KO_TO_SRC = {"계획서": "doc", "행정": "std", "doc": "doc", "std": "std", "계획": "doc"}
JUK_TO_KO = {"chong": "총괄", "op": "운영위", "dlr": "DLR", "ketep": "전담"}
KO_TO_JUK = {"총괄": "chong", "운영위": "op", "DLR": "dlr", "dlr": "dlr",
             "전담": "ketep", "chong": "chong", "op": "op", "ketep": "ketep"}

# ── 과제별 표시 메타데이터 (제목/아이콘만; 업무 내용은 엑셀에서) ───────────────
PROJECTS = [
    {"file": "총괄.xlsx", "target": "CATS", "id": "chong", "with_juk": True},
    {"file": "세부1.xlsx", "target": "SUB", "id": "sub1", "prefix": "u1", "nav": "🔬 세부1과제",
     "title": "세부1과제 · 2026년(2차년도) 수행 업무",
     "sub": "(편집 필요) 세부1과제 수행 업무와 일정 — 예시를 실제 내용으로 교체하세요",
     "chips": ["🧩 <b>세부1</b> 과제", "🤝 주관 <b>KIMM</b>", "🗓️ 2차년도 <b>2026</b>"]},
    {"file": "세부2.xlsx", "target": "SUB", "id": "sub2", "prefix": "u2", "nav": "🧪 세부2과제",
     "title": "세부2과제 · 2026년(2차년도) 수행 업무",
     "sub": "(편집 필요) 세부2과제 수행 업무와 일정 — 예시를 실제 내용으로 교체하세요",
     "chips": ["🧩 <b>세부2</b> 과제", "🤝 주관 <b>KIMM</b>", "🗓️ 2차년도 <b>2026</b>"]},
    {"file": "세부3.xlsx", "target": "SUB", "id": "sub3", "prefix": "u3", "nav": "⚙️ 세부3과제",
     "title": "세부3과제 · 2026년(2차년도) 수행 업무",
     "sub": "(편집 필요) 세부3과제 수행 업무와 일정 — 예시를 실제 내용으로 교체하세요",
     "chips": ["🧩 <b>세부3</b> 과제", "🤝 주관 <b>KIMM</b>", "🗓️ 2차년도 <b>2026</b>"]},
]

# ── 최초 시드 데이터 (--init 으로 엑셀 생성 시 사용) ──────────────────────────
SEED = {
    "chong": [
        {"key": "A", "title": "사업 운영·관리", "tasks": [
            {"name": "2차년도 협약 체결·연구비 카드 등록·집행계획 수립", "juk": ["chong"], "start": "2026-01-02", "end": "2026-03-12", "note": "협약 3/12 기준", "src": "std"},
            {"name": "참여연구원 등록·변경 및 인건비 계상 관리", "juk": ["chong"], "start": "2026-01-05", "end": "2026-12-31", "note": "연중 상시", "src": "std"},
            {"name": "운영위원회 개최 (연 1회 이상)", "juk": ["op"], "start": "2026-09-01", "end": "2026-09-30", "note": "KETEP 간사·국내전문가(가스터빈·수전해) 포함", "src": "doc"},
            {"name": "세부과제(1~5) 진척도 정기 점검 회의", "juk": ["chong"], "start": "2026-03-01", "end": "2026-12-31", "note": "분기별 (3·6·9·12월)", "src": "doc"},
            {"name": "KPI 기반 진척도 모니터링·조기경보(Early Warning)", "juk": ["chong"], "start": "2026-01-05", "end": "2026-12-31", "note": "연차·분기 보고 체계", "src": "doc"},
            {"name": "외부 전문가 Peer-review (내부 모니터링 병행)", "juk": ["chong"], "start": "2026-10-01", "end": "2026-11-30", "note": "", "src": "doc"},
            {"name": "부진과제 컨설팅·개선계획 수립 (필요시)", "juk": ["chong"], "start": "2026-11-01", "end": "2026-12-15", "note": "조건부 수행", "src": "doc"},
            {"name": "연구노트·보안등급·연구윤리 관리", "juk": ["chong"], "start": "2026-01-05", "end": "2026-12-31", "note": "보안등급 분류 기준", "src": "doc"},
        ]},
        {"key": "B", "title": "국제협력 (KIMM–DLR)", "tasks": [
            {"name": "KIMM–DLR 협력채널 유지·실무 협의", "juk": ["chong", "dlr"], "start": "2026-01-05", "end": "2026-12-31", "note": "내부협의체 정기 운영", "src": "doc"},
            {"name": "국제 워크숍 — 한국(기계연) 준비·개최", "juk": ["chong"], "start": "2026-06-16", "end": "2026-08-05", "note": "8/5 개최 (워크숍 탭 참조)", "src": "doc"},
            {"name": "국제 워크숍/기술교류 — 독일(DLR) 방문", "juk": ["chong", "dlr"], "start": "2026-10-01", "end": "2026-11-15", "note": "한·독 각 1회 이상", "src": "doc"},
            {"name": "연구인력 교류 — DLR 파견·방문연구·세미나 지원", "juk": ["chong", "dlr"], "start": "2026-04-01", "end": "2026-12-31", "note": "하반기 집중", "src": "doc"},
            {"name": "EKC 한-유럽 과학기술학술회의 특별세션 개설·발표", "juk": ["chong"], "start": "2026-05-01", "end": "2026-07-31", "note": "에너지/기계 분과 세션", "src": "doc"},
        ]},
        {"key": "C", "title": "성과 관리", "tasks": [
            {"name": "공동 논문(공동저자)·특허(공동출원) 성과 관리", "juk": ["chong", "dlr"], "start": "2026-03-01", "end": "2026-12-31", "note": "공동소유·활용 계약 기반", "src": "doc"},
            {"name": "DLR 공동특허 출원·기술이전 체계 운영", "juk": ["chong", "dlr"], "start": "2026-06-01", "end": "2026-12-31", "note": "", "src": "doc"},
            {"name": "수요기업 사업화 설명회 (단계별 1회)", "juk": ["chong"], "start": "2026-09-01", "end": "2026-10-31", "note": "국내 수요기업 대상", "src": "doc"},
            {"name": "성과집·기술 브로슈어 발간", "juk": ["chong"], "start": "2026-11-01", "end": "2026-12-31", "note": "", "src": "doc"},
        ]},
        {"key": "D", "title": "평가·보고 (연차 마감)", "tasks": [
            {"name": "2차년도 연차실적보고서 작성·제출", "juk": ["chong"], "start": "2026-12-01", "end": "2027-02-15", "note": "KETEP 제출", "src": "std"},
            {"name": "2차년도 연구비 집행 점검·정산", "juk": ["chong"], "start": "2027-01-02", "end": "2027-02-28", "note": "", "src": "std"},
            {"name": "자체평가·연차평가(점검) 대응", "juk": ["chong"], "start": "2026-12-15", "end": "2027-01-31", "note": "", "src": "std"},
            {"name": "3차년도 사업계획서·협약 준비", "juk": ["chong"], "start": "2026-11-01", "end": "2027-03-15", "note": "차년도 연속 수행", "src": "std"},
        ]},
    ],
}
for n in (1, 2, 3):
    SEED[f"sub{n}"] = [
        {"key": "1", "title": "(예시) 업무 구분 1", "tasks": [
            {"name": f"(예시) 세부{n} 핵심 연구·실험 항목", "juk": [], "start": "2026-01-05", "end": "2026-06-30", "note": "예시 — 실제 업무로 교체", "src": "doc"},
            {"name": f"(예시) 세부{n} 중간 점검·보고", "juk": [], "start": "2026-07-01", "end": "2026-09-30", "note": "예시", "src": "std"},
        ]},
    ]


# ───────────────────────── helpers ─────────────────────────
def js(s):
    return json.dumps(s, ensure_ascii=False)


def fmt_date(v):
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    m = re.match(r"^(\d{4})[-./](\d{1,2})[-./](\d{1,2})", s)
    if m:
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    return s


def norm_src(v):
    if v is None:
        return None
    s = str(v).strip()
    return KO_TO_SRC.get(s)


def norm_juk(v):
    if not v:
        return []
    out = []
    for tok in str(v).replace("，", ",").split(","):
        tok = tok.strip()
        if tok and tok in KO_TO_JUK and KO_TO_JUK[tok] not in out:
            out.append(KO_TO_JUK[tok])
    return out


# ───────────────────────── read xlsx → cats ─────────────────────────
def read_project(path):
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    cats, by_key = [], {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row is None:
            continue
        vals = (list(row) + [None] * 8)[:8]
        key, name, task, start, end, note, src, juk = vals
        key = "" if key is None else str(key).strip()
        task = "" if task is None else str(task).strip()
        if not task:
            continue
        if key not in by_key:
            cat = {"key": key or str(len(cats) + 1), "cls": CLS[len(cats) % 4],
                   "title": (str(name).strip() if name else key), "tasks": []}
            by_key[key] = cat
            cats.append(cat)
        t = {"name": task, "start": fmt_date(start), "end": fmt_date(end)}
        if note and str(note).strip():
            t["note"] = str(note).strip()
        s = norm_src(src)
        if s:
            t["src"] = s
        j = norm_juk(juk)
        if j:
            t["juk"] = j
        by_key[key]["tasks"].append(t)
    return cats


# ───────────────────────── cats → JS ─────────────────────────
def emit_task(t, with_juk):
    parts = [f"name:{js(t['name'])}"]
    if with_juk and t.get("juk"):
        parts.append("juk:[" + ",".join(js(x) for x in t["juk"]) + "]")
    parts += [f"start:{js(t['start'])}", f"end:{js(t['end'])}"]
    if t.get("note"):
        parts.append(f"note:{js(t['note'])}")
    if t.get("src"):
        parts.append(f"src:{js(t['src'])}")
    return "{ " + ", ".join(parts) + " }"


def emit_cats(cats, with_juk, pad):
    blocks = []
    for cat in cats:
        tl = (",\n").join(pad + "    " + emit_task(t, with_juk) for t in cat["tasks"])
        blocks.append(f"{pad}{{ key:{js(cat['key'])}, cls:{js(cat['cls'])}, "
                      f"title:{js(cat['title'])}, tasks:[\n{tl} ]}}")
    return "[\n" + ",\n".join(blocks) + " ]"


def build_cats_js(cats):
    return "const CATS = " + emit_cats(cats, True, "  ") + ";"


def build_subprojects_js(sub_defs):
    objs = []
    for p, cats in sub_defs:
        chips = "[" + ",".join(js(c) for c in p["chips"]) + "]"
        objs.append(
            f"  {{ id:{js(p['id'])}, prefix:{js(p['prefix'])}, nav:{js(p['nav'])},\n"
            f"    title:{js(p['title'])},\n"
            f"    sub:{js(p['sub'])},\n"
            f"    chips:{chips},\n"
            f"    cats:{emit_cats(cats, False, '    ')} }}")
    return "const SUBPROJECTS = [\n" + ",\n".join(objs) + "\n];"


def inject(html, marker, new_code):
    pat = re.compile(r"(/\* GEN:%s:START \*/\n).*?(\n/\* GEN:%s:END \*/)" % (marker, marker), re.DOTALL)
    if not pat.search(html):
        sys.exit(f"[오류] index.html에서 GEN:{marker} 마커를 찾을 수 없습니다.")
    return pat.sub(lambda m: m.group(1) + new_code + m.group(2), html)


# ───────────────────────── write xlsx (--init) ─────────────────────────
def write_xlsx(path, cats, with_juk):
    wb = Workbook()
    ws = wb.active
    ws.title = "업무"
    navy = PatternFill("solid", fgColor="16335E")
    head_font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    body_font = Font(name="Arial", size=10)
    thin = Side(style="thin", color="D0D7E2")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    wrap = Alignment(vertical="center", wrap_text=True)
    center = Alignment(horizontal="center", vertical="center")
    for c, h in enumerate(HEADERS, 1):
        cell = ws.cell(1, c, h)
        cell.fill = navy
        cell.font = head_font
        cell.alignment = center
        cell.border = border
    comments = {
        1: "같은 키끼리 한 카테고리로 묶임. 막대 색은 키 등장 순서대로 자동(파랑·보라·초록·주황).",
        7: "계획서 또는 행정 (비우면 배지 없음).",
        8: "총괄/운영위/DLR/전담 중 쉼표로 구분. 세부과제는 보통 비움.",
    }
    for col, txt in comments.items():
        ws.cell(1, col).comment = Comment(txt, "sync_excel")
    r = 2
    for cat in cats:
        for t in cat["tasks"]:
            ws.cell(r, 1, cat["key"])
            ws.cell(r, 2, cat["title"])
            ws.cell(r, 3, t["name"])
            for col, kk in ((4, "start"), (5, "end")):
                y, m, d = (int(x) for x in t[kk].split("-"))
                dc = ws.cell(r, col, datetime.date(y, m, d))
                dc.number_format = "yyyy-mm-dd"
            ws.cell(r, 6, t.get("note", ""))
            ws.cell(r, 7, SRC_TO_KO.get(t.get("src"), ""))
            ws.cell(r, 8, ", ".join(JUK_TO_KO[j] for j in t.get("juk", [])) if with_juk else "")
            for col in range(1, 9):
                cc = ws.cell(r, col)
                cc.font = body_font
                cc.border = border
                cc.alignment = wrap if col in (3, 6) else center if col in (1, 4, 5, 7) else Alignment(vertical="center")
            r += 1
    widths = [8, 20, 52, 13, 13, 30, 10, 16]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[chr(64 + i)].width = w
    ws.freeze_panes = "A2"
    dv = DataValidation(type="list", formula1='"계획서,행정"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"G2:G{max(r, 200)}")
    wb.save(path)


def do_init():
    os.makedirs(DATA, exist_ok=True)
    for p in PROJECTS:
        cats = SEED[p["id"]]
        # SEED uses list-of-cat dicts with tasks already; reuse directly
        write_xlsx(os.path.join(DATA, p["file"]), cats, p.get("with_juk", False))
        print(f"  생성: data/{p['file']}  (카테고리 {len(cats)}개)")
    print("초기 엑셀 생성 완료. 이제 내용을 편집한 뒤 `python sync_excel.py`로 반영하세요.")


def do_sync():
    if not os.path.exists(INDEX):
        sys.exit("[오류] index.html 없음")
    html = open(INDEX, encoding="utf-8").read()
    cats_chong, sub_defs, total = None, [], 0
    for p in PROJECTS:
        path = os.path.join(DATA, p["file"])
        if not os.path.exists(path):
            sys.exit(f"[오류] {path} 없음. 먼저 `python sync_excel.py --init` 실행.")
        cats = read_project(path)
        n = sum(len(c["tasks"]) for c in cats)
        total += n
        print(f"  읽음: data/{p['file']}  카테고리 {len(cats)}개 · 업무 {n}건")
        if p["target"] == "CATS":
            cats_chong = cats
        else:
            sub_defs.append((p, cats))
    html = inject(html, "CATS", build_cats_js(cats_chong))
    html = inject(html, "SUBPROJECTS", build_subprojects_js(sub_defs))
    open(INDEX, "w", encoding="utf-8").write(html)
    print(f"index.html 갱신 완료 (총 업무 {total}건). git push(gtfc156) 하면 사이트에 반영됩니다.")


if __name__ == "__main__":
    if "--init" in sys.argv:
        do_init()
    else:
        do_sync()
