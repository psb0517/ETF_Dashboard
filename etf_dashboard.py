import io, re, requests, time, json, html
from datetime import date, datetime, timedelta
from pathlib import Path
import pandas as pd

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}

def fetch_bytes(d: date, provider: str, etf_id):
    """URL에서 바이트를 받아 (content, ext) 반환. 없으면 (None, ext)."""
    if provider == "timefolio":
        url = f"https://timeetf.co.kr/pdf_excel.php?idx={etf_id}&cate=&pdfDate={d.isoformat()}&"
        ext = ".xlsx"
    elif provider == "koact":
        ymd = d.strftime("%Y%m%d")
        url = f"https://www.samsungactive.co.kr/excel_pdf.do?fId=2ETF{etf_id}&gijunYMD={ymd}"
        ext = ".xls"
    else:
        return None, ""
    try:
        resp = requests.get(url, headers=HEADERS, timeout=30)
    except requests.RequestException:
        return None, ext
    if resp.status_code != 200:
        return None, ext
    content = resp.content
    looks_like_html = (content[:15].lstrip().lower().startswith(b"<!doctype") or
                       content[:6].lstrip().lower().startswith(b"<html"))
    if len(content) < 200 or looks_like_html:
        return None, ext
    return content, ext


# ② 설정 — ETF 목록과 기간만 여기서 수정하세요
from pathlib import Path
from datetime import datetime as _dt

_SCRIPT_DIR = Path(__file__).parent

ETF_LIST = [
    # (표시이름,            운용사,        고유번호)
    ("Time_글로벌AI인공지능액티브",    "timefolio", 6),
    ("Time_나스닥100액티브",           "timefolio", 2),
    ("Time_글로벌우주방산액티브",       "timefolio", 20),
    ("Time_글로벌휴머노이드액티브",     "timefolio", 25),
    ("Time_코스닥액티브",              "timefolio", 24),
    ("Time_K바이오액티브",             "timefolio", 13),
    ("Time_K이노베이션액티브",         "timefolio", 17),
    ("KoAct_나스닥성장기업액티브",      "koact", "Q1"),
    ("KoAct_글로벌AI로봇액티브",       "koact", "L3"),
    ("KoAct_글로벌AI메모리반도체액티브","koact", "U5"),
    ("KoAct_코스닥액티브",             "koact", "U6"),
    ("KoAct_바이오헬스케어액티브",      "koact", "J9"),
    ("KoAct_미국로봇피지컬AI액티브",    "koact", "U7"),
    ("KoAct_수소전력ESS인프라액티브",   "koact", "T9"),
]

OUTPUT_HTML = _SCRIPT_DIR / "docs" / "index.html"
PLOTLY_CDN = "https://cdn.plot.ly/plotly-2.35.2.min.js"

OUTPUT_HTML.parent.mkdir(parents=True, exist_ok=True)
print(f"ETF: {len(ETF_LIST)}개  |  출력: {OUTPUT_HTML}")


# ③ 엔진 — 수정할 필요 없습니다. 실행만 하세요.
import re, json, io, time, html
from datetime import date as _date, datetime
import pandas as pd

HEADER_KEYS = {
    "code":   ["단축코드", "티커", "ticker", "code", "코드", "종목코드", "isin"],
    "name":   ["종목명", "name", "이름", "securityname", "secname"],
    "qty":    ["수량", "보유수량", "주수", "qty", "shares", "quantity"],
    "value":  ["평가금액", "평가", "금액", "평가액", "value", "marketvalue"],
    "weight": ["비중", "weight", "weights", "ratio", "%"],
}
DATE_RE = re.compile(r"(20\d{2})[-_.]?(\d{2})[-_.]?(\d{2})")
SKIP_NAME = ("합계", "총계", "합 계", "소계", "total", "sum")


def find_col(columns, role):
    """키 우선 탐색: 리스트 앞 키가 뒤 키보다 항상 우선 선택됨.
    (KoAct 등 '단축코드'와 '종목코드(ISIN)'가 공존할 때 단축코드를 우선 사용)"""
    keys = HEADER_KEYS[role]
    for k in keys:
        for col in columns:
            norm = str(col).strip().lower().replace(" ", "")
            if k in norm:
                return col
    return None


def to_number(series):
    """'1,234' 같은 문자열도 숫자로 변환."""
    return pd.to_numeric(
        series.astype(str).str.replace(",", "", regex=False).str.strip().replace({"": None, "nan": None, "None": None}),
        errors="coerce",
    )


def _read_excel_any(content, ext, header=0):
    """bytes + 확장자 → DataFrame. 포맷 자동 탐지, 실패 시 None."""
    if ext in (".xlsx", ".xlsm"):
        engines = ["openpyxl"]
    elif ext == ".xls":
        engines = ["xlrd", "openpyxl"]
    else:
        engines = [None]
    for eng in engines:
        try:
            kw = {"dtype": object, "header": header}
            if eng:
                kw["engine"] = eng
            return pd.read_excel(io.BytesIO(content), **kw)
        except Exception:
            continue
    # HTML 표 폴백 (lxml 필요)
    for kw in ({"header": header},
               {"header": header, "encoding": "utf-8"},
               {"header": header, "encoding": "cp949"}):
        try:
            tables = pd.read_html(io.BytesIO(content), **kw)
            if tables:
                return max(tables, key=lambda t: t.shape[0] * t.shape[1])
        except Exception:
            continue
    return None


def read_from_bytes(content, ext):
    """bytes → {key: {code,name,qty,value,weight}} 딕셔너리."""
    raw = _read_excel_any(content, ext, header=0)
    if raw is None:
        return None

    # 헤더가 첫 줄이 아닐 가능성 대비: 비중/수량 컬럼을 못 찾으면 상위 5행을 헤더 후보로 재시도
    cols = list(raw.columns)
    if find_col(cols, "name") is None or (find_col(cols, "qty") is None and find_col(cols, "value") is None):
        for hdr in range(1, 6):
            raw2 = _read_excel_any(content, ext, header=hdr)
            if raw2 is None:
                continue
            if find_col(raw2.columns, "name") and (find_col(raw2.columns, "qty") or find_col(raw2.columns, "value")):
                raw, cols = raw2, list(raw2.columns)
                break

    c_code = find_col(cols, "code")
    c_name = find_col(cols, "name")
    c_qty = find_col(cols, "qty")
    c_val = find_col(cols, "value")
    c_wt = find_col(cols, "weight")
    # 안전장치: 종목명이 코드열과 겹치거나 비면, 코드/수량/금액/비중이 아닌 다른 열에서 이름 찾기
    if c_name is not None and c_name == c_code:
        c_name = None
    if c_name is None:
        for col in cols:
            if col in (c_code, c_qty, c_val, c_wt):
                continue
            n = str(col).strip().lower().replace(" ", "")
            if "명" in n or "name" in n or "종목" in n:
                c_name = col
                break
    if c_name is None or (c_qty is None and c_val is None):
        print(f"  [건너뜀] 필수 컬럼 인식 실패 (컬럼: {cols})")
        return None

    df = pd.DataFrame()
    df["code"] = raw[c_code].astype(str).str.strip() if c_code else ""
    df["name"] = raw[c_name].astype(str).str.strip()
    df["qty"] = to_number(raw[c_qty]) if c_qty else 0
    df["value"] = to_number(raw[c_val]) if c_val else 0
    df["weight"] = to_number(raw[c_wt]) if c_wt else 0
    df = df.fillna({"qty": 0, "value": 0, "weight": 0})
    df["code"] = df["code"].replace({"nan": "", "None": "", "NaN": ""})

    holdings = {}
    for _, r in df.iterrows():
        name = r["name"]
        if not name or str(name).lower() in ("nan", "none"):
            continue
        if any(s in str(name).lower() for s in SKIP_NAME):
            continue
        code = str(r["code"]).strip()
        if code.lower() in ("", "nan", "none", "nat", "<na>"):
            code = ""
        key = code if code else name          # 종목코드 없으면(현금 등) 종목명을 키로
        holdings[key] = {
            "code": code,
            "name": name,
            "qty": float(r["qty"]),
            "value": float(r["value"]),
            "weight": float(r["weight"]),
        }
    # KoAct 등 비중이 소수(0.069)로 저장된 경우 100을 곱해 %로 통일
    weights = [v["weight"] for v in holdings.values() if v["weight"] > 0]
    if weights and max(weights) < 2.0:
        for v in holdings.values():
            v["weight"] *= 100
    return holdings


def build_all_etf_data():
    """ETF_LIST의 각 ETF에서 최신일, 1주 전, 4주 전 3개 날짜만 다운로드·파싱해 반환."""
    today = _date.today()

    all_data = {}
    for etf_name, provider, etf_id in ETF_LIST:
        print(f"[{etf_name}] 다운로드 중... ", end="", flush=True)
        holdings_by_date = {}

        def try_fetch(target: _date):
            """target 날짜부터 역순으로 최대 14일 탐색, 데이터 있는 첫 날짜 반환."""
            for offset in range(14):
                d = target - timedelta(days=offset)
                if d.weekday() >= 5:
                    continue
                if d.isoformat() in holdings_by_date:
                    return d
                content, ext = fetch_bytes(d, provider, etf_id)
                time.sleep(0.3)
                if content is None:
                    continue
                h = read_from_bytes(content, ext)
                if h:
                    holdings_by_date[d.isoformat()] = h
                    return d
            return None

        latest_d = try_fetch(today)
        if latest_d is None:
            print("데이터 없음")
            continue

        try_fetch(latest_d - timedelta(days=1))   # 전 거래일
        try_fetch(latest_d - timedelta(days=7))
        try_fetch(latest_d - timedelta(days=28))

        dates = sorted(holdings_by_date.keys())
        print(f"{len(dates)}일 인식 ({', '.join(dates)})")
        all_data[etf_name] = {"dates": dates, "holdings": holdings_by_date}
    return all_data


def ensure_plotly_local(_out_dir):
    return PLOTLY_CDN


def generate():
    print("ETF 대시보드 생성을 시작합니다.\n")
    etf_data = build_all_etf_data()

    if not etf_data:
        print("표시할 데이터가 없습니다. ETF_LIST 설정을 확인하세요.")
        return None

    plotly_src = ensure_plotly_local(OUTPUT_HTML.parent)

    template = TEMPLATE
    html_out = (
        template
        .replace("%%PLOTLY_SRC%%", html.escape(plotly_src, quote=True))
        .replace("%%DATA_JSON%%", json.dumps(etf_data, ensure_ascii=False))
        .replace("%%DEFAULT_ETF%%", json.dumps(list(etf_data.keys())[0], ensure_ascii=False))
        .replace("%%GENERATED%%", datetime.now().strftime("%Y-%m-%d %H:%M"))
    )
    OUTPUT_HTML.write_text(html_out, encoding="utf-8")
    print(f"완료 ✅  →  {OUTPUT_HTML}")
    print("브라우저로 열어 확인하세요.")
    return OUTPUT_HTML


# ===== 내장 HTML 템플릿 =====
TEMPLATE = r"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>ETF 구성종목 일별 변동 대시보드</title>
<link rel="preconnect" href="https://cdn.jsdelivr.net">
<link href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css" rel="stylesheet">
<link href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;700&display=swap" rel="stylesheet">
<style>
  :root{
    /* ── 네이비 다크 테마(이전보다 밝게) ── */
    --paper:#19202e; --panel:#222c3e; --panel2:#2a3650; --ink:#e9ecf3;
    --muted:#9aa3b5; --faint:#6b7689; --line:#374459; --line2:#2c384c;
    --accent:#7aa7ff; --accent-soft:#2b3850; --rule:#3f4d66; --hover:#2a3650;
    /* 한국 시장 관례: 상승=빨강 / 하락=파랑 */
    --up:#ff6058; --down:#6aa6ff;
    --new-line:#3fb950; --new-bg:rgba(63,185,80,.13);
    --sold-line:#f85149; --sold-bg:rgba(248,81,73,.12);
    --exp-text:#86e0a0; --exp-bg:rgba(63,185,80,.16); --exp-bd:rgba(63,185,80,.40);
    --con-text:#ff9d96; --con-bg:rgba(248,81,73,.15); --con-bd:rgba(248,81,73,.38);
    --hold:#9aa3b5;
    --mono:'JetBrains Mono',ui-monospace,'SF Mono',Menlo,Consolas,monospace;
    --sans:'Pretendard',-apple-system,BlinkMacSystemFont,'Apple SD Gothic Neo','Malgun Gothic','Noto Sans KR',sans-serif;
  }
  *{box-sizing:border-box}
  body{margin:0;background:var(--paper);color:var(--ink);font-family:var(--sans);
    font-size:14px;line-height:1.45;-webkit-font-smoothing:antialiased;}
  .wrap{max-width:1280px;margin:0 auto;padding:22px 24px 60px;}

  header{display:flex;align-items:flex-end;justify-content:space-between;gap:18px;
    flex-wrap:wrap;padding-bottom:16px;border-bottom:2px solid var(--rule);}
  .brand{display:flex;flex-direction:column;gap:3px}
  .eyebrow{font-size:11px;letter-spacing:.18em;text-transform:uppercase;color:var(--accent);font-weight:600}
  h1{font-size:23px;font-weight:800;margin:0;letter-spacing:-.01em}
  select{font-family:var(--mono);font-size:13px;font-weight:500;padding:8px 12px;border:1px solid var(--line);
    border-radius:8px;background:var(--panel);color:var(--ink);cursor:pointer;min-width:150px}

  /* ── 티커·종목명 검색 ── */
  .search-box{position:relative;display:flex;align-items:center;flex:0 1 260px}
  .search-box input{font-family:var(--mono);font-size:13px;padding:9px 30px 9px 32px;border:1px solid var(--line);
    border-radius:8px;background:var(--panel);color:var(--ink);width:100%;outline:none}
  .search-box input:focus{border-color:var(--accent)}
  .search-box input::placeholder{color:var(--faint)}
  .search-box .sicon{position:absolute;left:10px;top:50%;transform:translateY(-50%);color:var(--faint);pointer-events:none;font-size:12px}
  .search-box .sclear{position:absolute;right:8px;top:50%;transform:translateY(-50%);cursor:pointer;color:var(--faint);
    font-size:14px;line-height:1;display:none;padding:2px}
  .search-box .sclear:hover{color:var(--ink)}
  .sug{position:absolute;top:calc(100% + 6px);right:0;width:min(360px,86vw);max-height:340px;overflow:auto;z-index:40;
    background:var(--panel2);border:1px solid var(--line);border-radius:10px;box-shadow:0 12px 30px rgba(0,0,0,.45);display:none}
  .sug.on{display:block}
  .sug-item{padding:9px 12px;border-bottom:1px solid var(--line2);cursor:pointer;display:flex;justify-content:space-between;align-items:center;gap:10px}
  .sug-item:last-child{border-bottom:0}
  .sug-item:hover,.sug-item.active{background:var(--hover)}
  .sug-name{font-weight:600;font-size:13px}
  .sug-name .code{color:var(--faint);font-family:var(--mono);font-size:11px;margin-left:7px}
  .sug-cnt{font-family:var(--mono);font-size:11px;color:var(--accent);white-space:nowrap}
  .sug-empty{padding:14px 12px;color:var(--muted);font-size:12.5px}

  /* ── 검색 결과 패널 ── */
  .spanel{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:15px 16px;margin:18px 0 6px;display:none}
  .spanel.on{display:block}
  .sp-head{display:flex;justify-content:space-between;align-items:flex-start;gap:12px;margin-bottom:12px}
  .sp-title{font-size:16px;font-weight:800}
  .sp-title .code{font-family:var(--mono);font-size:12px;color:var(--faint);margin-left:8px}
  .sp-sub{font-size:12px;color:var(--muted);margin-top:3px}
  .sp-sub b{color:var(--accent);font-weight:700}
  .sp-close{appearance:none;border:1px solid var(--line);background:var(--panel2);color:var(--muted);
    border-radius:8px;padding:6px 11px;font-size:12px;font-weight:600;cursor:pointer;white-space:nowrap}
  .sp-close:hover{color:var(--ink);border-color:var(--accent)}
  .holder-row{display:grid;grid-template-columns:1.7fr .8fr .9fr .9fr .9fr auto;gap:10px;align-items:center;
    padding:9px 10px;border-bottom:1px solid var(--line2);cursor:pointer}
  .holder-row:hover{background:var(--hover)}
  .holder-row.head{cursor:default;font-size:11px;color:var(--muted);font-weight:700;letter-spacing:.02em;
    border-bottom:1px solid var(--line);background:transparent}
  .holder-row.head:hover{background:transparent}
  .h-etf{font-weight:600;font-size:13px;display:flex;align-items:center;gap:7px;min-width:0}
  .h-etf .lbl{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .h-num{font-family:var(--mono);text-align:right;font-variant-numeric:tabular-nums;font-size:12.5px}
  .h-num .code{color:var(--faint);font-size:11px}
  .h-jump{font-size:11px;color:var(--accent);white-space:nowrap;text-align:right}
  .gbadge{font-size:10px;font-weight:700;padding:1px 6px;border-radius:5px;letter-spacing:.02em;flex:none}
  .gbadge.time{background:var(--accent-soft);color:var(--accent)}
  .gbadge.koact{background:var(--exp-bg);color:var(--exp-text)}
  @media(max-width:880px){
    .holder-row{grid-template-columns:1.5fr .8fr .9fr auto;font-size:12px}
    .holder-row .h-rank,.holder-row .h-jump{display:none}
  }

  .cards{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin:20px 0 8px}
  .card{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:15px 16px;position:relative;overflow:hidden}
  .card::before{content:"";position:absolute;left:0;top:0;bottom:0;width:4px}
  .card.cnew::before{background:var(--new-line)} .card.csold::before{background:var(--sold-line)}
  .card.cup::before{background:var(--up)} .card.cdown::before{background:var(--down)}
  .card .k{font-size:11.5px;color:var(--muted);font-weight:600}
  .card .v{font-family:var(--mono);font-weight:700;font-size:30px;line-height:1.1;margin:8px 0 2px;letter-spacing:-.02em}
  .card .v small{font-size:14px;font-weight:600}
  .card .sub{font-size:11.5px;color:var(--faint);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
  .cnew .v{color:var(--new-line)} .csold .v{color:var(--sold-line)}
  .cup .v{color:var(--up)} .cdown .v{color:var(--down)}

  /* 전 ETF 교차 요약 */
  .xsum{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:14px 16px;margin:18px 0 6px}
  .xsum-head{font-size:13px;font-weight:800;margin-bottom:10px}
  .xsum-head .dim{color:var(--muted);font-weight:500;font-family:var(--mono);font-size:12px;margin-left:8px}
  .xsum-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px}
  .xsum-col{border:1px solid var(--line);border-radius:10px;padding:10px 12px}
  .xsum-col.buy{border-left:4px solid var(--new-line)} .xsum-col.sell{border-left:4px solid var(--sold-line)}
  .xc-title{font-size:12px;font-weight:700;margin-bottom:8px;display:flex;justify-content:space-between;align-items:center}
  .xsum-col.buy .xc-title{color:var(--new-line)} .xsum-col.sell .xc-title{color:var(--sold-line)}
  .xc-title span{font-size:11px;color:var(--muted);font-weight:600}
  .xsum-col ul{list-style:none;margin:0;padding:0;max-height:200px;overflow:auto}
  .xsum-col li{padding:5px 0;border-bottom:1px solid var(--line2);font-size:12.5px}
  .xsum-col li:last-child{border-bottom:0}
  .xc-cnt{font-family:var(--mono);font-weight:700;font-size:11px;color:var(--accent);margin:0 4px}
  .xc-etfs{color:var(--faint);font-size:11px}
  .xc-empty{color:var(--muted);font-size:12px;padding:8px 0}

  /* ETF 선택(운용사별 두 줄) */
  .etf-area{display:flex;flex-direction:column;gap:8px;margin:8px 0 4px}
  .etf-row{display:flex;align-items:center;gap:12px;flex-wrap:wrap}
  .etf-glabel{font-size:12px;font-weight:800;color:var(--accent);min-width:50px;letter-spacing:.04em}
  .etf-btns{display:flex;gap:6px;flex-wrap:wrap}
  .etf-btns button{appearance:none;border:1px solid var(--line);background:var(--panel);
    padding:7px 12px;border-radius:8px;font-family:var(--sans);font-size:13px;font-weight:600;
    color:var(--muted);cursor:pointer;white-space:nowrap}
  .etf-btns button:hover{border-color:var(--accent);color:var(--ink)}
  .etf-btns button.on{background:var(--accent);color:#16202e;border-color:var(--accent)}

  .tabs{display:flex;gap:2px;margin:24px 0 0;border-bottom:1px solid var(--line)}
  .tabs button{appearance:none;border:0;background:transparent;padding:11px 16px 12px;font-family:var(--sans);
    font-size:13.5px;font-weight:600;color:var(--muted);cursor:pointer;border-bottom:2px solid transparent;margin-bottom:-1px}
  .tabs button.on{color:var(--accent);border-bottom-color:var(--accent)}
  .panel-wrap{background:var(--panel);border:1px solid var(--line);border-top:0;border-radius:0 0 12px 12px;padding:18px 18px 22px;min-height:380px}
  .tabpanel{display:none} .tabpanel.on{display:block}
  .tabpanel h3{font-size:14px;margin:2px 0 10px;font-weight:700}
  .tabpanel h3 .dim{color:var(--muted);font-weight:500;font-family:var(--mono);font-size:12.5px;margin-left:8px}
  .hint{font-size:12px;color:var(--faint);margin:0 0 14px}

  table{width:100%;border-collapse:collapse;font-size:13px}
  thead th{position:sticky;top:0;background:var(--accent-soft);color:var(--accent);font-weight:700;
    font-size:11.5px;letter-spacing:.02em;text-align:right;padding:9px 12px;border-bottom:1px solid var(--line);white-space:nowrap}
  thead th.l{text-align:left}
  thead th.sortable{cursor:pointer;user-select:none}
  thead th.sortable:hover{color:#a7c5ff}
  thead th .sarrow{font-size:9px;opacity:.4;margin-left:3px}
  thead th.sortable.active .sarrow{opacity:1}
  tbody td{padding:8px 12px;border-bottom:1px solid var(--line2);text-align:right;font-family:var(--mono);font-variant-numeric:tabular-nums;white-space:nowrap}
  tbody td.l{text-align:left;font-family:var(--sans)}
  tbody tr:hover{background:var(--hover)}
  tr.r-new{background:var(--new-bg)} tr.r-new:hover{background:rgba(63,185,80,.20)}
  tr.r-sold{background:var(--sold-bg)} tr.r-sold:hover{background:rgba(248,81,73,.20)}
  tr.r-new td:first-child{box-shadow:inset 3px 0 0 var(--new-line)}
  tr.r-sold td:first-child{box-shadow:inset 3px 0 0 var(--sold-line)}
  tr.flash td{animation:flashrow 1.8s ease-out}
  @keyframes flashrow{0%,28%{background:var(--accent)}100%{background:transparent}}
  .tag{display:inline-block;font-family:var(--sans);font-size:11px;font-weight:700;padding:2px 8px;border-radius:20px;letter-spacing:.02em}
  .tag.new{background:var(--new-line);color:#16202e}
  .tag.sold{background:var(--sold-line);color:#16202e}
  .tag.expand{background:var(--exp-bg);color:var(--exp-text);border:1px solid var(--exp-bd)}
  .tag.contract{background:var(--con-bg);color:var(--con-text);border:1px solid var(--con-bd)}
  .tag.hold{background:#2c384c;color:var(--hold)}
  .code{color:var(--faint);font-size:11.5px}
  .up{color:var(--up)} .down{color:var(--down)} .zero{color:var(--faint)}
  .name b{font-weight:600}
  .scroll{max-height:560px;overflow:auto;border:1px solid var(--line);border-radius:10px}
  .empty{padding:46px 20px;text-align:center;color:var(--muted);font-size:13.5px}
  .period-block{margin-bottom:26px}
  .legend{display:flex;gap:14px;flex-wrap:wrap;margin:0 0 12px;font-size:12px;color:var(--muted)}
  .legend span{display:inline-flex;align-items:center;gap:6px}
  .dot{width:10px;height:10px;border-radius:50%;display:inline-block}
  footer{margin-top:26px;font-size:11.5px;color:var(--faint);text-align:right}
  .chart{width:100%}
  @media(max-width:880px){.cards{grid-template-columns:repeat(2,1fr)}}
</style>
</head>
<body>
<div class="wrap">
  <header>
    <div class="brand">
      <span class="eyebrow">ETF Holdings Monitor · 수량 기준</span>
      <h1>구성종목 일별 변동 대시보드</h1>
    </div>
    <div class="search-box">
      <span class="sicon">🔍</span>
      <input id="tickerSearch" type="text" placeholder="티커·종목명 검색" autocomplete="off" spellcheck="false">
      <span class="sclear" id="searchClear" title="지우기">✕</span>
      <div class="sug" id="sug"></div>
    </div>
    <select id="dateSel" style="display:none"></select>
    <span id="baseNote" style="display:none"></span>
  </header>

  <section class="spanel" id="searchPanel"></section>

  <section class="xsum" id="xsum"></section>

  <section class="etf-area" id="etfSeg"></section>

  <section class="cards" id="cards"></section>

  <nav class="tabs" id="tabs">
    <button class="on" data-tab="0">일일 변경 내역</button>
    <button data-tab="1">신규 · 매도 타임라인</button>
    <button data-tab="2">기간 비교 (1주 · 1개월)</button>
  </nav>
  <div class="panel-wrap">
    <div class="tabpanel on" id="tab0"></div>
    <div class="tabpanel" id="tab1">
      <h3>편출입 타임라인 <span class="dim">신규=초록, 전액 매도=빨강</span></h3>
      <div class="legend">
        <span><i class="dot" style="background:var(--new-line)"></i>신규 편입</span>
        <span><i class="dot" style="background:var(--sold-line)"></i>전액 매도(제외)</span>
      </div>
      <div id="timelineChart" class="chart"></div>
    </div>
    <div class="tabpanel" id="tab3"></div>
  </div>

  <footer>생성: %%GENERATED%% · 등락 색상은 한국 시장 관례(상승 빨강 / 하락 파랑)</footer>
</div>

<script src="%%PLOTLY_SRC%%"></script>
<script>
const ETF_DATA = %%DATA_JSON%%;
const DEFAULT_ETF = %%DEFAULT_ETF%%;
const EXPAND_PCT = 0.10;   // 수량 증감 ±10% 기준
const NO_SORT = ()=>({ day:{col:null,dir:"desc"}, wk:{col:null,dir:"desc"}, mo:{col:null,dir:"desc"} });
const state = { etf: DEFAULT_ETF, dateIdx: 0, tab: 0, sorts: NO_SORT(), searchKey: null, highlight: null };

// 표 컬럼 정의 (헤더 순서 = 본문 순서)
const COLS = [
  {key:"status", label:"상태",     type:"status", cls:"l"},
  {key:"name",   label:"종목명",    type:"str",    cls:"l"},
  {key:"code",   label:"코드",      type:"str",    cls:"l"},
  {key:"qtyB",   label:"수량",      type:"num", dateSlot:"b"},
  {key:"qtyT",   label:"수량",      type:"num", dateSlot:"c"},
  {key:"dQty",   label:"수량 증감", type:"num"},
  {key:"wB",     label:"비중",      type:"num", dateSlot:"b"},
  {key:"wT",     label:"비중",      type:"num", dateSlot:"c"},
  {key:"dW",     label:"비중 증감", type:"num"},
];
const STATUS_RANK = {NEW:0, SOLD:1, EXPAND:2, CONTRACT:3, HELD:4};
const TAG = {NEW:["new","신규"], SOLD:["sold","제외"], EXPAND:["expand","확대"], CONTRACT:["contract","축소"], HELD:["hold","유지"]};

// ---------- helpers ----------
const fmtInt = n => (n===null||n===undefined||isNaN(n)) ? "–" : Math.round(n).toLocaleString("ko-KR");
const fmtW   = n => (n===null||n===undefined||isNaN(n)) ? "–" : n.toFixed(2);
function signNum(n, isW){
  if(n===0||Math.abs(n)<(isW?0.005:0.5)) return `<span class="zero">0${isW?'%p':''}</span>`;
  const cls = n>0 ? "up":"down", s = n>0?"+":"−", v = Math.abs(n);
  return `<span class="${cls}">${s}${isW?v.toFixed(2)+'%p':fmtInt(v)}</span>`;
}
const esc = s => String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
function nameList(names){
  if(!names.length) return "";
  const head = names.slice(0,3).map(esc).join(", ");
  return head + (names.length>3 ? ` 외 ${names.length-3}종목` : "");
}

function dates(){ return ETF_DATA[state.etf].dates; }
function holdings(d){ return ETF_DATA[state.etf].holdings[d] || {}; }
function holdingsOf(etf, d){ return ETF_DATA[etf].holdings[d] || {}; }
function isActive(o){ return !!o && (o.qty!==0 || o.value!==0); }

// 운용사 그룹/접두어
function etfGroup(key){ return key.startsWith("KoAct_") ? "KoAct" : key.startsWith("Time_") ? "Time" : "기타"; }
function stripPrefix(key){ return key.replace(/^Time_|^KoAct_/, ""); }

// 한 종목의 전일 대비 상태 분류 (computeDiff와 동일 규칙)
function classify(t, b){
  const ta=isActive(t), ba=isActive(b);
  if(!ta && !ba) return null;
  if(ta && !ba) return "NEW";
  if(!ta && ba) return "SOLD";
  const qtyT=t?t.qty:0, qtyB=b?b.qty:0;
  if(qtyB>0){ const p=(qtyT-qtyB)/qtyB; return p>=EXPAND_PCT?"EXPAND":p<=-EXPAND_PCT?"CONTRACT":"HELD"; }
  return "HELD";
}

// 현금·선물·지수류 제외
const XSUM_EXCL_NAME = ["현금","cash","krw","mini","e-mini","선물","옵션","index","future"];
function isExcludedSec(code, name){
  if(!code) return true;                                   // 코드 없음(현금 등)
  const c=String(code).toLowerCase(), nm=String(name||"").toLowerCase();
  if(/index|future/.test(c)) return true;                  // 코드에 Index/Futures
  if(XSUM_EXCL_NAME.some(k=>nm.includes(k))) return true;  // 이름 키워드
  return false;
}

function shiftDate(iso, opt){
  const [y,m,d]=iso.split("-").map(Number);
  const dt=new Date(y, m-1, d);
  if(opt.days) dt.setDate(dt.getDate()-opt.days);
  if(opt.months) dt.setMonth(dt.getMonth()-opt.months);
  const p=n=>String(n).padStart(2,"0");
  return `${dt.getFullYear()}-${p(dt.getMonth()+1)}-${p(dt.getDate())}`;
}
function baselineDate(idx, mode){
  const ds=dates();
  if(mode==='prev') return idx>0 ? ds[idx-1] : null;
  const cur=ds[idx];
  const target = mode==='week' ? shiftDate(cur,{days:7}) : shiftDate(cur,{months:1});
  let res=null;
  for(let i=0;i<idx;i++){ if(ds[i]<=target) res=ds[i]; }
  return res;
}

function computeDiff(curDate, baseDate){
  const T=holdings(curDate), B=baseDate?holdings(baseDate):{};
  const keys=new Set([...Object.keys(T),...Object.keys(B)]);
  const rows=[]; let newC=0, soldC=0, expC=0, conC=0;
  keys.forEach(k=>{
    const t=T[k], b=B[k], ta=isActive(t), ba=isActive(b);
    if(!ta && !ba) return;
    const qtyT=t?t.qty:0, qtyB=b?b.qty:0, wT=t?t.weight:0, wB=b?b.weight:0;
    let status, pct=null;
    if(ta && !ba) status="NEW";
    else if(!ta && ba) status="SOLD";
    else if(qtyB>0){
      pct=(qtyT-qtyB)/qtyB;
      status = pct>=EXPAND_PCT ? "EXPAND" : pct<=-EXPAND_PCT ? "CONTRACT" : "HELD";
    } else status="HELD";
    rows.push({key:k, name:(t?t.name:b.name), code:(t?t.code:b.code), status, pct,
      qtyB, qtyT, dQty:qtyT-qtyB, wB, wT, dW:+(wT-wB).toFixed(2)});
    if(status==="NEW")newC++; else if(status==="SOLD")soldC++;
    else if(status==="EXPAND")expC++; else if(status==="CONTRACT")conC++;
  });
  rows.sort((a,b)=> STATUS_RANK[a.status]-STATUS_RANK[b.status] || Math.abs(b.dW)-Math.abs(a.dW));
  const up = rows.reduce((m,r)=> r.dW>(m?m.dW:-1e9)?r:m, null);
  const dn = rows.reduce((m,r)=> r.dW<(m?m.dW: 1e9)?r:m, null);
  return {rows, newC, soldC, expC, conC, up, dn};
}

function sortRows(rows, col, dir){
  const meta = COLS.find(c=>c.key===col); const sgn = dir==="asc"?1:-1;
  rows.sort((a,b)=>{
    if(meta.type==="str")    return String(a[col]||"").localeCompare(String(b[col]||""),"ko")*sgn;
    if(meta.type==="status") return (STATUS_RANK[a.status]-STATUS_RANK[b.status])*sgn;
    return ((a[col]||0)-(b[col]||0))*sgn;
  });
}

// ---------- cross-ETF summary ----------
function renderCrossSummary(){
  const sel = dates()[state.dateIdx];           // 현재 선택 기준일(문자열)
  const buy={}, sell={};                          // name -> {code, name, etfs:[...]}
  Object.keys(ETF_DATA).forEach(etf=>{
    const ds = ETF_DATA[etf].dates;
    let cur=null; for(const d of ds){ if(d<=sel) cur=d; }   // 이 ETF의 기준일(<=sel 중 최신)
    if(!cur) return;
    const idx=ds.indexOf(cur), base = idx>0 ? ds[idx-1] : null;
    if(!base) return;
    const T=holdingsOf(etf,cur), B=holdingsOf(etf,base);
    const keys=new Set([...Object.keys(T),...Object.keys(B)]);
    keys.forEach(k=>{
      const st=classify(T[k],B[k]); if(!st) return;
      const o=T[k]||B[k]; const code=o.code, name=o.name;
      if(isExcludedSec(code,name)) return;
      const dir = (st==="NEW"||st==="EXPAND") ? "buy" : (st==="SOLD"||st==="CONTRACT") ? "sell" : null;
      if(!dir) return;
      // 코드가 있으면 코드(케이스·공백 정규화)로, 없으면(현금 등) 이름으로 매칭
      const mapKey = code
        ? code.trim().toUpperCase().replace(/\s+/g, " ")
        : name.trim().toUpperCase().replace(/\s+/g, " ");
      const map = dir==="buy"?buy:sell;
      if(!map[mapKey]) map[mapKey]={code, name, etfs:[]};
      if(!map[mapKey].etfs.some(e=>e.etf===etf)){
        // 코드가 있고 기존 코드가 비어있으면 업데이트 (더 구체적인 코드 우선)
        if(code && !map[mapKey].code) map[mapKey].code = code;
        map[mapKey].etfs.push({etf, label:stripPrefix(etf), isTime:etf.startsWith("Time_")});
      }
    });
  });
  const build = map => Object.values(map)
    .filter(o=>o.etfs.length>=2)                   // 서로 다른 ETF 2개 이상
    .map(o=>{
      return {name:o.name, code:o.code, count:o.etfs.length, etfs:o.etfs.map(e=>e.label)};
    })
    .sort((a,b)=> b.count-a.count || a.name.localeCompare(b.name,"ko"));
  const buyList=build(buy), sellList=build(sell);
  const li = it => `<li><b>${esc(it.name)}</b><span class="xc-cnt">(${it.count})</span><span class="xc-etfs">${it.etfs.map(esc).join(", ")}</span></li>`;
  const col = (cls,title,list,empty) =>
    `<div class="xsum-col ${cls}"><div class="xc-title">${title}<span>${list.length}종목</span></div>`
    + (list.length? `<ul>${list.map(li).join("")}</ul>` : `<div class="xc-empty">${empty}</div>`) + `</div>`;
  document.getElementById("xsum").innerHTML =
    `<div class="xsum-head">전 ETF 교차 요약<span class="dim">${sel} · 전일 대비 · 2개 이상 ETF 동시 · 현금/선물 제외</span></div>
     <div class="xsum-grid">
       ${col("buy","📈 매수 방향 (신규·확대)", buyList, "2개 이상 ETF가 동시 매수한 종목 없음")}
       ${col("sell","📉 매도 방향 (제외·축소)", sellList, "2개 이상 ETF가 동시 매도한 종목 없음")}
     </div>`;
}

// ---------- cards ----------
function renderCards(){
  const cur=dates()[state.dateIdx], base=baselineDate(state.dateIdx,'prev');
  const d=computeDiff(cur, base);
  const newNames=d.rows.filter(r=>r.status==="NEW").map(r=>r.name);
  const soldNames=d.rows.filter(r=>r.status==="SOLD").map(r=>r.name);
  const up=d.up, dn=d.dn;
  document.getElementById("cards").innerHTML = `
    <div class="card cnew"><div class="k">🆕 신규 편입</div>
      <div class="v">${d.newC}<small> 종목</small></div>
      <div class="sub" title="${newNames.map(esc).join(', ')}">${d.newC? nameList(newNames) : "신규 없음"}</div></div>
    <div class="card csold"><div class="k">❌ 전액 매도(제외)</div>
      <div class="v">${d.soldC}<small> 종목</small></div>
      <div class="sub" title="${soldNames.map(esc).join(', ')}">${d.soldC? nameList(soldNames) : "제외 없음"}</div></div>
    <div class="card cup"><div class="k">📈 비중 최대 증가</div>
      <div class="v">${up&&up.dW>0?"+"+up.dW.toFixed(2):"–"}<small>${up&&up.dW>0?"%p":""}</small></div>
      <div class="sub">${up&&up.dW>0? esc(up.name) : "증가 종목 없음"}</div></div>
    <div class="card cdown"><div class="k">📉 비중 최대 감소</div>
      <div class="v">${dn&&dn.dW<0?"−"+Math.abs(dn.dW).toFixed(2):"–"}<small>${dn&&dn.dW<0?"%p":""}</small></div>
      <div class="sub">${dn&&dn.dW<0? esc(dn.name) : "감소 종목 없음"}</div></div>`;
  document.getElementById("baseNote").textContent = base ? `vs ${base}` : "이전 영업일 데이터 없음";
}

// ---------- table ----------
function headerHTML(baseDate, curDate, sortable, scope, ss){
  const bl=baseDate.slice(5), cl=curDate.slice(5);
  return COLS.map(c=>{
    let label=c.label;
    if(c.dateSlot==="b") label=`${c.label} (${bl})`;
    if(c.dateSlot==="c") label=`${c.label} (${cl})`;
    const base=c.cls==="l"?"l":"";
    if(!sortable) return `<th class="${base}">${label}</th>`;
    const active=ss && ss.col===c.key;
    const arrow=active?(ss.dir==="asc"?"▲":"▼"):"↕";
    return `<th class="${base} sortable${active?' active':''}" onclick="setSort('${scope}','${c.key}')">${label}<span class="sarrow">${arrow}</span></th>`;
  }).join("");
}
function rowHTML(r){
  const cls = r.status==="NEW"?"r-new":r.status==="SOLD"?"r-sold":"";
  const [tc,tl]=TAG[r.status];
  const pctTitle = (r.status==="EXPAND"||r.status==="CONTRACT")&&r.pct!=null
    ? ` title="수량 ${(r.pct>0?'+':'')}${(r.pct*100).toFixed(1)}%"` : "";
  return `<tr class="${cls}" data-code="${esc(r.code||"")}" data-name="${esc(r.name||"")}">
    <td class="l"><span class="tag ${tc}"${pctTitle}>${tl}</span></td>
    <td class="l name"><b>${esc(r.name)}</b></td>
    <td class="l code">${esc(r.code||"")}</td>
    <td>${r.status==="NEW"?'<span class="zero">–</span>':fmtInt(r.qtyB)}</td>
    <td>${r.status==="SOLD"?'<span class="zero">–</span>':fmtInt(r.qtyT)}</td>
    <td>${signNum(r.dQty,false)}</td>
    <td>${r.status==="NEW"?'<span class="zero">–</span>':fmtW(r.wB)}</td>
    <td>${r.status==="SOLD"?'<span class="zero">–</span>':fmtW(r.wT)}</td>
    <td>${signNum(r.dW,true)}</td></tr>`;
}
function tableHTML(rows, baseDate, curDate, opts){
  opts = opts || {};
  if(!baseDate) return `<div class="empty">비교할 기준일 데이터가 없습니다.<br>(가장 과거 날짜이거나 해당 기간 이전 파일이 없습니다)</div>`;
  if(!rows.length) return `<div class="empty">변동 내역이 없습니다.</div>`;
  return `<div class="scroll"><table>
    <thead><tr>${headerHTML(baseDate,curDate,opts.sortable,opts.scope,opts.sortState)}</tr></thead>
    <tbody>${rows.map(rowHTML).join("")}</tbody></table></div>`;
}

function setSort(scope, col){
  const s = state.sorts[scope];
  if(s.col===col){ s.dir = s.dir==="desc"?"asc":"desc"; }
  else { state.sorts[scope] = {col, dir:"desc"}; }
  if(scope==="day") renderTab0(); else renderTab3();
}

function renderTab0(){
  const cur=dates()[state.dateIdx], base=baselineDate(state.dateIdx,'prev');
  const d=computeDiff(cur, base);
  let rows=d.rows.slice();
  const ss=state.sorts.day;
  if(ss.col) sortRows(rows, ss.col, ss.dir);
  const sortMsg = ss.col
    ? `정렬: ${COLS.find(c=>c.key===ss.col).label} ${ss.dir==="asc"?"오름차순▲":"내림차순▼"} · 기본 정렬은 기준일 재선택`
    : "아무 열 머리글이나 누르면 정렬됩니다.";
  const tags = `<span class="legend" style="margin-bottom:6px">
    <span><span class="tag new">신규</span>편입</span>
    <span><span class="tag sold">제외</span>전액매도</span>
    <span><span class="tag expand">확대</span>수량 +10%↑</span>
    <span><span class="tag contract">축소</span>수량 −10%↓</span></span>`;
  document.getElementById("tab0").innerHTML =
    `<h3>일일 포트폴리오 변경 내역<span class="dim">${base?base+" → "+cur:cur}</span></h3>`
    + tags + `<p class="hint">${sortMsg}</p>`
    + tableHTML(rows, base, cur, {sortable:true, scope:"day", sortState:ss});
}

function renderTab3(){
  const cur=dates()[state.dateIdx];
  const wk=baselineDate(state.dateIdx,'week'), mo=baselineDate(state.dateIdx,'month');
  const blk=(title, base, scope)=>{
    const d=computeDiff(cur, base);
    const meta = base? `${base} → ${cur}` : "기간 이전 데이터 없음";
    let rows=d.rows.slice();
    const ss=state.sorts[scope];
    if(base && ss.col) sortRows(rows, ss.col, ss.dir);
    const hint = base
      ? `<p class="hint">${ss.col ? `정렬: ${COLS.find(c=>c.key===ss.col).label} ${ss.dir==="asc"?"오름차순▲":"내림차순▼"} · 기본 정렬은 기준일 재선택` : "아무 열 머리글이나 누르면 정렬됩니다."}</p>`
      : "";
    return `<div class="period-block"><h3>${title}<span class="dim">${meta}</span></h3>`
      + hint
      + tableHTML(rows, base, cur, {sortable:!!base, scope, sortState:ss}) + `</div>`;
  };
  document.getElementById("tab3").innerHTML = blk("1주 전 대비", wk, "wk") + blk("1개월 전 대비", mo, "mo");
}

// ---------- ticker search ----------
function refDate(){ return dates()[state.dateIdx]; }   // 현재 선택 기준일(선택 ETF 기준)

// 각 ETF의 asOf(<=ref 중 최신)와 그 직전일
function etfAsOf(etf, ref){
  const ds=ETF_DATA[etf].dates; let cur=null;
  for(const d of ds){ if(d<=ref) cur=d; }
  if(!cur) return null;
  const i=ds.indexOf(cur);
  return {cur, prev: i>0 ? ds[i-1] : null};
}

let _idxCache={ref:null, idx:null};
function buildSearchIndex(ref){
  if(_idxCache.ref===ref && _idxCache.idx) return _idxCache.idx;
  const idx={};   // mapKey -> {code, name, holders:[...]}
  Object.keys(ETF_DATA).forEach(etf=>{
    const ao=etfAsOf(etf, ref); if(!ao) return;
    const T=holdingsOf(etf, ao.cur), B=ao.prev?holdingsOf(etf, ao.prev):{};
    // 이 ETF의 활성 보유 → 비중 순위 산정
    const active=Object.values(T).filter(isActive).slice().sort((a,b)=>(b.weight||0)-(a.weight||0));
    const total=active.length;
    const rankOf=new Map();
    active.forEach((o,i)=> rankOf.set(o.code ? o.code+"|"+o.name : o.name, i+1));
    Object.keys(T).forEach(k=>{
      const o=T[k]; if(!isActive(o)) return;
      const code=(o.code||"").trim(), name=(o.name||"").trim();
      const mapKey = code ? code.toUpperCase().replace(/\s+/g," ") : name.toUpperCase().replace(/\s+/g," ");
      if(!idx[mapKey]) idx[mapKey]={code, name, holders:[]};
      if(code && !idx[mapKey].code) idx[mapKey].code=code;
      idx[mapKey].holders.push({
        etf, label:stripPrefix(etf), group:etfGroup(etf),
        weight:o.weight||0, qty:o.qty||0,
        rank:rankOf.get(o.code ? o.code+"|"+o.name : o.name)||null, total,
        status:classify(T[k], B[k]) || "HELD",
        code:o.code||"", name:o.name
      });
    });
  });
  _idxCache={ref, idx};
  return idx;
}

function searchMatches(q){
  q=q.trim().toLowerCase(); if(!q) return [];
  const idx=buildSearchIndex(refDate());
  const out=[];
  Object.keys(idx).forEach(mk=>{
    const e=idx[mk];
    const nm=(e.name||"").toLowerCase(), cd=(e.code||"").toLowerCase();
    if(nm.includes(q) || cd.includes(q))
      out.push({mapKey:mk, code:e.code, name:e.name, count:e.holders.length});
  });
  out.sort((a,b)=>{
    const as=(a.name||"").toLowerCase().startsWith(q)?0:1;
    const bs=(b.name||"").toLowerCase().startsWith(q)?0:1;
    return as-bs || b.count-a.count || (a.name||"").localeCompare(b.name||"","ko");
  });
  return out;
}

function renderSug(q){
  const box=document.getElementById("sug");
  document.getElementById("searchClear").style.display = q ? "block" : "none";
  if(!q.trim()){ box.classList.remove("on"); box.innerHTML=""; return; }
  const matches=searchMatches(q).slice(0,40);
  if(!matches.length){
    box.innerHTML=`<div class="sug-empty">일치하는 종목이 없어요.</div>`;
    box.classList.add("on"); return;
  }
  box.innerHTML=matches.map(m=>
    `<div class="sug-item" data-key="${esc(m.mapKey)}">
       <span class="sug-name">${esc(m.name)}${m.code?`<span class="code">${esc(m.code)}</span>`:""}</span>
       <span class="sug-cnt">${m.count}개 ETF</span>
     </div>`).join("");
  box.classList.add("on");
  box.querySelectorAll(".sug-item").forEach(it=> it.onclick=()=> selectSecurity(it.dataset.key));
}

function selectSecurity(mapKey){
  state.searchKey=mapKey;
  document.getElementById("sug").classList.remove("on");
  renderSearchPanel();
  document.getElementById("searchPanel").scrollIntoView({behavior:"smooth", block:"start"});
}

function clearSearch(){
  state.searchKey=null;
  const inp=document.getElementById("tickerSearch");
  inp.value="";
  document.getElementById("searchClear").style.display="none";
  document.getElementById("sug").classList.remove("on");
  renderSearchPanel();
}

function renderSearchPanel(){
  const panel=document.getElementById("searchPanel");
  if(!state.searchKey){ panel.classList.remove("on"); panel.innerHTML=""; return; }
  const e=buildSearchIndex(refDate())[state.searchKey];
  panel.classList.add("on");
  if(!e){   // 기준일이 바뀌어 현재 시점엔 어떤 ETF도 보유하지 않는 경우
    panel.innerHTML=`<div class="sp-head"><div>
        <div class="sp-title">${esc(state.searchKey)}</div>
        <div class="sp-sub">${esc(refDate())} 기준으로 이 종목을 보유한 ETF가 없어요.</div></div>
      <button class="sp-close" onclick="clearSearch()">닫기</button></div>`;
    return;
  }
  const holders=e.holders.slice().sort((a,b)=>(b.weight||0)-(a.weight||0));
  const rows=holders.map(h=>{
    const [tc,tl]=TAG[h.status]||TAG.HELD;
    return `<div class="holder-row" data-etf="${esc(h.etf)}" data-code="${esc(h.code||"")}" data-name="${esc(h.name||"")}">
      <span class="h-etf"><span class="gbadge ${h.group==='KoAct'?'koact':'time'}">${esc(h.group)}</span><span class="lbl">${esc(h.label)}</span></span>
      <span class="h-num">${h.weight? h.weight.toFixed(2)+'%' : '–'}</span>
      <span class="h-num">${fmtInt(h.qty)}</span>
      <span class="h-num h-rank">${h.rank? `${h.rank}위<span class="code"> / ${h.total}</span>` : '–'}</span>
      <span><span class="tag ${tc}">${tl}</span></span>
      <span class="h-jump">이동 →</span>
    </div>`;
  }).join("");
  panel.innerHTML=`
    <div class="sp-head">
      <div>
        <div class="sp-title">${esc(e.name)}${e.code?`<span class="code">${esc(e.code)}</span>`:""}</div>
        <div class="sp-sub"><b>${holders.length}개 ETF</b> 보유 · ${esc(refDate())} 기준 · 상태는 전일 대비</div>
      </div>
      <button class="sp-close" onclick="clearSearch()">닫기</button>
    </div>
    <div class="holder-row head">
      <span>ETF</span><span class="h-num">비중</span><span class="h-num">수량</span>
      <span class="h-num h-rank">ETF내 순위</span><span>전일 대비</span><span class="h-jump"></span>
    </div>
    ${rows}`;
  panel.querySelectorAll(".holder-row:not(.head)").forEach(r=>
    r.onclick=()=> jumpToEtf(r.dataset.etf, r.dataset.code, r.dataset.name));
}

function jumpToEtf(etf, code, name){
  const keepDate=dates()[state.dateIdx];        // 현재 기준일 유지 시도
  state.etf=etf; state.sorts=NO_SORT();
  document.querySelectorAll("#etfSeg button").forEach(b=> b.classList.toggle("on", b.dataset.etf===etf));
  buildDateSel(keepDate);
  // 일일 변경 내역 탭으로 전환
  state.tab=0;
  document.querySelectorAll("#tabs button").forEach(b=> b.classList.toggle("on", +b.dataset.tab===0));
  document.querySelectorAll(".tabpanel").forEach((p,i)=> p.classList.toggle("on", i===0));
  state.highlight={code:(code||"").trim().toUpperCase(), name:(name||"").trim().toUpperCase()};
  renderAll();
  requestAnimationFrame(highlightRow);
}

function highlightRow(){
  if(!state.highlight) return;
  const {code, name}=state.highlight;
  let target=null;
  document.querySelectorAll("#tab0 tbody tr").forEach(tr=>{
    if(target) return;
    const c=(tr.dataset.code||"").trim().toUpperCase();
    const n=(tr.dataset.name||"").trim().toUpperCase();
    if(code ? c===code : n===name) target=tr;
  });
  state.highlight=null;
  if(!target) return;
  target.scrollIntoView({behavior:"smooth", block:"center"});
  target.classList.remove("flash"); void target.offsetWidth; target.classList.add("flash");
}

// ---------- charts ----------
const PLOT_DARK = {
  font:{family:"'Pretendard',sans-serif",size:12,color:"#e9ecf3"},
  plot_bgcolor:"#222c3e", paper_bgcolor:"#222c3e", grid:"#374459", grid2:"#2c384c", zero:"#46536b"
};
const hasPlotly = ()=> typeof Plotly !== "undefined";
function chartUnavailable(id){
  document.getElementById(id).innerHTML =
    `<div class="empty">차트 라이브러리를 불러오지 못했습니다.<br>인터넷에 연결해 다시 열거나, 폴더의 plotly.min.js를 확인하세요.</div>`;
}

function renderTimeline(){
  const el="timelineChart";
  if(!hasPlotly()) return chartUnavailable(el);
  const ds=dates();
  const newX=[],newY=[],soldX=[],soldY=[], names=new Set();
  for(let i=1;i<ds.length;i++){
    const P=holdings(ds[i-1]), C=holdings(ds[i]);
    const keys=new Set([...Object.keys(P),...Object.keys(C)]);
    keys.forEach(k=>{
      const ca=isActive(C[k]), pa=isActive(P[k]);
      const nm=(C[k]?C[k].name:(P[k]?P[k].name:k));   // 종목명으로 표기
      if(ca&&!pa){ newX.push(ds[i]); newY.push(nm); names.add(nm); }
      if(!ca&&pa){ soldX.push(ds[i]); soldY.push(nm); names.add(nm); }
    });
  }
  if(!newX.length && !soldX.length){
    document.getElementById(el).innerHTML=`<div class="empty">기간 내 편출입 이벤트가 없습니다.</div>`; return;
  }
  const order=[...names].sort((a,b)=>a.localeCompare(b,"ko"));
  const h=Math.max(360, order.length*26+120);
  const base={mode:"markers", marker:{size:12, line:{width:1,color:"#222c3e"}}, hovertemplate:"%{y}<br>%{x}<br>%{text}<extra></extra>"};
  Plotly.newPlot(el, [
    {...base, name:"신규 편입", x:newX, y:newY, text:newX.map(()=>"신규 편입"),
      marker:{...base.marker, color:"#3fb950", symbol:"circle"}},
    {...base, name:"전액 매도", x:soldX, y:soldY, text:soldX.map(()=>"전액 매도"),
      marker:{...base.marker, color:"#f85149", symbol:"x"}}
  ],{
    height:h, margin:{l:230,r:24,t:10,b:46}, hovermode:"closest",
    font:PLOT_DARK.font, plot_bgcolor:PLOT_DARK.plot_bgcolor, paper_bgcolor:PLOT_DARK.paper_bgcolor,
    xaxis:{type:"category", categoryorder:"category ascending", tickangle:-40, gridcolor:PLOT_DARK.grid},
    yaxis:{type:"category", categoryarray:order, autorange:"reversed", gridcolor:PLOT_DARK.grid2, automargin:true},
    legend:{orientation:"h", y:1.06, x:0}, showlegend:true
  },{responsive:true, displaylogo:false});
}

// ---------- dispatch & events ----------
function renderActiveTab(){
  if(state.tab===0) renderTab0();
  else if(state.tab===1) renderTimeline();
  else if(state.tab===2) renderTab3();
}
function renderAll(){ renderCrossSummary(); renderCards(); renderActiveTab(); renderSearchPanel(); }

function buildEtfSeg(){
  const wrap=document.getElementById("etfSeg");
  const groups={}, order=[];
  Object.keys(ETF_DATA).forEach(k=>{
    const g=etfGroup(k);
    if(!groups[g]){ groups[g]=[]; order.push(g); }
    groups[g].push(k);
  });
  wrap.innerHTML = order.map(g=>
    `<div class="etf-row"><span class="etf-glabel">${esc(g)}</span><div class="etf-btns">`
    + groups[g].map(k=>`<button data-etf="${esc(k)}" class="${k===state.etf?'on':''}">${esc(stripPrefix(k))}</button>`).join("")
    + `</div></div>`).join("");
  wrap.querySelectorAll("button").forEach(b=> b.onclick=()=>{
    const keepDate = dates()[state.dateIdx];   // 현재 선택한 기준일 유지
    state.etf=b.dataset.etf; state.sorts=NO_SORT();
    buildDateSel(keepDate);                    // 같은 날짜가 있으면 그대로, 없으면 최신일
    wrap.querySelectorAll("button").forEach(x=>x.classList.toggle("on",x===b));
    renderAll();
  });
}
function buildDateSel(prefer){
  const sel=document.getElementById("dateSel"), ds=dates();
  sel.innerHTML = ds.map((d,i)=>`<option value="${i}">${d}</option>`).join("");
  let idx = ds.length-1;
  if(prefer){ const p=ds.indexOf(prefer); if(p>=0) idx=p; }
  state.dateIdx = idx;
  sel.value = idx;
}
function bindTabs(){
  document.querySelectorAll("#tabs button").forEach(b=> b.onclick=()=>{
    state.tab=+b.dataset.tab;
    document.querySelectorAll("#tabs button").forEach(x=>x.classList.toggle("on",x===b));
    document.querySelectorAll(".tabpanel").forEach((p,i)=>p.classList.toggle("on",i===state.tab));
    renderActiveTab();
    if(hasPlotly() && state.tab===1){ try{Plotly.Plots.resize(document.getElementById("timelineChart"));}catch(e){} }
  });
}
function bindSearch(){
  const inp=document.getElementById("tickerSearch");
  inp.addEventListener("input", e=> renderSug(e.target.value));
  inp.addEventListener("focus", e=>{ if(e.target.value.trim()) renderSug(e.target.value); });
  inp.addEventListener("keydown", e=>{
    if(e.key==="Enter"){ const first=document.querySelector("#sug .sug-item"); if(first) first.click(); }
    else if(e.key==="Escape"){ document.getElementById("sug").classList.remove("on"); inp.blur(); }
  });
  document.getElementById("searchClear").onclick=clearSearch;
  document.addEventListener("click", e=>{
    if(!e.target.closest(".search-box")) document.getElementById("sug").classList.remove("on");
  });
}
function init(){
  buildEtfSeg(); buildDateSel(); bindTabs(); bindSearch();
  document.getElementById("dateSel").onchange=e=>{ state.dateIdx=+e.target.value; state.sorts=NO_SORT(); renderAll(); };
  renderAll();
}
init();
</script>
</body>
</html>
"""


if __name__ == "__main__":
    out = generate()
    if out:
        print("\n저장 위치:", out.resolve())
    else:
        print("표시할 데이터가 없습니다. ETF_LIST 설정을 확인하세요.")
