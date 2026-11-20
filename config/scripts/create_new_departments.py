
import json
from pathlib import Path
from html import escape

ROOT = Path(__file__).resolve().parents[1]
CONFIG = ROOT / "config" / "new_departments.json"

with CONFIG.open(encoding="utf-8") as file:
    departments = json.load(file)["departments"]

for dept in departments:
    dept_id = dept["id"]
    name = escape(dept["name"])

    folder = ROOT / dept["report_path"]
    folder.mkdir(parents=True, exist_ok=True)

    page = folder / "index.html"

    # 이미 생성된 페이지는 덮어쓰지 않음
    if page.exists():
        print(f"SKIP: {page}")
        continue

    content = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport"
      content="width=device-width, initial-scale=1">
<title>Big Kids · {name}</title>
<style>
:root {{
  --bg:#f3f2f6;
  --ink:#1d1c22;
  --sub:#6a6a72;
  --bar:#1b1b22;
  --line:#e7e5ee;
}}
* {{ box-sizing:border-box; }}
body {{
  margin:0;
  font-family:-apple-system,"Segoe UI",
    "Apple SD Gothic Neo","Noto Sans KR",sans-serif;
  background:var(--bg);
  color:var(--ink);
}}
.topbar {{
  background:var(--bar);
  padding:14px 20px;
}}
.topbar a {{
  color:white;
  text-decoration:none;
  font-size:13px;
}}
.wrap {{
  max-width:900px;
  margin:0 auto;
  padding:26px 20px 56px;
}}
.hero {{
  background:#eef4fa;
  border-radius:16px;
  padding:24px;
  margin-bottom:20px;
}}
.hero h1 {{
  font-size:26px;
  margin:8px 0;
}}
.meta {{
  color:var(--sub);
  font-size:13px;
}}
.card {{
  background:white;
  border-radius:14px;
  padding:22px;
  margin-bottom:12px;
  box-shadow:0 1px 2px rgba(20,18,30,.06);
}}
.card h2 {{
  margin:0 0 10px;
  font-size:19px;
}}
.card p {{
  color:var(--sub);
  font-size:13px;
  line-height:1.6;
}}
.status {{
  display:inline-block;
  background:#f3f2f6;
  padding:5px 10px;
  border-radius:8px;
  font-size:12px;
}}
</style>
</head>
<body>

<div class="topbar">
  <a href="../../">📊 Big Kids Hub</a>
</div>

<div class="wrap">
  <div class="hero">
    <div class="meta">
      TARGET.COM · 격주 트래킹
    </div>
    <h1>{name}</h1>
    <div class="meta">
      Catalog Tracking + Guest Feedback
    </div>
  </div>

  <div class="card">
    <h2>🛒 T.COM 카탈로그 트래킹</h2>
    <p>
      신상품 · DROPPED · 재등장 · 가격변동
    </p>
    <span class="status">
      첫 데이터 수집 대기 중
    </span>
  </div>

  <div class="card">
    <h2>💬 Guest Feedback</h2>
    <p>
      신규 리뷰 · 누적 리뷰 · POS/NEG 분석
    </p>
    <span class="status">
      첫 데이터 수집 대기 중
    </span>
  </div>

</div>
</body>
</html>"""

    page.write_text(content, encoding="utf-8")
    print(f"CREATED: {page.relative_to(ROOT)}")

print("Finished creating new department pages.")
