import os, re, html

base = r"C:/Users/m1594/WorkBuddy/2026-06-25-04-43-16/outputs/portfolio"

files = {
    "elevator_gdd_master": "游戏设计总纲",
    "elevator_pm_master": "项目管理文档",
    "elevator_art_master": "美术需求文档",
    "elevator_audio_master": "音效需求文档",
}

CSS = """<style>
  :root {
    --bg: #f9f7f4; --card: #ffffff; --text: #1e1b18;
    --text-secondary: #5c554e; --text-muted: #9b9489; --accent: #b8753e;
    --accent-dim: #d4a76a; --accent-subtle: #f5ede0; --border: #e8e2d8;
    --highlight-bg: #faf7f0; --shadow-sm: 0 1px 3px rgba(80,60,30,0.04);
    --shadow-md: 0 4px 20px rgba(80,60,30,0.06);
  }
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, "PingFang SC", "Microsoft YaHei", sans-serif;
    background: var(--bg); color: var(--text); line-height: 1.85; -webkit-font-smoothing: antialiased;
  }
  body::before {
    content: ""; position: fixed; top: 0; left: 0; right: 0; bottom: 0;
    background-image: url("data:image/svg+xml,%3Csvg width='60' height='60' xmlns='http://www.w3.org/2000/svg'%3E%3Cpath d='M30 5 Q35 15 30 25 Q25 15 30 5Z' fill='%23d4c8b0' fill-opacity='0.04'/%3E%3C/svg%3E");
    background-size: 60px 60px; pointer-events: none; z-index: -1;
  }
  .wrapper { max-width: 860px; margin: 0 auto; padding: 3rem 2rem; }
  .back { display: inline-block; margin-bottom: 2rem; color: var(--accent); text-decoration: none; font-size: 0.85rem; }
  .back:hover { text-decoration: underline; }
  .doc-title { font-size: 1.8rem; font-weight: 700; color: var(--text); margin-bottom: 0.5rem; }
  .doc-sub { font-size: 1rem; color: var(--text-secondary); margin-bottom: 2rem; }
  .section { margin-bottom: 2.5rem; }
  .section h2 {
    font-size: 1.25rem; font-weight: 600; margin-bottom: 1rem; color: var(--accent);
    display: flex; align-items: center; gap: 0.6rem;
  }
  .section h2::before { content: ""; display: inline-block; width: 3px; height: 1.2em; background: var(--accent); border-radius: 2px; }
  .section h3 { font-size: 1.05rem; font-weight: 600; margin: 1.5rem 0 0.75rem; color: var(--text); }
  .section h4 { font-size: 0.95rem; font-weight: 600; margin: 1rem 0 0.5rem; color: var(--text-secondary); }
  .section p { color: var(--text-secondary); font-size: 0.92rem; margin-bottom: 0.75rem; line-height: 1.85; }
  table {
    width: 100%; border-collapse: collapse; margin: 1rem 0; font-size: 0.85rem;
    background: var(--card); border-radius: 8px; overflow: hidden; border: 1px solid var(--border);
  }
  th { background: var(--accent-subtle); padding: 0.6rem 0.75rem; text-align: left; font-weight: 600; color: var(--accent); font-size: 0.78rem; letter-spacing: 0.04em; }
  td { padding: 0.5rem 0.75rem; border-top: 1px solid var(--border); color: var(--text-secondary); }
  footer { text-align: center; color: var(--text-muted); font-size: 0.8rem; padding: 2rem; }
  @media (max-width: 768px) {
    .wrapper { padding: 2rem 1.25rem; }
    .doc-title { font-size: 1.4rem; }
  }
</style>"""

def parse_text(text):
    lines = text.strip().split('\n')
    sections = []
    current_section = {"title": "", "subsections": []}
    current_sub = {"title": "", "content": []}
    in_table = False
    table_lines = []
    
    for line in lines:
        line = line.strip()
        if not line:
            if current_sub["content"]:
                current_section["subsections"].append(current_sub)
                current_sub = {"title": "", "content": []}
            continue
        if line == '--- TABLES ---':
            if current_sub["content"]:
                current_section["subsections"].append(current_sub)
                current_sub = {"title": "", "content": []}
            in_table = True
            continue
        if in_table:
            table_lines.append(line)
            continue
        if line.startswith('### ') and not line.startswith('#### '):
            title = line[4:].strip()
            if current_section["title"]:
                if current_sub["content"]:
                    current_section["subsections"].append(current_sub)
                    current_sub = {"title": "", "content": []}
                sections.append(current_section)
            current_section = {"title": title, "subsections": [], "table": None}
            continue
        if re.match(r'^[一二三四五六七八九十]、', line) or re.match(r'^\d+\.\d+\s', line) or line.startswith('★'):
            if current_sub["content"]:
                current_section["subsections"].append(current_sub)
            current_sub = {"title": line, "content": []}
            continue
        current_sub["content"].append(line)
    
    if current_sub["content"]:
        current_section["subsections"].append(current_sub)
    if current_section["title"]:
        if table_lines:
            current_section["table"] = table_lines
        sections.append(current_section)
    return sections

def render_table(table_lines):
    if not table_lines:
        return ""
    rows = []
    for line in table_lines:
        cells = [c.strip() for c in line.split('|')]
        cells = [c for c in cells if c]
        rows.append(cells)
    if len(rows) < 2:
        return f"<pre>{chr(10).join(table_lines)}</pre>"
    h = '<div style="overflow-x:auto;"><table><thead><tr>'
    for cell in rows[0]:
        h += f'<th>{html.escape(cell)}</th>'
    h += '</tr></thead><tbody>'
    for row in rows[1:]:
        h += '<tr>'
        for cell in row:
            h += f'<td>{html.escape(cell)}</td>'
        h += '</tr>'
    h += '</tbody></table></div>'
    return h

def render_sections(sections):
    hp = []
    for sec in sections:
        t = html.escape(sec['title'])
        hp.append(f'<div class="section"><h2>{t}</h2>')
        for sub in sec['subsections']:
            st = sub['title']
            cl = sub['content']
            if st:
                if re.match(r'^[一二三四五六七八九十]、', st):
                    hp.append(f'<h3>{html.escape(st)}</h3>')
                elif re.match(r'^\d+\.\d+\s', st) or st.startswith('★'):
                    hp.append(f'<h4>{html.escape(st)}</h4>')
            if cl:
                para = []
                for line in cl:
                    if line.startswith(('●','•','-','·')):
                        if para:
                            hp.append(f'<p>{" ".join(html.escape(p) for p in para)}</p>')
                            para = []
                        hp.append(f'<p style="padding-left:1rem;">{html.escape(line)}</p>')
                    else:
                        para.append(line)
                if para:
                    hp.append(f'<p>{" ".join(html.escape(p) for p in para)}</p>')
        if sec.get('table'):
            hp.append(render_table(sec['table']))
        hp.append('</div>')
    return '\n'.join(hp)

for fname, title in files.items():
    txt_path = os.path.join(base, f"{fname}.txt")
    if not os.path.exists(txt_path):
        print(f"SKIP: {txt_path}")
        continue
    with open(txt_path, 'r', encoding='utf-8') as f:
        text = f.read()
    sections = parse_text(text)
    body = render_sections(sections)
    html_content = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{title} — Elevator — Don't Touch</title>
{CSS}
</head>
<body>
<div class="wrapper">
  <a href="index.html" class="back">← 返回作品集</a>
  <h1 class="doc-title">{title}</h1>
  <p class="doc-sub">Elevator — Don't Touch（别按那个键）· 第一人称 3D 密室互动恐怖游戏</p>
  {body}
</div>
<footer><p>© 2026 代立强 · Elevator — Don't Touch 项目文档</p></footer>
</body>
</html>"""
    out_path = os.path.join(base, f"{fname}.html")
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    print(f"Generated: {fname}.html")

print("Done!")
