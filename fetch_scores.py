import pandas as pd
from datetime import datetime

# ===== 你要修改的地方 =====
SHEET_ID = "1YVa3nLUBW80j2nA4mudEqLH91RJ0FSRytmoDqmbyUJk"
SHEET_NAME = "Sheet3"
# ========================

# 生成Google Sheets的CSV导出链接
url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet={SHEET_NAME}"

# 读取数据
print("正在从 Google Sheets 读取数据...")
df = pd.read_csv(url, header=None)
print(f"读取到 {len(df)} 行数据")

# ===== 直接把名字写死，绝对不会错 =====
members_list = [
    # ... 这里保持你原来的 members_list 不变 ...
]

# 从Google Sheets提取分数
people = []
for member in members_list:
    row_idx = member["row_index"]
    
    if row_idx < len(df):
        row = df.iloc[row_idx].tolist()
        
        # 计算总分：从第7列开始的所有数字加起来
        total_score = 0
        for j in range(7, len(row)):
            try:
                total_score += float(row[j])
            except (ValueError, TypeError):
                continue
        
        people.append({
            "group": member["group"],
            "name_cn": member["name_cn"],
            "name_en": member["name_en"],
            "class": member["class"],
            "student_id": member["student_id"],
            "total": total_score
        })
        print(f"✓ {member['name_cn']}: {total_score}分")
    else:
        print(f"✗ 找不到第{row_idx}行: {member['name_cn']}")

print(f"\n总共解析到 {len(people)} 位成员")

# 按组别整理
group_data = {}
group_totals = {}

for p in people:
    g = p["group"]
    if g not in group_data:
        group_data[g] = []
        group_totals[g] = 0
    group_data[g].append(p)
    group_totals[g] += p["total"]

print("\n✅ 组别总分：")
for g in ["星穹组", "夜曜组", "沧澜组"]:
    if g in group_totals:
        print(f"{g}: {int(group_totals[g])}分")

# 计算组排名
sorted_groups = sorted(group_totals.items(), key=lambda x: x[1], reverse=True)
group_rank = {g: i+1 for i, (g, _) in enumerate(sorted_groups)}

# 生成HTML
html = f"""<!DOCTYPE html>
<html lang="zh">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>训育处 - 学长团分数板</title>
<style>
/* 保留你原来的 CSS 样式 */
...
</style>
</head>
<body>
<div class="container">
<div class="header">
<h1>🏫 训育处 · 学长团分数板</h1>
<div class="update-time">最后更新：{datetime.now().strftime('%Y-%m-%d %H:%M')}</div>
<div class="search-box">
<input type="text" id="search" placeholder="🔍 输入姓名/班级/学号搜索...">
</div>
</div>

<div class="rank-bar">
"""

# 添加组排名
for g, total in sorted_groups:
    rank = group_rank[g]
    html += f"""
<div class="rank-item">
    <div class="rank-name">{g}</div>
    <div class="rank-score">{int(total)} <small>分</small></div>
    <div style="font-size:0.9rem; color:#7f8c8d;">第{rank}名</div>
</div>
"""

html += "</div><div class='group-container'>"

# 按顺序添加三个组
group_order = ["星穹组", "夜曜组", "沧澜组"]
for group_name in group_order:
    if group_name in group_data:
        members = group_data[group_name]
        
        html += f"""
<div class="group-card">
    <div class="group-header">
        <span class="group-title">{group_name} <small>第{group_rank[group_name]}名</small></span>
        <span class="group-score">{int(group_totals[group_name])}分</span>
    </div>
    <table>
        <thead>
            <tr>
                <th>序号</th>
                <th>姓名</th>
                <th>班级</th>
                <th>总分</th>
            </tr>
        </thead>
        <tbody>
"""
        for idx, p in enumerate(members, 1):
            name_en_display = p['name_en'][:25] + "..." if len(p['name_en']) > 25 else p['name_en']
            html += f"""
<tr data-name="{p['name_cn']} {p['name_en']} {p['class']} {p['student_id']}">
<td>{idx}</td>
<td><strong>{p['name_cn']}</strong><br><span style="font-size:0.75rem; color:#7f8c8d;">{name_en_display}</span></td>
<td>{p['class']}</td>
<td><span class="score-badge">{int(p['total'])}</span></td>
</tr>
"""
        html += "</tbody></table></div>"

html += """
</div>
<div class="footer">
训育处 · 数据每日自动更新 · 仅供参考
</div>
</div>

<script>
const searchInput = document.getElementById('search');
const allRows = document.querySelectorAll('tbody tr');

searchInput.addEventListener('input', (e) => {
    const searchTerm = e.target.value.toLowerCase().trim();
    allRows.forEach(row => {
        const searchText = row.getAttribute('data-name').toLowerCase();
        row.style.display = searchText.includes(searchTerm) ? '' : 'none';
    });
});
</script>
</body>
</html>
"""

# 保存HTML文件
with open("index.html", "w", encoding="utf-8") as f:
    f.write(html)

print(f"\n✅ 生成成功！共 {len(people)} 人")
