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

# ===== 组别位置 =====
group_positions = {
    2: "星穹组",   # 第3行
    14: "夜曜组",  # 第15行
    27: "沧澜组"   # 第28行
}

# 解析数据
people = []
current_group = None

print("开始解析数据...")

for i in range(len(df)):
    row = df.iloc[i].tolist()
    
    # 检查是否是组别行
    if i in group_positions:
        current_group = group_positions[i]
        print(f"第{i}行: 找到组别 {current_group}")
        continue
    
    # 检查是否是人员数据行（A列有序号）
    if len(row) > 0 and pd.notna(row[0]) and isinstance(row[0], (int, float)):
        # 跳过TOTAL行
        if len(row) > 4 and row[4] == "TOTAL":
            continue
            
        if current_group:
            # 提取基本信息
            name_cn = row[3] if len(row) > 3 and pd.notna(row[3]) else ""
            name_en = row[4] if len(row) > 4 and pd.notna(row[4]) else ""
            
            # 只添加有名字的人
            if name_cn or name_en:
                # 计算总分：从第7列开始的所有数字加起来
                total_score = 0
                for j in range(7, len(row)):
                    if pd.notna(row[j]) and isinstance(row[j], (int, float)):
                        total_score += row[j]
                
                people.append({
                    "group": current_group,
                    "name_cn": name_cn,
                    "name_en": name_en,
                    "class": row[2] if len(row) > 2 and pd.notna(row[2]) else "",
                    "student_id": row[1] if len(row) > 1 and pd.notna(row[1]) else "",
                    "total": total_score,
                    "original_order": i  # 保留原始顺序
                })

print(f"总共解析到 {len(people)} 位成员")

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

# 每个组内按原始顺序排序（绝对不能乱）
for g in group_data:
    group_data[g].sort(key=lambda x: x["original_order"])

print("\n✅ 解析结果：")
for g in group_data:
    print(f"{g}: {len(group_data[g])} 人, 组总分: {int(group_totals[g])}")
    for p in group_data[g]:
        print(f"  - {p['name_cn']}: {int(p['total'])}分")

# 计算组排名
sorted_groups = sorted(group_totals.items(), key=lambda x: x[1], reverse=True)
group_rank = {}
for i, (g, _) in enumerate(sorted_groups, 1):
    group_rank[g] = i

# 生成HTML
html = f"""<!DOCTYPE html>
<html lang="zh">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>训育处 - 学长团分数板</title>
    <style>
        * {{
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }}
        body {{
            font-family: 'Segoe UI', 'Microsoft YaHei', sans-serif;
            background: #f5f5f5;
            padding: 24px;
            color: #2c3e50;
        }}
        .container {{
            max-width: 1200px;
            margin: 0 auto;
        }}
        .header {{
            background: white;
            border-radius: 12px;
            padding: 24px;
            margin-bottom: 24px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        }}
        h1 {{
            font-size: 1.8rem;
            font-weight: 500;
            margin-bottom: 8px;
        }}
        .update-time {{
            color: #95a5a6;
            font-size: 0.9rem;
            margin-top: 8px;
        }}
        .search-box {{
            margin-top: 16px;
        }}
        #search {{
            width: 100%;
            max-width: 400px;
            padding: 10px 16px;
            border: 1px solid #dce1e5;
            border-radius: 30px;
            font-size: 0.95rem;
        }}
        .rank-bar {{
            display: flex;
            gap: 16px;
            margin-bottom: 24px;
            flex-wrap: wrap;
        }}
        .rank-item {{
            flex: 1;
            min-width: 200px;
            background: white;
            border-radius: 10px;
            padding: 16px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.05);
            border-left: 6px solid;
        }}
        .rank-1 {{ border-left-color: #b0bec5; }}
        .rank-2 {{ border-left-color: #b0bec5; }}
        .rank-3 {{ border-left-color: #b0bec5; }}
        .rank-name {{
            font-size: 1.2rem;
            font-weight: 500;
            margin-bottom: 8px;
        }}
        .rank-score {{
            font-size: 1.8rem;
            font-weight: 500;
        }}
        .rank-score small {{
            font-size: 1rem;
            font-weight: 400;
            color: #7f8c8d;
        }}
        .group-container {{
            display: flex;
            gap: 24px;
            flex-wrap: wrap;
            margin-top: 24px;
        }}
        .group-card {{
            flex: 1;
            min-width: 300px;
            background: white;
            border-radius: 12px;
            padding: 20px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        }}
        .group-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 16px;
            padding-bottom: 12px;
            border-bottom: 2px solid #ecf0f1;
        }}
        .group-title {{
            font-size: 1.4rem;
            font-weight: 500;
        }}
        .group-title small {{
            font-size: 0.9rem;
            color: #95a5a6;
            margin-left: 8px;
        }}
        .group-score {{
            font-size: 1.2rem;
            font-weight: 500;
            background: #f8fafc;
            padding: 4px 12px;
            border-radius: 20px;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
        }}
        th {{
            text-align: left;
            padding: 10px 8px;
            background: #f8fafc;
            font-weight: 500;
            color: #4a5b6b;
            font-size: 0.9rem;
        }}
        td {{
            padding: 10px 8px;
            border-bottom: 1px solid #edf2f7;
        }}
        .score-badge {{
            background: #ecf0f1;
            padding: 4px 10px;
            border-radius: 20px;
            font-weight: 500;
            display: inline-block;
        }}
        .footer {{
            margin-top: 32px;
            text-align: center;
            color: #95a5a6;
            font-size: 0.85rem;
            padding: 16px;
        }}
        @media (max-width: 768px) {{
            .group-container {{
                flex-direction: column;
            }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🏫 训育处 · 学长团分数板</h1>
            <div class="update-time">最后更新：{datetime.now().strftime('%Y-%m-%d %H:%M')}</div>
            <div class="search-box">
                <input type="text" id="search" placeholder="🔍 输入姓名或学号搜索...">
            </div>
        </div>

        <div class="rank-bar">
"""

# 添加组排名
for g, total in sorted_groups:
    rank = group_rank[g]
    html += f"""
            <div class="rank-item rank-{rank}">
                <div class="rank-name">{g}</div>
                <div class="rank-score">{int(total)} <small>分</small></div>
                <div style="font-size:0.9rem; color:#7f8c8d;">第{rank}名</div>
            </div>
"""

html += """
        </div>

        <div class="group-container">
"""

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
        
        # 严格按照原始顺序显示
        for idx, p in enumerate(members, 1):
            # 处理英文名
            name_en_display = p['name_en'][:15] + "..." if len(p['name_en']) > 15 else p['name_en']
            
            html += f"""
                        <tr data-name="{p['name_cn']} {p['name_en']}">
                            <td>{idx}</td>
                            <td>
                                <strong>{p['name_cn']}</strong><br>
                                <span style="font-size:0.75rem; color:#7f8c8d;">{name_en_display}</span>
                            </td>
                            <td>{p['class']}</td>
                            <td><span class="score-badge">{int(p['total'])}</span></td>
                        </tr>
            """
        
        html += """
                    </tbody>
                </table>
            </div>
        """

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
            
            if (searchTerm === '') {
                allRows.forEach(row => row.style.display = '');
                return;
            }
            
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
for g in group_order:
    if g in group_data:
        print(f"{g}: {len(group_data[g])} 人, 总分: {int(group_totals[g])}, 第{group_rank[g]}名")
