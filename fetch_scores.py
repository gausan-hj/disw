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

# 按组别整理，保持原始顺序
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
    # 打印所有人确认
    for p in group_data[g]:
        print(f"  - {p['name_cn']}: {int(p['total'])}分")

# 按总分排序组别（用于顶部排名）
sorted_groups = sorted(group_totals.items(), key=lambda x: x[1], reverse=True)

# 生成HTML - 保持你之前的格式
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
            border: 1px solid #eaeef2;
        }}
        h1 {{
            font-size: 1.8rem;
            font-weight: 500;
            color: #2c3e50;
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
        #search:focus {{
            outline: none;
            border-color: #9aa6b2;
        }}
        .group-rank-header {{
            background: white;
            border-radius: 12px;
            padding: 20px;
            margin-bottom: 24px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.05);
            border: 1px solid #eaeef2;
        }}
        .group-rank-title {{
            font-size: 1.2rem;
            font-weight: 500;
            color: #34495e;
            margin-bottom: 16px;
            padding-bottom: 12px;
            border-bottom: 1px solid #ecf0f1;
        }}
        .group-rank-list {{
            display: flex;
            gap: 16px;
            flex-wrap: wrap;
        }}
        .group-rank-item {{
            flex: 1;
            min-width: 180px;
            background: #f8fafc;
            border-radius: 10px;
            padding: 16px;
            border-left: 4px solid;
            transition: all 0.2s;
            border: 1px solid #e9ecef;
        }}
        .group-rank-name {{
            font-size: 1.1rem;
            font-weight: 500;
            margin-bottom: 8px;
            color: #2c3e50;
        }}
        .group-rank-score {{
            font-size: 1.6rem;
            font-weight: 500;
            color: #34495e;
        }}
        .group-rank-score small {{
            font-size: 0.9rem;
            font-weight: 400;
            color: #7f8c8d;
        }}
        .tabs {{
            display: flex;
            gap: 8px;
            margin-bottom: 20px;
            flex-wrap: wrap;
        }}
        .tab {{
            padding: 10px 24px;
            background: white;
            border: 1px solid #dce1e5;
            border-radius: 30px;
            font-size: 0.95rem;
            font-weight: 500;
            color: #5d6d7e;
            cursor: pointer;
            transition: all 0.2s;
        }}
        .tab:hover {{
            background: #f1f4f8;
        }}
        .tab.active {{
            background: #2c3e50;
            color: white;
            border-color: #2c3e50;
        }}
        .group-section {{
            display: none;
            background: white;
            border-radius: 12px;
            padding: 24px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.05);
            border: 1px solid #eaeef2;
        }}
        .group-section.active {{
            display: block;
        }}
        .group-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
            padding-bottom: 16px;
            border-bottom: 1px solid #ecf0f1;
        }}
        .group-name {{
            font-size: 1.4rem;
            font-weight: 500;
            color: #2c3e50;
        }}
        .group-total {{
            font-size: 1.2rem;
            font-weight: 500;
            color: #5d6d7e;
            background: #f8fafc;
            padding: 6px 16px;
            border-radius: 30px;
            border: 1px solid #e9ecef;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
        }}
        th {{
            text-align: left;
            padding: 14px 12px;
            background: #f8fafc;
            font-weight: 500;
            color: #4a5b6b;
            border-bottom: 2px solid #e2e8f0;
        }}
        td {{
            padding: 14px 12px;
            border-bottom: 1px solid #edf2f7;
        }}
        tr:hover {{
            background: #fafcfd;
        }}
        .score-badge {{
            background: #ecf0f1;
            padding: 6px 14px;
            border-radius: 30px;
            font-weight: 500;
            display: inline-block;
            font-size: 1rem;
        }}
        .footer {{
            margin-top: 32px;
            text-align: center;
            color: #95a5a6;
            font-size: 0.85rem;
            padding: 16px;
            border-top: 1px solid #e9ecef;
        }}
        
        /* 组别颜色 - 素色 */
        .border-star {{ border-left-color: #9aa6b2; }}
        .border-night {{ border-left-color: #8a9aa8; }}
        .border-ocean {{ border-left-color: #7f8fa6; }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🏫 训育处 · 学长团个人分数板</h1>
            <div class="update-time">最后更新：{datetime.now().strftime('%Y-%m-%d %H:%M')}</div>
            <div class="search-box">
                <input type="text" id="search" placeholder="🔍 输入姓名或学号搜索...">
            </div>
        </div>

        <div class="group-rank-header">
            <div class="group-rank-title">📊 组别总分排名</div>
            <div class="group-rank-list" id="groupRankList">
"""

# 添加组别排名卡片
group_border = {
    "星穹组": "border-star",
    "夜曜组": "border-night",
    "沧澜组": "border-ocean"
}

for i, (group_name, total) in enumerate(sorted_groups):
    border_class = group_border.get(group_name, "border-star")
    rank_text = "第1名" if i == 0 else "第2名" if i == 1 else "第3名"
    html += f"""
                <div class="group-rank-item {border_class}" data-group="{group_name}">
                    <div class="group-rank-name">{group_name}</div>
                    <div class="group-rank-score">{int(total)} <small>分</small></div>
                    <div style="font-size:0.8rem; color:#95a5a6;">{rank_text}</div>
                </div>
    """

html += """
            </div>
        </div>

        <div class="tabs" id="tabs">
"""

# 添加标签页
for group_name, _ in sorted_groups:
    active_class = "active" if group_name == sorted_groups[0][0] else ""
    html += f"""
            <button class="tab {active_class}" data-group="{group_name}">{group_name}</button>
"""

html += """
        </div>
"""

# 添加每个组别的表格 - 严格按照原始顺序！
for group_name, members in group_data.items():
    active_class = "active" if group_name == sorted_groups[0][0] else ""
    
    html += f"""
        <div class="group-section {active_class}" data-group="{group_name}">
            <div class="group-header">
                <span class="group-name">{group_name}</span>
                <span class="group-total">总分 {int(group_totals[group_name])}</span>
            </div>
            <table>
                <thead>
                    <tr>
                        <th style="width: 60px">序号</th>
                        <th>姓名</th>
                        <th style="width: 100px">班级</th>
                        <th style="width: 120px">学号</th>
                        <th style="width: 100px">总分</th>
                    </tr>
                </thead>
                <tbody>
    """
    
    # 严格按照原始顺序显示（不排序！）
    for idx, p in enumerate(members, 1):
        # 处理学号显示
        student_id_display = ""
        if isinstance(p['student_id'], (int, float)):
            student_id_display = str(int(p['student_id']))
        else:
            student_id_display = str(p['student_id'])
        
        # 处理英文名过长
        name_en_display = p['name_en'][:25] + "..." if len(p['name_en']) > 25 else p['name_en']
        
        html += f"""
                    <tr data-name="{p['name_cn']} {p['name_en']} {student_id_display}">
                        <td>{idx}</td>
                        <td>
                            <strong>{p['name_cn']}</strong><br>
                            <span style="font-size:0.8rem; color:#7f8c8d;">{name_en_display}</span>
                        </td>
                        <td>{p['class']}</td>
                        <td>{student_id_display}</td>
                        <td><span class="score-badge">{int(p['total'])}</span></td>
                    </tr>
        """
    
    html += """
                </tbody>
            </table>
        </div>
    """

html += """
        <div class="footer">
            训育处 · 数据每日自动更新 · 仅供参考
        </div>
    </div>

    <script>
        const tabs = document.querySelectorAll('.tab');
        const sections = document.querySelectorAll('.group-section');
        const groupRankItems = document.querySelectorAll('.group-rank-item');
        
        tabs.forEach(tab => {
            tab.addEventListener('click', () => {
                const groupName = tab.dataset.group;
                tabs.forEach(t => t.classList.remove('active'));
                tab.classList.add('active');
                sections.forEach(section => {
                    section.classList.remove('active');
                    if (section.dataset.group === groupName) {
                        section.classList.add('active');
                    }
                });
            });
        });

        groupRankItems.forEach(item => {
            item.addEventListener('click', () => {
                const groupName = item.dataset.group;
                const targetTab = Array.from(tabs).find(tab => tab.dataset.group === groupName);
                if (targetTab) targetTab.click();
            });
        });

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
print("组别统计：")
for g in ["星穹组", "夜曜组", "沧澜组"]:
    if g in group_data:
        print(f"  {g}: {len(group_data[g])} 人, 总分: {int(group_totals[g])}")
        # 打印所有人确认
        for p in group_data[g]:
            print(f"    {p['name_cn']}: {int(p['total'])}分")
