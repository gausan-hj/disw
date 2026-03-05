import pandas as pd
from datetime import datetime

# ===== 你要修改的地方 =====
SHEET_ID = "1YVa3nLUBW80j2nA4mudEqLH91RJ0FSRytmoDqmbyUJk"
SHEET_NAME = "Sheet3"
# ========================

# 生成Google Sheets的CSV导出链接
url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet={SHEET_NAME}"

# 读取数据，不把第一行当列名
print("正在从 Google Sheets 读取数据...")
df = pd.read_csv(url, header=None)

print(f"读取到 {len(df)} 行数据")

# 解析数据
people = []
current_group = None
dates = []  # 存储日期

# 先读取第一行获取日期
if len(df) > 0:
    first_row = df.iloc[0].tolist()
    # 从第7列（索引7）开始是日期
    for i in range(7, len(first_row)):
        if pd.notna(first_row[i]):
            dates.append(first_row[i])
        else:
            dates.append(f"Day{i-6}")

print(f"找到 {len(dates)} 个日期")

# 遍历每一行
for i in range(len(df)):
    row = df.iloc[i].tolist()
    
    # 检查是否是组别标题行（B列有"星穹组"、"夜曜组"等）
    if len(row) > 1 and pd.notna(row[1]) and isinstance(row[1], str) and "组" in row[1]:
        current_group = row[1]
        print(f"第{i}行: 找到组别 {current_group}")
        continue
    
    # 检查是否是人员数据行（A列有序号，且是数字）
    if len(row) > 0 and pd.notna(row[0]) and isinstance(row[0], (int, float)):
        if current_group:
            # 提取基本信息
            # A列: 序号
            # B列: 学号
            # C列: 班级
            # D列: 中文名
            # E列: 英文名
            # F列: 性别
            # G列: 是否留宿
            # H列开始: 每日分数
            
            person = {
                "group": current_group,
                "no": row[0],
                "student_id": row[1] if len(row) > 1 else "",
                "class": row[2] if len(row) > 2 else "",
                "name_cn": row[3] if len(row) > 3 else "",
                "name_en": row[4] if len(row) > 4 else "",
                "gender": row[5] if len(row) > 5 else "",
                "boarding": row[6] if len(row) > 6 else "",
                "daily_scores": [],
                "dates": dates
            }
            
            # 提取每天分数（从第8列开始，索引7）
            for j in range(7, len(row)):
                if j < 7 + len(dates):  # 只取有日期的列
                    if pd.notna(row[j]) and isinstance(row[j], (int, float)):
                        person["daily_scores"].append(row[j])
                    else:
                        person["daily_scores"].append(0)
            
            # 计算总分
            person["total"] = sum(person["daily_scores"])
            
            people.append(person)

print(f"总共解析到 {len(people)} 位成员")

# 按组别整理
group_data = {}
for p in people:
    g = p["group"]
    if g not in group_data:
        group_data[g] = []
    group_data[g].append(p)

print(f"组别分布: {list(group_data.keys())}")

# 如果没有解析到任何人，输出错误信息
if len(people) == 0:
    print("❌ 错误：没有解析到任何成员！")
    print("请检查：")
    print("1. Google Sheets 的权限是否设置为“任何知道链接的人可查看”")
    print("2. Sheet 名字是否正确（当前是 Sheet3）")
    print("3. 数据格式是否和 Excel 一致")
    exit(1)

# 计算各组总分
group_totals = {}
for g, members in group_data.items():
    group_totals[g] = sum([m["total"] for m in members])

# 按总分排序组别
sorted_groups = sorted(group_totals.items(), key=lambda x: x[1], reverse=True)

# 生成HTML
html = f"""
<!DOCTYPE html>
<html lang="zh">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>学长团个人分数板</title>
    <style>
        * {{
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }}
        body {{
            font-family: system-ui, -apple-system, 'Segoe UI', Roboto, 'Helvetica Neue', sans-serif;
            background: #f3f4f6;
            padding: 20px;
            min-height: 100vh;
        }}
        .container {{
            max-width: 1200px;
            margin: 0 auto;
        }}
        h1 {{
            font-size: 2rem;
            font-weight: 700;
            color: #1f2937;
            margin-bottom: 8px;
        }}
        .update-time {{
            color: #6b7280;
            margin-bottom: 24px;
            font-size: 0.9rem;
        }}
        .search-box {{
            margin-bottom: 24px;
        }}
        #search {{
            width: 100%;
            padding: 12px 16px;
            font-size: 1rem;
            border: 1px solid #e5e7eb;
            border-radius: 12px;
            box-shadow: 0 1px 2px rgba(0,0,0,0.05);
            transition: all 0.2s;
        }}
        #search:focus {{
            outline: none;
            border-color: #3b82f6;
            box-shadow: 0 0 0 3px rgba(59,130,246,0.2);
        }}
        .group-rank-header {{
            background: white;
            border-radius: 16px;
            padding: 20px;
            margin-bottom: 24px;
            box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);
        }}
        .group-rank-title {{
            font-size: 1.25rem;
            font-weight: 600;
            color: #374151;
            margin-bottom: 16px;
        }}
        .group-rank-list {{
            display: flex;
            gap: 16px;
            flex-wrap: wrap;
        }}
        .group-rank-item {{
            flex: 1;
            min-width: 200px;
            background: #f9fafb;
            border-radius: 12px;
            padding: 16px;
            border-left: 6px solid;
            transition: transform 0.2s;
            cursor: pointer;
        }}
        .group-rank-item:hover {{
            transform: translateY(-2px);
            box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);
        }}
        .group-rank-name {{
            font-size: 1.1rem;
            font-weight: 600;
            margin-bottom: 8px;
        }}
        .group-rank-score {{
            font-size: 1.5rem;
            font-weight: 700;
        }}
        .group-rank-score small {{
            font-size: 0.9rem;
            font-weight: 400;
            color: #6b7280;
        }}
        .tabs {{
            display: flex;
            gap: 8px;
            margin-bottom: 24px;
            flex-wrap: wrap;
        }}
        .tab {{
            padding: 10px 20px;
            background: white;
            border: 1px solid #e5e7eb;
            border-radius: 30px;
            font-size: 1rem;
            font-weight: 500;
            color: #4b5563;
            cursor: pointer;
            transition: all 0.2s;
        }}
        .tab:hover {{
            background: #f3f4f6;
        }}
        .tab.active {{
            background: #1f2937;
            color: white;
            border-color: #1f2937;
        }}
        .group-section {{
            display: none;
            background: white;
            border-radius: 20px;
            padding: 20px;
            box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);
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
            border-bottom: 2px solid #e5e7eb;
        }}
        .group-name {{
            font-size: 1.5rem;
            font-weight: 700;
        }}
        .group-total {{
            font-size: 1.25rem;
            font-weight: 600;
            color: #4b5563;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
        }}
        th {{
            text-align: left;
            padding: 12px 8px;
            background: #f9fafb;
            font-weight: 600;
            color: #374151;
            border-radius: 8px 8px 0 0;
        }}
        td {{
            padding: 12px 8px;
            border-bottom: 1px solid #e5e7eb;
        }}
        .rank-1 {{
            background: rgba(255,215,0,0.1);
            font-weight: 600;
        }}
        .rank-1 td:first-child::before {{
            content: "👑 ";
        }}
        .score-badge {{
            background: #e5e7eb;
            padding: 4px 8px;
            border-radius: 20px;
            font-size: 0.85rem;
            font-weight: 600;
        }}
        .daily-scores {{
            display: flex;
            gap: 4px;
            flex-wrap: wrap;
            margin-top: 4px;
        }}
        .daily-score {{
            background: #f3f4f6;
            padding: 2px 6px;
            border-radius: 12px;
            font-size: 0.75rem;
            color: #4b5563;
        }}
        .search-highlight {{
            background: #fef3c7;
        }}
        .color-star {{
            border-left-color: #3b82f6;
        }}
        .color-night {{
            border-left-color: #8b5cf6;
        }}
        .color-ocean {{
            border-left-color: #10b981;
        }}
        .bg-star {{ background: #3b82f6; color: white; }}
        .bg-night {{ background: #8b5cf6; color: white; }}
        .bg-ocean {{ background: #10b981; color: white; }}
        .text-star {{ color: #3b82f6; }}
        .text-night {{ color: #8b5cf6; }}
        .text-ocean {{ color: #10b981; }}
        @media (max-width: 640px) {{
            body {{ padding: 12px; }}
            .group-rank-list {{ flex-direction: column; }}
            th, td {{ padding: 8px 4px; font-size: 0.9rem; }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1>🏆 学长团个人分数板</h1>
        <div class="update-time">最后更新：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</div>
        
        <div class="search-box">
            <input type="text" id="search" placeholder="🔍 输入姓名搜索...">
        </div>

        <div class="group-rank-header">
            <div class="group-rank-title">📊 组别总分排名</div>
            <div class="group-rank-list" id="groupRankList">
"""

# 添加组别排名卡片
group_colors = {
    "星穹组": "color-star",
    "夜曜组": "color-night", 
    "沧澜组": "color-ocean"
}

group_color_classes = {
    "星穹组": "star",
    "夜曜组": "night",
    "沧澜组": "ocean"
}

for i, (group_name, total) in enumerate(sorted_groups):
    color_class = group_colors.get(group_name, "color-star")
    badge_class = group_color_classes.get(group_name, "star")
    rank_emoji = "🥇" if i == 0 else "🥈" if i == 1 else "🥉" if i == 2 else f"{i+1}."
    html += f"""
                <div class="group-rank-item {color_class}" data-group="{group_name}">
                    <div class="group-rank-name">{rank_emoji} {group_name}</div>
                    <div class="group-rank-score">{int(total)} <small>分</small></div>
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

# 添加每个组别的表格
for group_name, members in group_data.items():
    # 按总分排序成员
    sorted_members = sorted(members, key=lambda x: x["total"], reverse=True)
    active_class = "active" if group_name == sorted_groups[0][0] else ""
    color_class = group_color_classes.get(group_name, "star")
    
    html += f"""
        <div class="group-section {active_class}" id="group-{group_name}" data-group="{group_name}">
            <div class="group-header">
                <span class="group-name">{group_name}</span>
                <span class="group-total">总分：{int(group_totals[group_name])}</span>
            </div>
            <table>
                <thead>
                    <tr>
                        <th>排名</th>
                        <th>姓名</th>
                        <th>班级</th>
                        <th>学号</th>
                        <th>总分</th>
                        <th>每日得分</th>
                    </tr>
                </thead>
                <tbody>
    """
    
    for rank, p in enumerate(sorted_members, 1):
        rank_class = "rank-1" if rank == 1 else ""
        # 生成每日得分小标签
        daily_html = ""
        for i, score in enumerate(p["daily_scores"]):
            if score > 0:
                date_str = p["dates"][i] if i < len(p["dates"]) else f"D{i+1}"
                # 格式化日期，只显示月-日
                if isinstance(date_str, str) and "00:00" in date_str:
                    date_str = date_str[5:10]  # 取 "03-02" 这样的格式
                daily_html += f'<span class="daily-score" title="{date_str}">{int(score)}</span>'
        
        html += f"""
                    <tr class="{rank_class}" data-name="{p['name_cn']} {p['name_en']}">
                        <td>{rank}</td>
                        <td><strong>{p['name_cn']}</strong><br><small>{p['name_en']}</small></td>
                        <td>{p['class']}</td>
                        <td>{int(p['student_id']) if isinstance(p['student_id'], (int, float)) else p['student_id']}</td>
                        <td><span class="score-badge bg-{color_class}">{int(p['total'])}</span></td>
                        <td><div class="daily-scores">{daily_html}</div></td>
                    </tr>
        """
    
    html += """
                </tbody>
            </table>
        </div>
    """

# 添加搜索脚本
html += """
    </div>

    <script>
        // 标签页切换
        const tabs = document.querySelectorAll('.tab');
        const sections = document.querySelectorAll('.group-section');
        const groupRankItems = document.querySelectorAll('.group-rank-item');
        
        tabs.forEach(tab => {
            tab.addEventListener('click', () => {
                const groupName = tab.dataset.group;
                
                // 更新标签页激活状态
                tabs.forEach(t => t.classList.remove('active'));
                tab.classList.add('active');
                
                // 更新显示区域
                sections.forEach(section => {
                    section.classList.remove('active');
                    if (section.dataset.group === groupName) {
                        section.classList.add('active');
                    }
                });
            });
        });

        // 点击组别排名卡片切换到对应组
        groupRankItems.forEach(item => {
            item.addEventListener('click', () => {
                const groupName = item.dataset.group;
                const targetTab = Array.from(tabs).find(tab => tab.dataset.group === groupName);
                if (targetTab) {
                    targetTab.click();
                }
            });
        });

        // 搜索功能
        const searchInput = document.getElementById('search');
        const allRows = document.querySelectorAll('tbody tr');
        
        searchInput.addEventListener('input', (e) => {
            const searchTerm = e.target.value.toLowerCase().trim();
            
            if (searchTerm === '') {
                allRows.forEach(row => {
                    row.style.display = '';
                    row.classList.remove('search-highlight');
                });
                return;
            }
            
            allRows.forEach(row => {
                row.style.display = 'none';
                row.classList.remove('search-highlight');
            });
            
            allRows.forEach(row => {
                const nameAttr = row.dataset.name ? row.dataset.name.toLowerCase() : '';
                const nameCell = row.querySelector('td:nth-child(2)');
                const nameText = nameCell ? nameCell.innerText.toLowerCase() : '';
                
                if (nameAttr.includes(searchTerm) || nameText.includes(searchTerm)) {
                    row.style.display = '';
                    row.classList.add('search-highlight');
                    
                    // 自动切换到该组
                    const groupSection = row.closest('.group-section');
                    if (groupSection) {
                        const groupName = groupSection.dataset.group;
                        const targetTab = Array.from(tabs).find(tab => tab.dataset.group === groupName);
                        if (targetTab && !targetTab.classList.contains('active')) {
                            targetTab.click();
                        }
                    }
                }
            });
        });
    </script>
</body>
</html>
"""

# 保存HTML文件
with open("index.html", "w", encoding="utf-8") as f:
    f.write(html)

print(f"✅ HTML 生成成功！共 {len(people)} 位成员")
print(f"组别: {list(group_data.keys())}")
for g, members in group_data.items():
    print(f"  {g}: {len(members)} 人")
