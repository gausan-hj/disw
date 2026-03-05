import pandas as pd
import json
from datetime import datetime

# ===== 你要修改的地方 =====
SHEET_ID = "https://docs.google.com/spreadsheets/d/1YVa3nLUBW80j2nA4mudEqLH91RJ0FSRytmoDqmbyUJk/edit?usp=drive_link"  # 从链接中获取，比如：1AbcDefGhiJklMnoPqrStUvWxYz
SHEET_NAME = "Sheet3"  # 默认是Sheet3，如果不对就改
# ========================

# 生成Google Sheets的CSV导出链接
url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet={SHEET_NAME}"

# 读取数据
df = pd.read_csv(url)

# 打印前几行看看结构（调试用）
print("数据预览：")
print(df.head())
print("列名：")
print(df.columns.tolist())

# 手动处理数据
# 根据你给的Excel，数据结构是：
# 第0行：标题 "学长团个人分数板"
# 第1行：空 + 组名 "星穹组"
# 第2行开始：序号、学号、班级、中文名、英文名、性别、是否留宿、然后每天分数...

# 找到组别的分隔行
groups = []
current_group = None
start_row = 0

people = []

for i, row in df.iterrows():
    row_list = row.tolist()
    
    # 检查是否是组别标题行
    if pd.notna(row_list[1]) and isinstance(row_list[1], str) and "组" in row_list[1]:
        current_group = row_list[1]
        groups.append({"name": current_group, "start": i + 1})
        continue
    
    # 如果是人员数据行（有序号，且是数字）
    if pd.notna(row_list[0]) and isinstance(row_list[0], (int, float)):
        if current_group:
            # 提取基本信息
            person = {
                "group": current_group,
                "no": row_list[0],
                "student_id": row_list[1],
                "class": row_list[2],
                "name_cn": row_list[3],
                "name_en": row_list[4],
                "gender": row_list[5],
                "boarding": row_list[6],
                "daily_scores": []
            }
            
            # 提取每天分数（从第7列开始）
            for j in range(7, len(row_list)):
                if pd.notna(row_list[j]) and isinstance(row_list[j], (int, float)):
                    person["daily_scores"].append(row_list[j])
                else:
                    person["daily_scores"].append(0)
            
            # 计算总分
            person["total"] = sum(person["daily_scores"])
            
            people.append(person)

# 按组别整理
group_data = {}
for p in people:
    g = p["group"]
    if g not in group_data:
        group_data[g] = []
    group_data[g].append(p)

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
                daily_html += f'<span class="daily-score">D{i+1}:{int(score)}</span>'
        
        html += f"""
                    <tr class="{rank_class}" data-name="{p['name_cn']} {p['name_en']}">
                        <td>{rank}</td>
                        <td><strong>{p['name_cn']}</strong><br><small>{p['name_en']}</small></td>
                        <td>{p['class']}</td>
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
                // 找到对应的标签并点击
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
                // 如果搜索框为空，显示所有行
                allRows.forEach(row => {
                    row.style.display = '';
                    row.classList.remove('search-highlight');
                });
                return;
            }
            
            // 先隐藏所有行
            allRows.forEach(row => {
                row.style.display = 'none';
                row.classList.remove('search-highlight');
            });
            
            // 搜索匹配的行
            allRows.forEach(row => {
                const nameAttr = row.dataset.name.toLowerCase();
                const nameText = row.querySelector('td:nth-child(2)').innerText.toLowerCase();
                
                if (nameAttr.includes(searchTerm) || nameText.includes(searchTerm)) {
                    row.style.display = '';
                    row.classList.add('search-highlight');
                    
                    // 自动切换到该组
                    const groupSection = row.closest('.group-section');
                    const groupName = groupSection.dataset.group;
                    const targetTab = Array.from(tabs).find(tab => tab.dataset.group === groupName);
                    if (targetTab && !targetTab.classList.contains('active')) {
                        targetTab.click();
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

print("✅ HTML 生成成功！")
