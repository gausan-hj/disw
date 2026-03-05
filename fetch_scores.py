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

# ===== 根据你提供的信息，组别分别在以下行（0-based索引）=====
# 第2行（索引2）：星穹组
# 第14行（索引14）：夜曜组
# 第27行（索引27）：沧澜组
# ====================================================

group_positions = {
    2: "星穹组",
    14: "夜曜组", 
    27: "沧澜组"
}

current_group = None
current_group_start_row = 0

# 遍历每一行
for i in range(len(df)):
    row = df.iloc[i].tolist()
    
    # 检查这一行是不是组别标题行
    if i in group_positions:
        current_group = group_positions[i]
        current_group_start_row = i
        print(f"第{i}行: 找到组别 {current_group}")
        continue
    
    # 检查是否是人员数据行（A列有序号，且是数字）
    if len(row) > 0 and pd.notna(row[0]) and isinstance(row[0], (int, float)):
        if current_group:
            # 提取基本信息
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
            
            # 打印找到的人员（调试用）
            if len(people) < 5:  # 只打印前几个
                print(f"  找到人员: {person['name_cn']} ({person['group']})")

print(f"总共解析到 {len(people)} 位成员")

# 按组别整理
group_data = {}
for p in people:
    g = p["group"]
    if g not in group_data:
        group_data[g] = []
    group_data[g].append(p)

print(f"组别分布: {list(group_data.keys())}")
for g, members in group_data.items():
    print(f"  {g}: {len(members)} 人")

# 如果没有解析到任何人，输出错误信息
if len(people) == 0:
    print("❌ 错误：没有解析到任何成员！")
    print("请检查：")
    print("1. Google Sheets 的权限是否设置为“任何知道链接的人可查看”")
    print("2. 组别行号是否正确（当前设置：2=星穹组, 14=夜曜组, 27=沧澜组）")
    exit(1)

# 计算各组总分
group_totals = {}
for g, members in group_data.items():
    group_totals[g] = sum([m["total"] for m in members])

# 按总分排序组别
sorted_groups = sorted(group_totals.items(), key=lambda x: x[1], reverse=True)

# 生成HTML - 素色背景风格
html = f"""
<!DOCTYPE html>
<html lang="zh">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>训育处 - 学长团个人分数板</title>
    <style>
        * {{
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }}
        body {{
            font-family: 'Segoe UI', 'Microsoft YaHei', system-ui, -apple-system, sans-serif;
            background: #f5f5f5;
            padding: 24px;
            min-height: 100vh;
            color: #2c3e50;
        }}
        .container {{
            max-width: 1400px;
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
            letter-spacing: 1px;
        }}
        .subtitle {{
            color: #7f8c8d;
            font-size: 0.95rem;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}
        .update-time {{
            color: #95a5a6;
            font-size: 0.9rem;
        }}
        .search-box {{
            margin-top: 16px;
        }}
        #search {{
            width: 100%;
            max-width: 400px;
            padding: 10px 16px;
            font-size: 0.95rem;
            border: 1px solid #dce1e5;
            border-radius: 30px;
            background: white;
            transition: all 0.2s;
        }}
        #search:focus {{
            outline: none;
            border-color: #9aa6b2;
            box-shadow: 0 0 0 2px rgba(44,62,80,0.1);
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
            cursor: pointer;
            border: 1px solid #e9ecef;
            box-shadow: 0 1px 3px rgba(0,0,0,0.02);
        }}
        .group-rank-item:hover {{
            background: #f1f4f8;
            transform: translateY(-1px);
            box-shadow: 0 4px 8px rgba(0,0,0,0.05);
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
            box-shadow: 0 1px 3px rgba(0,0,0,0.02);
        }}
        .tab:hover {{
            background: #f1f4f8;
            border-color: #b8c5d0;
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
            font-size: 0.95rem;
        }}
        th {{
            text-align: left;
            padding: 14px 12px;
            background: #f8fafc;
            font-weight: 500;
            color: #4a5b6b;
            border-bottom: 1px solid #e2e8f0;
            font-size: 0.9rem;
            letter-spacing: 0.3px;
        }}
        td {{
            padding: 14px 12px;
            border-bottom: 1px solid #edf2f7;
            color: #3a4a5a;
        }}
        tr:hover {{
            background: #fafcfd;
        }}
        .rank-1 {{
            background: #fcfaf7;
        }}
        .rank-1 td:first-child {{
            font-weight: 600;
            color: #b85c5c;
        }}
        .rank-1 td:first-child::before {{
            content: "👑 ";
            font-size: 0.9rem;
        }}
        .score-badge {{
            background: #ecf0f1;
            padding: 4px 10px;
            border-radius: 30px;
            font-size: 0.85rem;
            font-weight: 500;
            color: #2c3e50;
            display: inline-block;
        }}
        .daily-scores {{
            display: flex;
            gap: 6px;
            flex-wrap: wrap;
        }}
        .daily-score {{
            background: #f1f5f9;
            padding: 3px 8px;
            border-radius: 16px;
            font-size: 0.8rem;
            color: #4a5b6b;
            border: 1px solid #e2e8f0;
            min-width: 32px;
            text-align: center;
        }}
        .search-highlight {{
            background: #fff9e6;
        }}
        .border-star {{ border-left-color: #9aa6b2; }}
        .border-night {{ border-left-color: #7f8c8d; }}
        .border-ocean {{ border-left-color: #95a5a6; }}
        
        /* 素色风格 - 只有极淡的区分 */
        .group-rank-item.border-star {{ border-left-color: #b0bec5; }}
        .group-rank-item.border-night {{ border-left-color: #90a4ae; }}
        .group-rank-item.border-ocean {{ border-left-color: #aab7b8; }}
        
        .footer {{
            margin-top: 32px;
            text-align: center;
            color: #95a5a6;
            font-size: 0.85rem;
            padding: 16px;
            border-top: 1px solid #e9ecef;
        }}
        @media (max-width: 640px) {{
            body {{ padding: 12px; }}
            .group-rank-list {{ flex-direction: column; }}
            th, td {{ padding: 10px 6px; font-size: 0.85rem; }}
            .daily-score {{ padding: 2px 4px; font-size: 0.7rem; }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🏫 训育处 · 学长团个人分数板</h1>
            <div class="subtitle">
                <span>Prefects' Personal Scoreboard</span>
                <span class="update-time">更新：{datetime.now().strftime('%Y-%m-%d %H:%M')}</span>
            </div>
            <div class="search-box">
                <input type="text" id="search" placeholder="🔍 输入姓名或学号搜索...">
            </div>
        </div>

        <div class="group-rank-header">
            <div class="group-rank-title">📋 组别总分排名</div>
            <div class="group-rank-list" id="groupRankList">
"""

# 添加组别排名卡片 - 素色风格
group_border_classes = {
    "星穹组": "border-star",
    "夜曜组": "border-night", 
    "沧澜组": "border-ocean"
}

for i, (group_name, total) in enumerate(sorted_groups):
    border_class = group_border_classes.get(group_name, "border-star")
    rank_display = "第1名" if i == 0 else "第2名" if i == 1 else "第3名" if i == 2 else f"第{i+1}名"
    html += f"""
                <div class="group-rank-item {border_class}" data-group="{group_name}">
                    <div class="group-rank-name">{group_name}</div>
                    <div class="group-rank-score">{int(total)} <small>分</small></div>
                    <div style="font-size:0.8rem; color:#95a5a6; margin-top:4px;">{rank_display}</div>
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
    
    html += f"""
        <div class="group-section {active_class}" id="group-{group_name}" data-group="{group_name}">
            <div class="group-header">
                <span class="group-name">{group_name}</span>
                <span class="group-total">总分 {int(group_totals[group_name])}</span>
            </div>
            <table>
                <thead>
                    <tr>
                        <th style="width:60px">名次</th>
                        <th>姓名</th>
                        <th style="width:80px">班级</th>
                        <th style="width:100px">学号</th>
                        <th style="width:80px">总分</th>
                        <th>每日得分</th>
                    </tr>
                </thead>
                <tbody>
    """
    
    for rank, p in enumerate(sorted_members, 1):
        rank_class = "rank-1" if rank == 1 else ""
        # 生成每日得分小标签
        daily_html = ""
        score_count = 0
        for i, score in enumerate(p["daily_scores"]):
            if score > 0:
                score_count += 1
                date_str = p["dates"][i] if i < len(p["dates"]) else f"D{i+1}"
                # 格式化日期，只显示月-日
                if isinstance(date_str, str) and "00:00" in date_str:
                    try:
                        date_str = date_str[5:10]  # 取 "03-02" 这样的格式
                    except:
                        date_str = f"D{i+1}"
                daily_html += f'<span class="daily-score" title="{date_str}">{int(score)}</span>'
        
        # 如果没有分数，显示横线
        if score_count == 0:
            daily_html = '<span style="color:#bdc3c7; font-size:0.8rem;">—</span>'
        
        # 处理学号显示
        student_id_display = ""
        if isinstance(p['student_id'], (int, float)):
            student_id_display = str(int(p['student_id']))
        else:
            student_id_display = p['student_id']
        
        html += f"""
                    <tr class="{rank_class}" data-name="{p['name_cn']} {p['name_en']} {student_id_display}">
                        <td>{rank}</td>
                        <td><strong>{p['name_cn']}</strong><br><span style="font-size:0.8rem; color:#7f8c8d;">{p['name_en'][:20] + '...' if len(p['name_en']) > 20 else p['name_en']}</span></td>
                        <td>{p['class']}</td>
                        <td>{student_id_display}</td>
                        <td><span class="score-badge">{int(p['total'])}</span></td>
                        <td><div class="daily-scores">{daily_html}</div></td>
                    </tr>
        """
    
    html += """
                </tbody>
            </table>
        </div>
    """

# 添加页脚和搜索脚本
html += """
        <div class="footer">
            ⚖️ 训育处  ·  数据每日自动更新  ·  仅供参考
        </div>
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
                const idCell = row.querySelector('td:nth-child(4)');
                const idText = idCell ? idCell.innerText.toLowerCase() : '';
                
                if (nameAttr.includes(searchTerm) || nameText.includes(searchTerm) || idText.includes(searchTerm)) {
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
