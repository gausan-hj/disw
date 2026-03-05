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
    # 星穹组 (10人)
    {"group": "星穹组", "name_cn": "陈展艺", "name_en": "IVAN TAN ZHAN YI", "class": "S2FA", "student_id": "22038", "row_index": 3},
    {"group": "星穹组", "name_cn": "侯展扬", "name_en": "HOW ZHAN YANG", "class": "S2Y", "student_id": "22100", "row_index": 4},
    {"group": "星穹组", "name_cn": "邱嘉瑞", "name_en": "KATHERINE KHOO JAI RUI", "class": "J3T", "student_id": "24076", "row_index": 5},
    {"group": "星穹组", "name_cn": "莎哈娜", "name_en": "SADHANA A/P SASU BAKIAN", "class": "J3T", "student_id": "24078", "row_index": 6},
    {"group": "星穹组", "name_cn": "李韡翰", "name_en": "LEE WEI HANN", "class": "J3K", "student_id": "24068", "row_index": 7},
    {"group": "星穹组", "name_cn": "彭绍洋", "name_en": "PEH SHAO YANG", "class": "J3T", "student_id": "24088", "row_index": 8},
    {"group": "星穹组", "name_cn": "梁纹璇", "name_en": "LEONG WEN XUAN", "class": "J2Y", "student_id": "25035", "row_index": 9},
    {"group": "星穹组", "name_cn": "尤嘉乐", "name_en": "JUSTIN YEW JIA LE", "class": "J2Y", "student_id": "25046", "row_index": 10},
    {"group": "星穹组", "name_cn": "许艳棋", "name_en": "KHOR YAN QI", "class": "J1Y", "student_id": "26018", "row_index": 11},
    {"group": "星穹组", "name_cn": "林隽毓", "name_en": "LIM JOON YI", "class": "J1Y", "student_id": "26032", "row_index": 12},
    
    # 夜曜组 (11人)
    {"group": "夜曜组", "name_cn": "李竑证", "name_en": "LEE HOONG ZHENG", "class": "S2FA", "student_id": "22040", "row_index": 15},
    {"group": "夜曜组", "name_cn": "廖若含", "name_en": "LIEW XIN YU", "class": "S2Y", "student_id": "22029", "row_index": 16},
    {"group": "夜曜组", "name_cn": "林芷嫣", "name_en": "LIM ZHI YAN", "class": "S2Y", "student_id": "22083", "row_index": 17},
    {"group": "夜曜组", "name_cn": "周柔慈", "name_en": "CHEO ROU ZHI", "class": "S2K", "student_id": "22051", "row_index": 18},
    {"group": "夜曜组", "name_cn": "林骏喨", "name_en": "LIM TEIK LIANG", "class": "J3T", "student_id": "24083", "row_index": 19},
    {"group": "夜曜组", "name_cn": "林宜彤", "name_en": "LIM YEE TONG", "class": "J2Y", "student_id": "25036", "row_index": 20},
    {"group": "夜曜组", "name_cn": "潘宛瑜", "name_en": "TRACY PHUAH WANYU", "class": "J2Y", "student_id": "25071", "row_index": 21},
    {"group": "夜曜组", "name_cn": "符传吉", "name_en": "FOO CHUAN JI", "class": "J2Y", "student_id": "25044", "row_index": 22},
    {"group": "夜曜组", "name_cn": "陈欣怡", "name_en": "CINDY TAN XIN YI", "class": "J2F", "student_id": "25058", "row_index": 23},
    {"group": "夜曜组", "name_cn": "丽亚", "name_en": "DHIYA ZULAIKHA DARWISYAH BINTI YUSNIZAN", "class": "J2F", "student_id": "25059", "row_index": 24},
    {"group": "夜曜组", "name_cn": "郑宜桐", "name_en": "TEH YEE THONG", "class": "J1Y", "student_id": "26024", "row_index": 25},
    
    # 沧澜组 (10人)
    {"group": "沧澜组", "name_cn": "浦源政", "name_en": "POH YUAN ZHENG", "class": "S2Y", "student_id": "22044", "row_index": 28},
    {"group": "沧澜组", "name_cn": "吴贝优", "name_en": "GOH BEI YO", "class": "S2Y", "student_id": "22021", "row_index": 29},
    {"group": "沧澜组", "name_cn": "林沛筠", "name_en": "LIM PEI JUN", "class": "S2Y", "student_id": "22030", "row_index": 30},
    {"group": "沧澜组", "name_cn": "陈诗惠", "name_en": "CHAN SHI HUI", "class": "S2FA", "student_id": "22017", "row_index": 31},
    {"group": "沧澜组", "name_cn": "郑憶欣", "name_en": "TEE YEE XIN", "class": "S1Y", "student_id": "23065", "row_index": 32},
    {"group": "沧澜组", "name_cn": "谢楷棋", "name_en": "CHEAH KHAI QI", "class": "S1T", "student_id": "23013", "row_index": 33},
    {"group": "沧澜组", "name_cn": "蔡善恩", "name_en": "CHUAH SHAN EN", "class": "J3F", "student_id": "24039", "row_index": 34},
    {"group": "沧澜组", "name_cn": "许家绮", "name_en": "KOO JIA QI", "class": "J2Y", "student_id": "25031", "row_index": 35},
    {"group": "沧澜组", "name_cn": "张子欣", "name_en": "TEON ZI XIN", "class": "J2F", "student_id": "25070", "row_index": 36},
    {"group": "沧澜组", "name_cn": "施锦轩", "name_en": "SEE JIN XUAN", "class": "J1T", "student_id": "26092", "row_index": 37}
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
            if pd.notna(row[j]) and isinstance(row[j], (int, float)):
                total_score += row[j]
        
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
group_rank = {}
for i, (g, _) in enumerate(sorted_groups, 1):
    group_rank[g] = i

# 生成HTML（用之前的版本）
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
            border-left: 6px solid #b0bec5;
        }}
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
            <div class="rank-item">
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
        
        # 按原顺序显示
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
