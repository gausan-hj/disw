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
try:
    df = pd.read_csv(url, header=None, encoding='utf-8-sig')
except:
    try:
        df = pd.read_csv(url, header=None, encoding='latin1')
    except:
        df = pd.read_csv(url, header=None, encoding='cp1252')

print(f"读取到 {len(df)} 行数据")

# ===== 成员名单 =====
members_list = [
    # 星穹组 (10人)
    {"group": "星穹组", "name_cn": "陈展艺", "name_en": "IVAN TAN ZHAN YI", "class": "S2FA", "student_id": "22038"},
    {"group": "星穹组", "name_cn": "侯展扬", "name_en": "HOW ZHAN YANG", "class": "S2Y", "student_id": "22100"},
    {"group": "星穹组", "name_cn": "邱嘉瑞", "name_en": "KATHERINE KHOO JAI RUI", "class": "J3T", "student_id": "24076"},
    {"group": "星穹组", "name_cn": "莎哈娜", "name_en": "SADHANA A/P SASU BAKIAN", "class": "J3T", "student_id": "24078"},
    {"group": "星穹组", "name_cn": "李韡翰", "name_en": "LEE WEI HANN", "class": "J3K", "student_id": "24068"},
    {"group": "星穹组", "name_cn": "彭绍洋", "name_en": "PEH SHAO YANG", "class": "J3T", "student_id": "24088"},
    {"group": "星穹组", "name_cn": "梁纹璇", "name_en": "LEONG WEN XUAN", "class": "J2Y", "student_id": "25035"},
    {"group": "星穹组", "name_cn": "尤嘉乐", "name_en": "JUSTIN YEW JIA LE", "class": "J2Y", "student_id": "25046"},
    {"group": "星穹组", "name_cn": "许艳棋", "name_en": "KHOR YAN QI", "class": "J1Y", "student_id": "26018"},
    {"group": "星穹组", "name_cn": "林隽毓", "name_en": "LIM JOON YI", "class": "J1Y", "student_id": "26032"},
    
    # 夜曜组 (11人)
    {"group": "夜曜组", "name_cn": "李竑证", "name_en": "LEE HOONG ZHENG", "class": "S2FA", "student_id": "22040"},
    {"group": "夜曜组", "name_cn": "廖若含", "name_en": "LIEW XIN YU", "class": "S2Y", "student_id": "22029"},
    {"group": "夜曜组", "name_cn": "林芷嫣", "name_en": "LIM ZHI YAN", "class": "S2Y", "student_id": "22083"},
    {"group": "夜曜组", "name_cn": "周柔慈", "name_en": "CHEO ROU ZHI", "class": "S2K", "student_id": "22051"},
    {"group": "夜曜组", "name_cn": "林骏喨", "name_en": "LIM TEIK LIANG", "class": "J3T", "student_id": "24083"},
    {"group": "夜曜组", "name_cn": "林宜彤", "name_en": "LIM YEE TONG", "class": "J2Y", "student_id": "25036"},
    {"group": "夜曜组", "name_cn": "潘宛瑜", "name_en": "TRACY PHUAH WANYU", "class": "J2Y", "student_id": "25071"},
    {"group": "夜曜组", "name_cn": "符传吉", "name_en": "FOO CHUAN JI", "class": "J2Y", "student_id": "25044"},
    {"group": "夜曜组", "name_cn": "陈欣怡", "name_en": "CINDY TAN XIN YI", "class": "J2F", "student_id": "25058"},
    {"group": "夜曜组", "name_cn": "丽亚", "name_en": "DHIYA ZULAIKHA DARWISYAH BINTI YUSNIZAN", "class": "J2F", "student_id": "25059"},
    {"group": "夜曜组", "name_cn": "郑宜桐", "name_en": "TEH YEE THONG", "class": "J1Y", "student_id": "26024"},
    
    # 沧澜组 (10人)
    {"group": "沧澜组", "name_cn": "浦源政", "name_en": "POH YUAN ZHENG", "class": "S2Y", "student_id": "22044"},
    {"group": "沧澜组", "name_cn": "吴贝优", "name_en": "GOH BEI YO", "class": "S2Y", "student_id": "22021"},
    {"group": "沧澜组", "name_cn": "林沛筠", "name_en": "LIM PEI JUN", "class": "S2Y", "student_id": "22030"},
    {"group": "沧澜组", "name_cn": "陈诗惠", "name_en": "CHAN SHI HUI", "class": "S2FA", "student_id": "22017"},
    {"group": "沧澜组", "name_cn": "郑憶欣", "name_en": "TEE YEE XIN", "class": "S1Y", "student_id": "23065"},
    {"group": "沧澜组", "name_cn": "谢楷棋", "name_en": "CHEAH KHAI QI", "class": "S1T", "student_id": "23013"},
    {"group": "沧澜组", "name_cn": "蔡善恩", "name_en": "CHUAH SHAN EN", "class": "J3F", "student_id": "24039"},
    {"group": "沧澜组", "name_cn": "许家绮", "name_en": "KOO JIA QI", "class": "J2Y", "student_id": "25031"},
    {"group": "沧澜组", "name_cn": "张子欣", "name_en": "TEON ZI XIN", "class": "J2F", "student_id": "25070"},
    {"group": "沧澜组", "name_cn": "施锦轩", "name_en": "SEE JIN XUAN", "class": "J1T", "student_id": "26092"}
]

# 自动匹配每一行
people = []

print("\n开始匹配成员...")

for member in members_list:
    found = False
    
    for i in range(len(df)):
        row = df.iloc[i].tolist()
        
        if len(row) > 4:
            row_name_cn = str(row[3]) if len(row) > 3 and pd.notna(row[3]) else ""
            row_name_en = str(row[4]) if len(row) > 4 and pd.notna(row[4]) else ""
            
            if (member["name_cn"] in row_name_cn or 
                member["name_en"][:20] in row_name_en[:20]):
                
                # 提取分数
                scores = []
                total = 0
                for j in range(7, len(row)):
                    val = row[j]
                    if pd.notna(val):
                        try:
                            num = float(val)
                            scores.append(num)
                            total += num
                        except (ValueError, TypeError):
                            scores.append(0)
                    else:
                        scores.append(0)
                
                people.append({
                    "group": member["group"],
                    "name_cn": member["name_cn"],
                    "name_en": member["name_en"],
                    "class": member["class"],
                    "student_id": member["student_id"],
                    "scores": scores,
                    "total": total
                })
                print(f"✓ 找到 {member['name_cn']} (总分: {total})")
                found = True
                break
    
    if not found:
        print(f"✗ 找不到 {member['name_cn']}")

print(f"\n总共找到 {len(people)} 人")

# 按组别整理
group_data = {g: [] for g in ["星穹组", "夜曜组", "沧澜组"]}
group_totals = {g: 0 for g in ["星穹组", "夜曜组", "沧澜组"]}

for p in people:
    group_data[p["group"]].append(p)
    group_totals[p["group"]] += p["total"]

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
            max-width: 1400px;
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
        
        /* 组排名卡片 */
        .rank-bar {{
            display: flex;
            gap: 16px;
            margin-bottom: 24px;
            flex-wrap: wrap;
        }}
        .rank-card {{
            flex: 1;
            min-width: 200px;
            background: white;
            border-radius: 12px;
            padding: 20px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.05);
            display: flex;
            align-items: center;
            border-left: 6px solid;
        }}
        .rank-1 {{ border-left-color: #ffd700; }}
        .rank-2 {{ border-left-color: #c0c0c0; }}
        .rank-3 {{ border-left-color: #cd7f32; }}
        
        .rank-icon {{
            font-size: 2rem;
            margin-right: 16px;
        }}
        .rank-info {{
            flex: 1;
        }}
        .rank-name {{
            font-size: 1.2rem;
            font-weight: 500;
            margin-bottom: 4px;
        }}
        .rank-score {{
            font-size: 1.5rem;
            font-weight: 600;
        }}
        .rank-score small {{
            font-size: 0.9rem;
            font-weight: 400;
            color: #7f8c8d;
        }}
        
        /* 三栏布局 */
        .group-container {{
            display: flex;
            gap: 24px;
            flex-wrap: wrap;
        }}
        .group-column {{
            flex: 1;
            min-width: 350px;
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
            font-size: 1.3rem;
            font-weight: 500;
        }}
        .group-rank-badge {{
            background: #2c3e50;
            color: white;
            padding: 4px 12px;
            border-radius: 30px;
            font-size: 0.9rem;
        }}
        
        /* 成员列表 */
        .member-list {{
            display: flex;
            flex-direction: column;
            gap: 8px;
        }}
        .member-row {{
            display: flex;
            align-items: center;
            padding: 8px;
            background: #f8fafc;
            border-radius: 8px;
            border-left: 3px solid transparent;
        }}
        .member-row:hover {{
            background: #f1f4f8;
        }}
        .member-info {{
            flex: 2;
            min-width: 120px;
        }}
        .member-name {{
            font-weight: 500;
        }}
        .member-class {{
            font-size: 0.75rem;
            color: #7f8c8d;
        }}
        
        /* 每日得分 - 小球设计 */
        .daily-scores {{
            flex: 3;
            display: flex;
            gap: 4px;
            flex-wrap: wrap;
            justify-content: flex-end;
        }}
        .score-dot {{
            width: 24px;
            height: 24px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 0.7rem;
            font-weight: 500;
            cursor: pointer;
            transition: transform 0.2s;
            background: #ecf0f1;
            color: #2c3e50;
        }}
        .score-dot:hover {{
            transform: scale(1.2);
            box-shadow: 0 2px 8px rgba(0,0,0,0.2);
        }}
        .score-positive {{
            background: #d4edda;
            color: #155724;
            font-weight: 600;
        }}
        
        /* 总分徽章 */
        .total-badge {{
            flex: 0.5;
            min-width: 50px;
            text-align: right;
            font-weight: 600;
            color: #2c3e50;
            background: #e9ecef;
            padding: 4px 8px;
            border-radius: 30px;
            margin-left: 8px;
            font-size: 0.9rem;
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
            .member-row {{
                flex-wrap: wrap;
            }}
            .daily-scores {{
                margin-top: 8px;
                justify-content: flex-start;
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
                <input type="text" id="search" placeholder="🔍 输入姓名搜索...">
            </div>
        </div>

        <!-- 组排名 -->
        <div class="rank-bar">
"""

# 添加组排名卡片
rank_icons = {1: "🥇", 2: "🥈", 3: "🥉"}
for i, (g, total) in enumerate(sorted_groups, 1):
    html += f"""
            <div class="rank-card rank-{i}">
                <div class="rank-icon">{rank_icons[i]}</div>
                <div class="rank-info">
                    <div class="rank-name">{g}</div>
                    <div class="rank-score">{int(total)} <small>分</small></div>
                </div>
            </div>
"""

html += """
        </div>

        <!-- 三组并列 -->
        <div class="group-container">
"""

# 按顺序添加三个组
for group_name in ["星穹组", "夜曜组", "沧澜组"]:
    members = group_data[group_name]
    rank = group_rank[group_name]
    
    html += f"""
            <div class="group-column">
                <div class="group-header">
                    <span class="group-title">{group_name}</span>
                    <span class="group-rank-badge">第{rank}名 · {int(group_totals[group_name])}分</span>
                </div>
                <div class="member-list">
    """
    
    for member in members:
        # 生成每日得分小球
        dots_html = ""
        for i, score in enumerate(member["scores"]):
            if score > 0:
                dots_html += f'<span class="score-dot score-positive" title="Day{i+1}: {int(score)}分">{int(score)}</span>'
            else:
                dots_html += f'<span class="score-dot" title="Day{i+1}: 0分">-</span>'
        
        # 英文名截断
        name_en_short = member['name_en'][:15] + "..." if len(member['name_en']) > 15 else member['name_en']
        
        html += f"""
                    <div class="member-row" data-name="{member['name_cn']} {member['name_en']}">
                        <div class="member-info">
                            <div class="member-name">{member['name_cn']}</div>
                            <div class="member-class">{member['class']} · {name_en_short}</div>
                        </div>
                        <div class="daily-scores">
                            {dots_html}
                        </div>
                        <div class="total-badge">{int(member['total'])}</div>
                    </div>
        """
    
    html += """
                </div>
            </div>
    """

html += """
        </div>
        <div class="footer">
            训育处 · 数据每日自动更新 · 鼠标悬停小球查看每日分数
        </div>
    </div>

    <script>
        const searchInput = document.getElementById('search');
        const allRows = document.querySelectorAll('.member-row');
        
        searchInput.addEventListener('input', (e) => {
            const searchTerm = e.target.value.toLowerCase().trim();
            
            if (searchTerm === '') {
                allRows.forEach(row => row.style.display = 'flex');
                return;
            }
            
            allRows.forEach(row => {
                const searchText = row.getAttribute('data-name').toLowerCase();
                if (searchText.includes(searchTerm)) {
                    row.style.display = 'flex';
                } else {
                    row.style.display = 'none';
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

print(f"\n✅ 生成成功！共 {len(people)} 人")
for g in ["星穹组", "夜曜组", "沧澜组"]:
    print(f"  {g}: {len(group_data[g])}人, 总分: {int(group_totals[g])}, 第{group_rank[g]}名")
