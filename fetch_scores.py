import pandas as pd
from datetime import datetime
import base64
import io

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

# 获取日期（第一行）
dates = []
if len(df) > 0:
    first_row = df.iloc[0].tolist()
    for j in range(7, len(first_row)):
        if pd.notna(first_row[j]):
            date_str = str(first_row[j])
            try:
                if "00:00" in date_str:
                    date_str = date_str[5:10]
                dates.append(date_str)
            except:
                dates.append("")
        else:
            dates.append("")

print(f"找到 {len(dates)} 个日期")

# ===== 成员名单 =====
members_list = [
    # 星穹组 (10人)
    {"group": "星穹组", "name_cn": "陈展艺", "name_en": "IVAN TAN ZHAN YI", "class": "S2FA", "student_id": "22038", "order": 1},
    {"group": "星穹组", "name_cn": "侯展扬", "name_en": "HOW ZHAN YANG", "class": "S2Y", "student_id": "22100", "order": 2},
    {"group": "星穹组", "name_cn": "邱嘉瑞", "name_en": "KATHERINE KHOO JAI RUI", "class": "J3T", "student_id": "24076", "order": 3},
    {"group": "星穹组", "name_cn": "莎哈娜", "name_en": "SADHANA A/P SASU BAKIAN", "class": "J3T", "student_id": "24078", "order": 4},
    {"group": "星穹组", "name_cn": "李韡翰", "name_en": "LEE WEI HANN", "class": "J3K", "student_id": "24068", "order": 5},
    {"group": "星穹组", "name_cn": "彭绍洋", "name_en": "PEH SHAO YANG", "class": "J3T", "student_id": "24088", "order": 6},
    {"group": "星穹组", "name_cn": "梁纹璇", "name_en": "LEONG WEN XUAN", "class": "J2Y", "student_id": "25035", "order": 7},
    {"group": "星穹组", "name_cn": "尤嘉乐", "name_en": "JUSTIN YEW JIA LE", "class": "J2Y", "student_id": "25046", "order": 8},
    {"group": "星穹组", "name_cn": "许艳棋", "name_en": "KHOR YAN QI", "class": "J1Y", "student_id": "26018", "order": 9},
    {"group": "星穹组", "name_cn": "林隽毓", "name_en": "LIM JOON YI", "class": "J1Y", "student_id": "26032", "order": 10},
    
    # 夜曜组 (11人)
    {"group": "夜曜组", "name_cn": "李竑证", "name_en": "LEE HOONG ZHENG", "class": "S2FA", "student_id": "22040", "order": 1},
    {"group": "夜曜组", "name_cn": "廖若含", "name_en": "LIEW XIN YU", "class": "S2Y", "student_id": "22029", "order": 2},
    {"group": "夜曜组", "name_cn": "林芷嫣", "name_en": "LIM ZHI YAN", "class": "S2Y", "student_id": "22083", "order": 3},
    {"group": "夜曜组", "name_cn": "周柔慈", "name_en": "CHEO ROU ZHI", "class": "S2K", "student_id": "22051", "order": 4},
    {"group": "夜曜组", "name_cn": "林骏喨", "name_en": "LIM TEIK LIANG", "class": "J3T", "student_id": "24083", "order": 5},
    {"group": "夜曜组", "name_cn": "林宜彤", "name_en": "LIM YEE TONG", "class": "J2Y", "student_id": "25036", "order": 6},
    {"group": "夜曜组", "name_cn": "潘宛瑜", "name_en": "TRACY PHUAH WANYU", "class": "J2Y", "student_id": "25071", "order": 7},
    {"group": "夜曜组", "name_cn": "符传吉", "name_en": "FOO CHUAN JI", "class": "J2Y", "student_id": "25044", "order": 8},
    {"group": "夜曜组", "name_cn": "陈欣怡", "name_en": "CINDY TAN XIN YI", "class": "J2F", "student_id": "25058", "order": 9},
    {"group": "夜曜组", "name_cn": "丽亚", "name_en": "DHIYA ZULAIKHA DARWISYAH BINTI YUSNIZAN", "class": "J2F", "student_id": "25059", "order": 10},
    {"group": "夜曜组", "name_cn": "郑宜桐", "name_en": "TEH YEE THONG", "class": "J1Y", "student_id": "26024", "order": 11},
    
    # 沧澜组 (10人)
    {"group": "沧澜组", "name_cn": "浦源政", "name_en": "POH YUAN ZHENG", "class": "S2Y", "student_id": "22044", "order": 1},
    {"group": "沧澜组", "name_cn": "吴贝优", "name_en": "GOH BEI YO", "class": "S2Y", "student_id": "22021", "order": 2},
    {"group": "沧澜组", "name_cn": "林沛筠", "name_en": "LIM PEI JUN", "class": "S2Y", "student_id": "22030", "order": 3},
    {"group": "沧澜组", "name_cn": "陈诗惠", "name_en": "CHAN SHI HUI", "class": "S2FA", "student_id": "22017", "order": 4},
    {"group": "沧澜组", "name_cn": "郑憶欣", "name_en": "TEE YEE XIN", "class": "S1Y", "student_id": "23065", "order": 5},
    {"group": "沧澜组", "name_cn": "谢楷棋", "name_en": "CHEAH KHAI QI", "class": "S1T", "student_id": "23013", "order": 6},
    {"group": "沧澜组", "name_cn": "蔡善恩", "name_en": "CHUAH SHAN EN", "class": "J3F", "student_id": "24039", "order": 7},
    {"group": "沧澜组", "name_cn": "许家绮", "name_en": "KOO JIA QI", "class": "J2Y", "student_id": "25031", "order": 8},
    {"group": "沧澜组", "name_cn": "张子欣", "name_en": "TEON ZI XIN", "class": "J2F", "student_id": "25070", "order": 9},
    {"group": "沧澜组", "name_cn": "施锦轩", "name_en": "SEE JIN XUAN", "class": "J1T", "student_id": "26092", "order": 10}
]

# 自动匹配每一行
people = []
group_totals = {g: 0 for g in ["星穹组", "夜曜组", "沧澜组"]}
group_members_count = {g: 0 for g in ["星穹组", "夜曜组", "沧澜组"]}

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
                
                total = 0
                score_dict = {}
                
                for j in range(7, len(row)):
                    if j-7 < len(dates):
                        date = dates[j-7]
                        val = row[j]
                        if pd.notna(val):
                            try:
                                num = float(val)
                                total += num
                                if num > 0:
                                    score_dict[date] = num
                            except (ValueError, TypeError):
                                pass
                
                people.append({
                    "group": member["group"],
                    "order": member["order"],
                    "name_cn": member["name_cn"],
                    "name_en": member["name_en"],
                    "class": member["class"],
                    "student_id": member["student_id"],
                    "score_dict": score_dict,
                    "total": total
                })
                print(f"✓ 找到 {member['name_cn']} (总分: {total})")
                found = True
                break
    
    if not found:
        print(f"✗ 找不到 {member['name_cn']}")

print(f"\n总共找到 {len(people)} 人")

if len(people) == 0:
    print("❌ 没有找到任何人！请检查：")
    print("1. Google Sheets 权限是否设置为 '任何知道链接的人可查看'")
    print("2. Sheet 名字是否正确（当前是 Sheet3）")
    print("3. 数据格式是否和 Excel 一致")
    exit(1)

# 按组别整理
group_data = {g: [] for g in ["星穹组", "夜曜组", "沧澜组"]}
for p in people:
    group_data[p["group"]].append(p)
    group_totals[p["group"]] += p["total"]
    group_members_count[p["group"]] += 1

# 每个组内按order排序
for g in group_data:
    group_data[g].sort(key=lambda x: x["order"])

# 计算当前排名
sorted_groups = sorted(group_totals.items(), key=lambda x: x[1], reverse=True)
group_rank = {}
for i, (g, _) in enumerate(sorted_groups, 1):
    group_rank[g] = i

# 计算每组平均分
group_averages = {}
for g in ["星穹组", "夜曜组", "沧澜组"]:
    if group_data[g]:
        group_averages[g] = group_totals[g] / len(group_data[g])
    else:
        group_averages[g] = 0

# 计算每个人的达标状态
for g in group_data:
    for p in group_data[g]:
        rank = group_rank[g]
        avg = group_averages[g]
        
        if rank == 1:  # 第一名组别
            p["reward_status"] = "✅" if p["total"] >= avg / 2 else "❌"
            p["reward_class"] = "reward-pass" if p["total"] >= avg / 2 else "reward-fail"
        elif rank == 2:  # 第二名组别
            p["reward_status"] = "✅" if p["total"] >= avg else "❌"
            p["reward_class"] = "reward-pass" if p["total"] >= avg else "reward-fail"
        else:  # 第三名组别
            # 找出组内前三名
            top3 = sorted(group_data[g], key=lambda x: x["total"], reverse=True)[:3]
            top3_names = [t["name_cn"] for t in top3]
            if p["name_cn"] in top3_names:
                p["reward_status"] = "✅"
                p["reward_class"] = "reward-pass"
            else:
                p["reward_status"] = "❌"
                p["reward_class"] = "reward-fail"

# 生成HTML - 融合版（无双击切换，无热力图）
html = '''<!DOCTYPE html>
<html lang="zh">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=yes, viewport-fit=cover">
    <title>训育处 - 学长团分数板</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
            -webkit-tap-highlight-color: transparent;
        }

        :root {
            --bg-primary: #f5f7fc;
            --bg-secondary: #ffffff;
            --card-bg: #ffffff;
            --text-primary: #1a2b3c;
            --text-secondary: #2c3e50;
            --text-tertiary: #5a6b7a;
            --border-light: #e1e8f0;
            --border-subtle: #eef2f6;
            --shadow-sm: 0 2px 8px rgba(0,0,0,0.02);
            --shadow-md: 0 4px 12px rgba(0,0,0,0.04);
            --shadow-lg: 0 8px 20px rgba(0,0,0,0.06);
            
            --star-primary: #eab308;
            --star-light: #fef9c3;
            --star-bg: #fefae8;
            --night-primary: #a855f7;
            --night-light: #f3e8ff;
            --night-bg: #faf5ff;
            --ocean-primary: #3b82f6;
            --ocean-light: #dbeafe;
            --ocean-bg: #f0f7ff;
            
            --score-bg: #f1f5f9;
            --score-text: #475569;
            --score-highlight: #dbeafe;
            --score-highlight-text: #1e40af;
            --reward-pass: #dcfce7;
            --reward-pass-text: #166534;
            --reward-fail: #fee2e2;
            --reward-fail-text: #991b1b;
            
            --safe-top: env(safe-area-inset-top);
            --safe-bottom: env(safe-area-inset-bottom);
        }

        body.night-mode {
            --bg-primary: #0f1825;
            --bg-secondary: #1e2a3a;
            --card-bg: #1f2c3d;
            --text-primary: #e6edf5;
            --text-secondary: #cbd5e1;
            --text-tertiary: #94a3b8;
            --border-light: #2d3a4d;
            --border-subtle: #253141;
            --shadow-sm: 0 2px 8px rgba(0,0,0,0.3);
            --shadow-md: 0 4px 12px rgba(0,0,0,0.4);
            --shadow-lg: 0 8px 20px rgba(0,0,0,0.5);
            
            --star-light: #423d2a;
            --star-bg: #2f2a1f;
            --night-light: #2f2740;
            --night-bg: #252033;
            --ocean-light: #1f3045;
            --ocean-bg: #1c2638;
            
            --score-bg: #2d3a4f;
            --score-text: #cbd5e1;
            --score-highlight: #1e3a6f;
            --score-highlight-text: #9ac7ff;
            --reward-pass: #1f4a3a;
            --reward-pass-text: #bbf7d0;
            --reward-fail: #562b2b;
            --reward-fail-text: #fecaca;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Microsoft YaHei', sans-serif;
            background: var(--bg-primary);
            color: var(--text-primary);
            padding: 12px;
            padding-top: max(12px, var(--safe-top));
            padding-bottom: max(12px, var(--safe-bottom));
            min-height: 100vh;
            transition: background 0.2s ease, color 0.2s ease;
            line-height: 1.4;
        }

        .container {
            max-width: 100%;
            margin: 0 auto;
        }

        /* 头部 */
        .header {
            background: var(--card-bg);
            border-radius: 24px;
            padding: 16px;
            margin-bottom: 16px;
            box-shadow: var(--shadow-md);
            border: 1px solid var(--border-subtle);
        }

        .header-top {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 12px;
            flex-wrap: wrap;
            gap: 8px;
        }

        .title-group {
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .school-icon {
            font-size: 1.6rem;
        }

        h1 {
            font-size: 1.3rem;
            font-weight: 600;
            color: var(--text-primary);
        }

        .action-buttons {
            display: flex;
            gap: 8px;
            align-items: center;
        }

        /* 深色模式按钮 */
        .theme-toggle {
            background: var(--bg-primary);
            border: 1px solid var(--border-light);
            border-radius: 30px;
            padding: 6px 12px;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 4px;
            font-size: 0.8rem;
            color: var(--text-primary);
            transition: background 0.15s ease;
            white-space: nowrap;
        }

        .theme-toggle:active {
            background: var(--border-light);
            transform: scale(0.98);
        }

        .moon-icon {
            display: inline-block;
            font-size: 1rem;
            transition: transform 0.2s ease;
        }

        body.night-mode .moon-icon {
            transform: rotate(0deg);
        }

        body:not(.night-mode) .moon-icon {
            transform: rotate(180deg);
        }

        /* 下载按钮 */
        .download-btn {
            background: var(--star-primary);
            border: none;
            border-radius: 30px;
            padding: 6px 12px;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 4px;
            font-size: 0.8rem;
            color: #1a2b3c;
            font-weight: 500;
            transition: all 0.15s ease;
            white-space: nowrap;
        }

        .download-btn:active {
            transform: scale(0.98);
            opacity: 0.9;
        }

        body.night-mode .download-btn {
            color: #ffffff;
        }

        .meta-info {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 12px;
            padding-bottom: 8px;
            border-bottom: 1px dashed var(--border-subtle);
            color: var(--text-tertiary);
            font-size: 0.7rem;
        }

        .date-badge {
            background: var(--bg-primary);
            padding: 3px 10px;
            border-radius: 20px;
            font-size: 0.65rem;
            border: 1px solid var(--border-subtle);
            white-space: nowrap;
        }

        .search-wrapper {
            position: relative;
        }

        .search-icon {
            position: absolute;
            left: 12px;
            top: 50%;
            transform: translateY(-50%);
            color: var(--text-tertiary);
            font-size: 0.9rem;
        }

        #search {
            width: 100%;
            padding: 10px 12px 10px 36px;
            border: 1px solid var(--border-subtle);
            border-radius: 30px;
            font-size: 0.85rem;
            background: var(--bg-primary);
            color: var(--text-primary);
            -webkit-appearance: none;
        }

        #search:focus {
            outline: none;
            border-color: var(--text-tertiary);
        }

        /* 统计图卡片 */
        .chart-card {
            background: var(--card-bg);
            border-radius: 20px;
            padding: 16px;
            margin-bottom: 16px;
            box-shadow: var(--shadow-md);
            border: 1px solid var(--border-subtle);
            display: none;
        }

        .chart-card.show {
            display: block;
            animation: slideDown 0.3s ease;
        }

        @keyframes slideDown {
            from {
                opacity: 0;
                transform: translateY(-10px);
            }
            to {
                opacity: 1;
                transform: translateY(0);
            }
        }

        .chart-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 12px;
        }

        .chart-title {
            font-size: 1rem;
            font-weight: 600;
            display: flex;
            align-items: center;
            gap: 6px;
        }

        .chart-actions {
            display: flex;
            gap: 8px;
        }

        .save-chart-btn {
            background: var(--star-primary);
            border: none;
            border-radius: 20px;
            padding: 4px 12px;
            font-size: 0.7rem;
            cursor: pointer;
            color: #1a2b3c;
            font-weight: 500;
            display: flex;
            align-items: center;
            gap: 4px;
        }

        .close-chart {
            background: var(--bg-primary);
            border: 1px solid var(--border-subtle);
            border-radius: 20px;
            padding: 4px 10px;
            font-size: 0.7rem;
            cursor: pointer;
            color: var(--text-secondary);
        }

        .chart-container {
            position: relative;
            height: 200px;
            width: 100%;
            margin-bottom: 16px;
        }

        .stats-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 8px;
            margin-top: 12px;
        }

        .stat-item {
            background: var(--bg-primary);
            border-radius: 16px;
            padding: 10px;
            text-align: center;
        }

        .stat-label {
            font-size: 0.65rem;
            color: var(--text-tertiary);
            margin-bottom: 4px;
        }

        .stat-value {
            font-size: 1rem;
            font-weight: 700;
            color: var(--text-primary);
        }

        .stat-rank {
            font-size: 0.6rem;
            color: var(--text-secondary);
            margin-top: 2px;
        }

        /* 组排名卡片 */
        .rank-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 8px;
            margin-bottom: 16px;
        }

        .rank-card {
            background: var(--card-bg);
            border-radius: 18px;
            padding: 10px 8px;
            box-shadow: var(--shadow-sm);
            border: 1px solid var(--border-subtle);
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 6px;
            border-left: 3px solid;
            transition: transform 0.1s ease;
        }

        .rank-card:active {
            transform: scale(0.98);
        }

        .rank-card[data-group="星穹组"] {
            border-left-color: #eab308 !important;
            background: linear-gradient(to right, var(--star-bg), var(--card-bg));
        }
        .rank-card[data-group="夜曜组"] {
            border-left-color: #a855f7 !important;
            background: linear-gradient(to right, var(--night-bg), var(--card-bg));
        }
        .rank-card[data-group="沧澜组"] {
            border-left-color: #3b82f6 !important;
            background: linear-gradient(to right, var(--ocean-bg), var(--card-bg));
        }

        .rank-icon {
            font-size: 1.4rem;
        }

        .rank-info {
            flex: 1;
        }

        .rank-name {
            font-weight: 600;
            font-size: 0.75rem;
            margin-bottom: 2px;
        }

        .rank-score {
            font-weight: 700;
            font-size: 1rem;
            display: flex;
            align-items: baseline;
            gap: 2px;
        }

        .rank-score small {
            font-size: 0.55rem;
            font-weight: 400;
            color: var(--text-tertiary);
        }

        /* 组卡片 */
        .groups {
            display: flex;
            flex-direction: column;
            gap: 12px;
        }

        .group-card {
            background: var(--card-bg);
            border-radius: 20px;
            padding: 14px;
            box-shadow: var(--shadow-md);
            border: 1px solid var(--border-subtle);
            scroll-margin-top: 12px;
            border-top: 3px solid;
        }

        .group-card[data-group="星穹组"] { border-top-color: #eab308 !important; }
        .group-card[data-group="夜曜组"] { border-top-color: #a855f7 !important; }
        .group-card[data-group="沧澜组"] { border-top-color: #3b82f6 !important; }

        .group-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 10px;
            padding-bottom: 8px;
            border-bottom: 1px solid var(--border-subtle);
        }

        .group-title-wrapper {
            display: flex;
            align-items: center;
            gap: 6px;
        }

        .group-emoji {
            font-size: 1.2rem;
        }

        .group-title {
            font-size: 1.1rem;
            font-weight: 600;
        }

        .group-stats {
            display: flex;
            gap: 6px;
            align-items: center;
        }

        .group-avg {
            font-size: 0.7rem;
            padding: 3px 8px;
            background: var(--bg-primary);
            border-radius: 20px;
            color: var(--text-secondary);
            border: 1px solid var(--border-subtle);
            white-space: nowrap;
        }

        .group-badge {
            font-size: 0.7rem;
            padding: 3px 8px;
            border-radius: 20px;
            color: white;
            white-space: nowrap;
        }

        .group-card[data-group="星穹组"] .group-badge { background: #eab308; }
        .group-card[data-group="夜曜组"] .group-badge { background: #a855f7; }
        .group-card[data-group="沧澜组"] .group-badge { background: #3b82f6; }

        /* 表格 */
        .table-container {
            overflow-x: auto;
            -webkit-overflow-scrolling: touch;
            margin: 0 -4px;
            padding: 0 4px;
        }

        .member-table {
            width: 100%;
            border-collapse: collapse;
            min-width: 750px;  /* 增加宽度以容纳完整英文名 */
            font-size: 0.8rem;
        }

        .member-table th {
            text-align: left;
            padding: 8px 4px;
            font-size: 0.65rem;
            font-weight: 600;
            color: var(--text-tertiary);
            text-transform: uppercase;
            border-bottom: 1px solid var(--border-subtle);
        }

        .member-table td {
            padding: 8px 4px;
            border-bottom: 1px solid var(--border-subtle);
        }

        .member-table tr:last-child td {
            border-bottom: none;
        }

        .member-table th:nth-child(1) { width: 30px; text-align: center; }
        .member-table th:nth-child(2) { width: 180px; }  /* 增加姓名列宽度 */
        .member-table th:nth-child(3) { width: 45px; }
        .member-table th:nth-child(4) { width: 60px; }
        .member-table th:nth-child(5) { width: auto; }
        .member-table th:nth-child(6) { width: 40px; text-align: right; }
        .member-table th:nth-child(7) { width: 35px; text-align: center; }

        .member-table td:nth-child(1) { text-align: center; }
        .member-table td:nth-child(6) { text-align: right; font-weight: 600; }
        .member-table td:nth-child(7) { text-align: center; }

        .name-cell {
            max-width: 180px;
        }

        .name-cn {
            font-weight: 600;
            font-size: 0.8rem;
            margin-bottom: 2px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .name-en {
            font-size: 0.65rem;
            color: var(--text-tertiary);
            white-space: normal;  /* 允许换行 */
            line-height: 1.3;
            word-break: break-word;  /* 长单词换行 */
        }

        .info-cell {
            font-size: 0.7rem;
            color: var(--text-secondary);
        }

        /* 分数标签 */
        .score-tags {
            display: flex;
            gap: 2px;
            flex-wrap: wrap;
        }

        .score-item {
            padding: 2px 5px;
            border-radius: 12px;
            font-size: 0.55rem;
            font-weight: 500;
            background: var(--score-bg);
            color: var(--score-text);
            border: 1px solid var(--border-subtle);
        }

        .score-item.has-score {
            background: var(--score-highlight);
            color: var(--score-highlight-text);
        }

        .score-date { opacity: 0.7; margin-right: 1px; }
        .score-value { font-weight: 600; }

        .reward-pass, .reward-fail {
            padding: 2px 5px;
            border-radius: 12px;
            font-size: 0.55rem;
            font-weight: 600;
            display: inline-block;
            min-width: 30px;
            text-align: center;
        }

        .reward-pass {
            background: var(--reward-pass);
            color: var(--reward-pass-text);
        }

        .reward-fail {
            background: var(--reward-fail);
            color: var(--reward-fail-text);
        }

        /* 奖励机制卡片 - 从测试网站融合过来 */
        .reward-section {
            background: linear-gradient(135deg, var(--card-bg) 0%, var(--bg-primary) 100%);
            border-radius: 20px;
            padding: 16px;
            margin-top: 20px;
            margin-bottom: 16px;
            box-shadow: var(--shadow-lg);
            border: 1px solid var(--border-light);
            border-left: 4px solid #eab308;
        }

        .reward-header {
            display: flex;
            align-items: center;
            gap: 8px;
            margin-bottom: 12px;
        }

        .reward-icon {
            font-size: 1.6rem;
        }

        .reward-header h2 {
            font-size: 1.1rem;
            font-weight: 600;
        }

        .reward-subtitle {
            font-size: 0.7rem;
            color: var(--text-tertiary);
            margin-bottom: 12px;
            padding-left: 8px;
            border-left: 2px solid #eab308;
        }

        .reward-content {
            display: flex;
            flex-direction: column;
            gap: 8px;
            margin-bottom: 12px;
        }

        .reward-item {
            background: var(--bg-primary);
            border-radius: 16px;
            padding: 10px;
            border: 1px solid var(--border-subtle);
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .reward-rank {
            display: flex;
            align-items: center;
            gap: 6px;
            min-width: 60px;
        }

        .rank-medal {
            font-size: 1.4rem;
        }

        .rank-title {
            font-size: 0.9rem;
            font-weight: 600;
        }

        .reward-condition {
            background: var(--card-bg);
            border-radius: 16px;
            padding: 3px 8px;
            font-size: 0.65rem;
            font-weight: 500;
            border: 1px solid var(--border-light);
            margin-bottom: 4px;
            white-space: nowrap;
        }

        .reward-benefit {
            color: var(--star-primary);
            font-weight: 500;
            font-size: 0.65rem;
        }

        .reward-note {
            background: var(--bg-primary);
            border-radius: 16px;
            padding: 10px;
            margin-top: 8px;
            border: 1px dashed var(--border-light);
            font-size: 0.7rem;
            color: var(--text-secondary);
            text-align: center;
        }

        .reward-extra {
            color: var(--night-primary);
            font-weight: 500;
        }

        .reward-footer {
            margin-top: 12px;
            text-align: center;
            font-size: 0.65rem;
            color: var(--text-tertiary);
            font-style: italic;
        }

        .footer {
            margin-top: 16px;
            text-align: center;
            color: var(--text-tertiary);
            font-size: 0.6rem;
            padding: 12px;
            border-top: 1px solid var(--border-subtle);
        }

        /* 下载提示 */
        .download-toast {
            position: fixed;
            bottom: 20px;
            left: 50%;
            transform: translateX(-50%) translateY(100px);
            background: var(--card-bg);
            color: var(--text-primary);
            padding: 10px 20px;
            border-radius: 30px;
            box-shadow: var(--shadow-lg);
            border: 1px solid var(--border-light);
            display: flex;
            align-items: center;
            gap: 8px;
            transition: transform 0.3s ease;
            z-index: 1000;
            font-size: 0.8rem;
        }

        .download-toast.show {
            transform: translateX(-50%) translateY(0);
        }

        .toast-icon {
            width: 20px;
            height: 20px;
            border-radius: 50%;
            background: #eab308;
            color: #1a2b3c;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 0.7rem;
        }

        @media (max-width: 360px) {
            .rank-name { font-size: 0.7rem; }
            .rank-score { font-size: 0.9rem; }
            .group-title { font-size: 1rem; }
            .member-table { min-width: 650px; }
            .name-en { font-size: 0.6rem; }
            .action-buttons { width: 100%; justify-content: flex-end; }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <div class="header-top">
                <div class="title-group">
                    <span class="school-icon">🏫</span>
                    <h1>学长团分数板</h1>
                </div>
                <div class="action-buttons">
                    <button class="download-btn" id="downloadBtn">
                        <span>📊</span>
                        <span>下载统计</span>
                    </button>
                    <div class="theme-toggle" onclick="document.body.classList.toggle('night-mode')">
                        <span class="moon-icon">🌓</span>
                        <span>深色</span>
                    </div>
                </div>
            </div>
            <div class="meta-info">
                <span>Prefects' Scoreboard</span>
                <span class="date-badge">''' + datetime.now().strftime('%m/%d %H:%M') + '''</span>
            </div>
            <div class="search-wrapper">
                <span class="search-icon">🔍</span>
                <input type="text" id="search" placeholder="搜姓名/班级/学号...">
            </div>
        </div>

        <!-- 统计图卡片 -->
        <div class="chart-card" id="chartCard">
            <div class="chart-header">
                <span class="chart-title">
                    <span>📊</span>
                    各组总分对比
                </span>
                <div class="chart-actions">
                    <button class="save-chart-btn" id="saveChartBtn">
                        <span>💾</span>
                        <span>保存到相册</span>
                    </button>
                    <button class="close-chart" id="closeChart">关闭</button>
                </div>
            </div>
            <div class="chart-container">
                <canvas id="groupChart"></canvas>
            </div>
            <div class="stats-grid" id="statsGrid"></div>
        </div>

        <!-- 组排名卡片 -->
        <div class="rank-grid">
'''

# 添加组排名卡片
rank_icons = {1: "🥇", 2: "🥈", 3: "🥉"}
group_ids = {"星穹组": "group-xingqiong", "夜曜组": "group-yeyao", "沧澜组": "group-canglan"}
group_list = []
total_list = []

for i, (g, total) in enumerate(sorted_groups, 1):
    group_id = group_ids[g]
    group_list.append(g)
    total_list.append(int(total))
    
    html += f'''
            <div class="rank-card" data-group="{g}" onclick="document.getElementById('{group_id}').scrollIntoView({{behavior: 'smooth'}})">
                <span class="rank-icon">{rank_icons[i]}</span>
                <div class="rank-info">
                    <div class="rank-name">{g}</div>
                    <div class="rank-score">{int(total)}<small>分</small></div>
                </div>
            </div>
'''

# 准备统计数据
stats_data = []
for g in ["星穹组", "夜曜组", "沧澜组"]:
    if g in group_data:
        stats_data.append({
            "group": g,
            "total": int(group_totals[g]),
            "rank": group_rank[g],
            "members": len(group_data[g]),
            "avg": int(group_averages[g]),
            "color": "#eab308" if g == "星穹组" else "#a855f7" if g == "夜曜组" else "#3b82f6"
        })

html += '''
        </div>

        <!-- 组别详情 -->
        <div class="groups">
'''

# 添加三个组
group_emojis = {"星穹组": "✨", "夜曜组": "🌙", "沧澜组": "🌊"}
for group_name in ["星穹组", "夜曜组", "沧澜组"]:
    members = group_data[group_name]
    rank = group_rank[group_name]
    group_id = group_ids[group_name]
    avg_score = group_averages[group_name]
    
    html += f'''
            <div class="group-card" data-group="{group_name}" id="{group_id}">
                <div class="group-header">
                    <div class="group-title-wrapper">
                        <span class="group-emoji">{group_emojis[group_name]}</span>
                        <span class="group-title">{group_name}</span>
                    </div>
                    <div class="group-stats">
                        <span class="group-avg">平均{int(avg_score)}</span>
                        <span class="group-badge">第{rank}名</span>
                    </div>
                </div>
                <div class="table-container">
                    <table class="member-table">
                        <thead>
                            <tr>
                                <th>#</th>
                                <th>姓名</th>
                                <th>班</th>
                                <th>学号</th>
                                <th>每日得分</th>
                                <th>总分</th>
                                <th>奖</th>
                            </tr>
                        </thead>
                        <tbody>
'''

    for member in members:
        # 生成每日得分标签
        score_tags = ""
        sorted_dates = sorted(member["score_dict"].keys())
        for date in sorted_dates[-5:]:
            if date:
                score = member["score_dict"][date]
                score_tags += f'<span class="score-item has-score"><span class="score-date">{date}</span><span class="score-value">{int(score)}</span></span>'
        
        if not score_tags:
            score_tags = '<span class="score-item">—</span>'
        
        # 完整显示英文名（不截断）
        name_en_full = member['name_en']
        
        html += f'''
                        <tr data-search="{member['name_cn']} {member['name_en']} {member['class']} {member['student_id']}">
                            <td>{member['order']}</td>
                            <td class="name-cell">
                                <div class="name-cn">{member['name_cn']}</div>
                                <div class="name-en">{name_en_full}</div>
                            </td>
                            <td class="info-cell">{member['class']}</td>
                            <td class="info-cell">{member['student_id']}</td>
                            <td><div class="score-tags">{score_tags}</div></td>
                            <td>{int(member['total'])}</td>
                            <td><span class="{member['reward_class']}">{member['reward_status']}</span></td>
                        </tr>
'''

    html += '''
                        </tbody>
                    </table>
                </div>
            </div>
'''

html += '''
        </div>

        <!-- 奖励机制卡片 - 从测试网站融合过来 -->
        <div class="reward-section">
            <div class="reward-header">
                <span class="reward-icon">🎁</span>
                <h2>本轮奖励机制</h2>
            </div>
            <div class="reward-subtitle">
                公平原则 · 不负每一位付出的学长
            </div>
            <div class="reward-content">
                <div class="reward-item">
                    <div class="reward-rank">
                        <span class="rank-medal">🥇</span>
                        <span class="rank-title">第1名</span>
                    </div>
                    <div>
                        <div class="reward-condition">分数 ≥ 平均分÷2</div>
                        <div class="reward-benefit">✨免搬椅子+减免操步</div>
                    </div>
                </div>
                <div class="reward-item">
                    <div class="reward-rank">
                        <span class="rank-medal">🥈</span>
                        <span class="rank-title">第2名</span>
                    </div>
                    <div>
                        <div class="reward-condition">分数 ≥ 平均分</div>
                        <div class="reward-benefit">✨免搬椅子+减免操步</div>
                    </div>
                </div>
                <div class="reward-item">
                    <div class="reward-rank">
                        <span class="rank-medal">🥉</span>
                        <span class="rank-title">第3名</span>
                    </div>
                    <div>
                        <div class="reward-condition">组内前三名</div>
                        <div class="reward-benefit">✨免搬椅子+减免操步</div>
                    </div>
                </div>
            </div>
            <div class="reward-note">
                <span class="reward-extra">🎁 第1名达标者额外奖励一份</span>
            </div>
            <div class="reward-footer">
                * 未达标者工作照旧 *
            </div>
        </div>

        <div class="footer">
            👆 点击排名卡片跳转 · 🌓切换深色 · 📊下载统计 · 最近5次得分
        </div>
    </div>

    <!-- 下载提示 -->
    <div class="download-toast" id="downloadToast">
        <span class="toast-icon">✓</span>
        <span id="toastMessage">统计图已生成</span>
    </div>

    <script>
        // 检查系统主题偏好
        if (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) {
            document.body.classList.add('night-mode');
        }
        
        const searchInput = document.getElementById('search');
        const allRows = document.querySelectorAll('tbody tr');
        const downloadBtn = document.getElementById('downloadBtn');
        const saveChartBtn = document.getElementById('saveChartBtn');
        const chartCard = document.getElementById('chartCard');
        const closeChart = document.getElementById('closeChart');
        const downloadToast = document.getElementById('downloadToast');
        const toastMessage = document.getElementById('toastMessage');
        const statsGrid = document.getElementById('statsGrid');
        
        // 统计数据
        const statsData = [
'''

# 按固定顺序输出统计图数据
for g in ["星穹组", "夜曜组", "沧澜组"]:
    for stat in stats_data:
        if stat["group"] == g:
            html += f'''            {{ group: "{stat['group']}", total: {stat['total']}, rank: {stat['rank']}, members: {stat['members']}, avg: {stat['avg']}, color: "{stat['color']}" }},\n'''

html += '''        ];

        // 组数据
        const groups = ["星穹组", "夜曜组", "沧澜组"];
        const scores = [
            statsData.find(s => s.group === "星穹组").total,
            statsData.find(s => s.group === "夜曜组").total,
            statsData.find(s => s.group === "沧澜组").total
        ];
        const colors = ["#eab308", "#a855f7", "#3b82f6"];
        
        let chart = null;

        // 显示提示
        function showToast(message, isSuccess = true) {
            toastMessage.textContent = message;
            downloadToast.classList.add('show');
            setTimeout(() => {
                downloadToast.classList.remove('show');
            }, 2000);
        }

        // 生成统计图
        function generateChart() {
            const ctx = document.getElementById('groupChart').getContext('2d');
            const isNightMode = document.body.classList.contains('night-mode');
            const textColor = isNightMode ? '#94a3b8' : '#5a6b7a';
            const gridColor = isNightMode ? '#2d3a4d' : '#e1e8f0';
            
            if (chart) {
                chart.destroy();
            }
            
            chart = new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: groups,
                    datasets: [{
                        label: '总分',
                        data: scores,
                        backgroundColor: colors,
                        borderRadius: 6,
                        barPercentage: 0.6
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: { display: false },
                        tooltip: {
                            callbacks: {
                                label: function(context) {
                                    const value = context.raw;
                                    const group = groups[context.dataIndex];
                                    const stat = statsData.find(s => s.group === group);
                                    return [
                                        `总分: ${value}分`,
                                        `排名: 第${stat.rank}名`,
                                        `人数: ${stat.members}人`,
                                        `平均: ${stat.avg}分`
                                    ];
                                }
                            }
                        }
                    },
                    scales: {
                        y: {
                            beginAtZero: true,
                            grid: { color: gridColor },
                            ticks: { color: textColor }
                        },
                        x: {
                            grid: { display: false },
                            ticks: { color: textColor }
                        }
                    }
                }
            });
        }

        // 生成统计卡片
        function generateStatsGrid() {
            let html = '';
            statsData.forEach(stat => {
                html += `
                    <div class="stat-item">
                        <div class="stat-label">${stat.group}</div>
                        <div class="stat-value">${stat.total}</div>
                        <div class="stat-rank">第${stat.rank}名 · ${stat.members}人</div>
                    </div>
                `;
            });
            statsGrid.innerHTML = html;
        }

        // 保存统计图到相册
        async function saveChartToGallery() {
            const chartCard = document.getElementById('chartCard');
            
            try {
                showToast('📸 正在生成图片...');
                
                if (!chart) {
                    generateChart();
                }
                
                setTimeout(async () => {
                    const canvas = await html2canvas(chartCard, {
                        scale: 2,
                        backgroundColor: getComputedStyle(document.body).backgroundColor,
                        allowTaint: false,
                        useCORS: true,
                        logging: false
                    });
                    
                    const link = document.createElement('a');
                    link.download = `学长团统计_${new Date().toISOString().slice(0,10)}.png`;
                    link.href = canvas.toDataURL('image/png');
                    link.click();
                    
                    showToast('✅ 已保存到相册');
                }, 500);
                
            } catch (error) {
                console.error('保存失败:', error);
                showToast('❌ 保存失败');
            }
        }

        downloadBtn.addEventListener('click', () => {
            chartCard.classList.add('show');
            generateChart();
            generateStatsGrid();
            showToast('📊 统计图已生成');
        });

        if (saveChartBtn) {
            saveChartBtn.addEventListener('click', saveChartToGallery);
        }

        closeChart.addEventListener('click', () => {
            chartCard.classList.remove('show');
        });

        searchInput.addEventListener('input', (e) => {
            const searchTerm = e.target.value.toLowerCase().trim();
            allRows.forEach(row => {
                const searchText = row.getAttribute('data-search').toLowerCase();
                row.style.display = searchText.includes(searchTerm) ? '' : 'none';
            });
        });

        const observer = new MutationObserver(() => {
            if (chart && chartCard.classList.contains('show')) {
                generateChart();
            }
        });
        
        observer.observe(document.body, { attributes: true, attributeFilter: ['class'] });
    </script>
</body>
</html>'''

# 保存HTML文件
with open("index.html", "w", encoding="utf-8") as f:
    f.write(html)

print(f"\n✅ 生成成功！共 {len(people)} 人")
print("\n📊 奖励统计:")
for g in ["星穹组", "夜曜组", "沧澜组"]:
    if g in group_data:
        members = group_data[g]
        pass_count = sum(1 for m in members if m["reward_status"] == "✅")
        print(f"  {g}: {pass_count}/{len(members)} 人达标 ({int(pass_count/len(members)*100)}%)")
print("✨ 融合版：英文名完整显示 + 奖励机制卡片 + 统计图下载")
