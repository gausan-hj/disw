import pandas as pd
from datetime import datetime
import base64
import io
import json

# ===== 你要修改的地方 =====
SHEET_ID = "1YVa3nLUBW80j2nA4mudEqLH91RJ0FSRytmoDqmbyUJk"
SHEET_NAME = "Sheet3"
LANGUAGES_JSON_PATH = "languages.json"  # 语言文件路径
# ========================

# ===== 2026年联课活动数据（多语言）=====
cca_data = [
    {"date": "2026-02-24", "activity": "活动顾问+总学长+副总学长", "activity_en": "Advisor + Head Prefect + Vice Head", "activity_ms": "Penasihat + Ketua Pengawas + Timbalan", "uniform": "校服"},
    {"date": "2026-02-26", "activity": "操步", "activity_en": "Marching", "activity_ms": "Kawad", "uniform": "体育衣"},
    {"date": "2026-03-03", "activity": "操步", "activity_en": "Marching", "activity_ms": "Kawad", "uniform": "体育衣"},
    {"date": "2026-03-05", "activity": "活动（文书+财政+查账）", "activity_en": "Activity (Secretary + Treasurer + Auditor)", "activity_ms": "Aktiviti (Setiausaha + Bendahari + Juruaudit)", "uniform": "校服"},
    {"date": "2026-03-10", "activity": "操步", "activity_en": "Marching", "activity_ms": "Kawad", "uniform": "体育衣"},
    {"date": "2026-03-12", "activity": "活动（行动组）", "activity_en": "Activity (Action Group)", "activity_ms": "Aktiviti (Kumpulan Tindakan)", "uniform": "校服"},
    {"date": "2026-03-17", "activity": "操步", "activity_en": "Marching", "activity_ms": "Kawad", "uniform": "体育衣"},
    {"date": "2026-03-31", "activity": "操步", "activity_en": "Marching", "activity_ms": "Kawad", "uniform": "体育衣"},
    {"date": "2026-04-02", "activity": "活动（值岗组）", "activity_en": "Activity (Duty Group)", "activity_ms": "Aktiviti (Kumpulan Bertugas)", "uniform": "校服"},
    {"date": "2026-04-07", "activity": "操步", "activity_en": "Marching", "activity_ms": "Kawad", "uniform": "体育衣"},
    {"date": "2026-04-09", "activity": "活动（督察组）", "activity_en": "Activity (Supervisory Group)", "activity_ms": "Aktiviti (Kumpulan Penyelia)", "uniform": "校服"},
    {"date": "2026-04-14", "activity": "操步", "activity_en": "Marching", "activity_ms": "Kawad", "uniform": "体育衣"},
    {"date": "2026-04-16", "activity": "活动（星穹组）", "activity_en": "Activity (Star Group)", "activity_ms": "Aktiviti (Kumpulan Bintang)", "uniform": "校服"},
    {"date": "2026-04-21", "activity": "操步", "activity_en": "Marching", "activity_ms": "Kawad", "uniform": "体育衣"},
    {"date": "2026-04-23", "activity": "活动（夜曜组）", "activity_en": "Activity (Moon Group)", "activity_ms": "Aktiviti (Kumpulan Bulan)", "uniform": "校服"},
    {"date": "2026-04-28", "activity": "操步", "activity_en": "Marching", "activity_ms": "Kawad", "uniform": "体育衣"},
    {"date": "2026-04-30", "activity": "活动（沧澜组）", "activity_en": "Activity (Ocean Group)", "activity_ms": "Aktiviti (Kumpulan Lautan)", "uniform": "校服"},
    {"date": "2026-05-05", "activity": "第一学期检讨", "activity_en": "First Semester Review", "activity_ms": "Semakan Semester Pertama", "uniform": "校服"},
    {"date": "2026-05-07", "activity": "操步", "activity_en": "Marching", "activity_ms": "Kawad", "uniform": "体育衣"}
]

cca_json = json.dumps(cca_data, ensure_ascii=False)

# ===== 加载语言文件 =====
print(f"正在加载语言文件: {LANGUAGES_JSON_PATH}")
try:
    with open(LANGUAGES_JSON_PATH, 'r', encoding='utf-8') as f:
        languages = json.load(f)
    print("✅ 语言文件加载成功")
except FileNotFoundError:
    print(f"❌ 找不到语言文件: {LANGUAGES_JSON_PATH}")
    print("请确保 languages.json 在正确的路径")
    # 创建一个默认的语言文件
    languages = {
        "zh": {
            "app": {
                "title": "学长团分数板",
                "subtitle": "Prefects' Scoreboard",
                "search": "搜姓名/班级/学号...",
                "download": "下载统计",
                "dark": "深色",
                "footer": "👆 双击屏幕 / 按两次空格切换深色 · 📊下载统计 · 颜色越浅分数越高",
                "heatmap": {
                    "low": "低分",
                    "high": "高分"
                },
                "chart": {
                    "title": "各组总分对比",
                    "save": "保存到相册",
                    "close": "关闭",
                    "total": "总分",
                    "rank": "第{rank}名",
                    "members": "{count}人",
                    "average": "平均{avg}分"
                },
                "groups": {
                    "xingqiong": "星穹组",
                    "yeyao": "夜曜组",
                    "canglan": "沧澜组",
                    "average": "平均"
                },
                "table": {
                    "number": "#",
                    "name": "姓名",
                    "class": "班",
                    "id": "学号",
                    "daily": "每日得分",
                    "total": "总分",
                    "reward": "奖"
                },
                "reward": {
                    "title": "本轮奖励机制",
                    "subtitle": "公平原则 · 不负每一位付出的学长",
                    "first": {
                        "condition": "分数 ≥ 平均分÷2",
                        "benefit": "✨免搬椅子+减免操步"
                    },
                    "second": {
                        "condition": "分数 ≥ 平均分",
                        "benefit": "✨免搬椅子+减免操步"
                    },
                    "third": {
                        "condition": "组内前三名",
                        "benefit": "✨免搬椅子+减免操步"
                    },
                    "extra": "🎁 第1名达标者额外奖励一份",
                    "footer": "* 未达标者工作照旧 *"
                },
                "reminder": {
                    "button": "开启提醒",
                    "title": "每日提醒时间",
                    "confirm": "知道了，开启提醒",
                    "morning": "起床提醒",
                    "evening": "明天衣服提醒",
                    "night": "再次提醒",
                    "bedtime": "睡前提醒"
                },
                "uniform": {
                    "sports": "明天穿体育衣 - {activity}",
                    "uniform": "明天穿校服 - {activity}",
                    "default": "明天没有联课活动"
                },
                "toast": {
                    "generating": "📸 正在生成图片...",
                    "saved": "✅ 已保存到相册",
                    "saveFailed": "❌ 保存失败",
                    "chartGenerated": "📊 统计图已生成"
                }
            }
        },
        "en": {
            "app": {
                "title": "Prefects' Scoreboard",
                "subtitle": "Prefects' Scoreboard",
                "search": "Search name/class/id...",
                "download": "Download Chart",
                "dark": "Dark",
                "footer": "👆 Double tap / double space for dark mode · 📊Download chart · Lighter color = higher score",
                "heatmap": {
                    "low": "Low",
                    "high": "High"
                },
                "chart": {
                    "title": "Group Total Comparison",
                    "save": "Save to Gallery",
                    "close": "Close",
                    "total": "Total",
                    "rank": "Rank {rank}",
                    "members": "{count} members",
                    "average": "Avg {avg}"
                },
                "groups": {
                    "xingqiong": "Star Group",
                    "yeyao": "Moon Group",
                    "canglan": "Ocean Group",
                    "average": "Avg"
                },
                "table": {
                    "number": "#",
                    "name": "Name",
                    "class": "Class",
                    "id": "ID",
                    "daily": "Daily",
                    "total": "Total",
                    "reward": "Reward"
                },
                "reward": {
                    "title": "Reward System",
                    "subtitle": "Fairness · Every effort counts",
                    "first": {
                        "condition": "Score ≥ Avg ÷ 2",
                        "benefit": "✨No chair moving + Less marching"
                    },
                    "second": {
                        "condition": "Score ≥ Avg",
                        "benefit": "✨No chair moving + Less marching"
                    },
                    "third": {
                        "condition": "Top 3 in group",
                        "benefit": "✨No chair moving + Less marching"
                    },
                    "extra": "🎁 Extra reward for 1st place achievers",
                    "footer": "* Others continue duties *"
                },
                "reminder": {
                    "button": "Enable Reminders",
                    "title": "Daily Reminder Times",
                    "confirm": "Got it, enable",
                    "morning": "Wake up",
                    "evening": "Tomorrow's uniform",
                    "night": "Reminder again",
                    "bedtime": "Bedtime"
                },
                "uniform": {
                    "sports": "Tomorrow: Sports uniform - {activity}",
                    "uniform": "Tomorrow: School uniform - {activity}",
                    "default": "No CCA tomorrow"
                },
                "toast": {
                    "generating": "📸 Generating image...",
                    "saved": "✅ Saved to gallery",
                    "saveFailed": "❌ Save failed",
                    "chartGenerated": "📊 Chart generated"
                }
            }
        },
        "ms": {
            "app": {
                "title": "Papan Skor Pengawas",
                "subtitle": "Papan Skor Pengawas",
                "search": "Cari nama/kelas/id...",
                "download": "Muat Turun",
                "dark": "Gelap",
                "footer": "👆 Ketuk dua kali / ruang dua kali untuk mod gelap · 📊Muat turun · Warna cerah = skor tinggi",
                "heatmap": {
                    "low": "Rendah",
                    "high": "Tinggi"
                },
                "chart": {
                    "title": "Perbandingan Jumlah Kumpulan",
                    "save": "Simpan ke Galeri",
                    "close": "Tutup",
                    "total": "Jumlah",
                    "rank": "Kedudukan {rank}",
                    "members": "{count} ahli",
                    "average": "Purata {avg}"
                },
                "groups": {
                    "xingqiong": "Kumpulan Bintang",
                    "yeyao": "Kumpulan Bulan",
                    "canglan": "Kumpulan Lautan",
                    "average": "Purata"
                },
                "table": {
                    "number": "#",
                    "name": "Nama",
                    "class": "Kelas",
                    "id": "ID",
                    "daily": "Harian",
                    "total": "Jumlah",
                    "reward": "Ganjaran"
                },
                "reward": {
                    "title": "Sistem Ganjaran",
                    "subtitle": "Keadilan · Setiap usaha dihargai",
                    "first": {
                        "condition": "Skor ≥ Purata ÷ 2",
                        "benefit": "✨Tiada kerusi + Kurang kawad"
                    },
                    "second": {
                        "condition": "Skor ≥ Purata",
                        "benefit": "✨Tiada kerusi + Kurang kawad"
                    },
                    "third": {
                        "condition": "3 teratas dalam kumpulan",
                        "benefit": "✨Tiada kerusi + Kurang kawad"
                    },
                    "extra": "🎁 Ganjaran tambahan untuk tempat pertama",
                    "footer": "* Yang lain terus bertugas *"
                },
                "reminder": {
                    "button": "Aktifkan Peringatan",
                    "title": "Masa Peringatan Harian",
                    "confirm": "Faham, aktifkan",
                    "morning": "Bangun tidur",
                    "evening": "Pakaian esok",
                    "night": "Peringatan lagi",
                    "bedtime": "Waktu tidur"
                },
                "uniform": {
                    "sports": "Esok: Pakaian sukan - {activity}",
                    "uniform": "Esok: Pakaian sekolah - {activity}",
                    "default": "Tiada CCA esok"
                },
                "toast": {
                    "generating": "📸 Menjana imej...",
                    "saved": "✅ Disimpan ke galeri",
                    "saveFailed": "❌ Gagal menyimpan",
                    "chartGenerated": "📊 Carta dijana"
                }
            }
        }
    }
    print("✅ 使用默认语言配置")

languages_json = json.dumps(languages, ensure_ascii=False)

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

# 计算每组最高分和最低分，用于热力图
group_max_scores = {}
group_min_scores = {}
for g in group_data:
    if group_data[g]:
        group_max_scores[g] = max(m["total"] for m in group_data[g])
        group_min_scores[g] = min(m["total"] for m in group_data[g])
    else:
        group_max_scores[g] = 0
        group_min_scores[g] = 0

# 生成HTML - 添加OneSignal通知权限
html = '''<!DOCTYPE html>
<html lang="zh">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=yes, viewport-fit=cover">
    <title>训育处 - 学长团分数板</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
    
    <!-- OneSignal SDK - 用于通知权限 -->
    <script src="https://cdn.onesignal.com/sdks/web/v16/OneSignalSDK.page.js" defer></script>
    
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
            
            --heat-1: #083344;
            --heat-2: #164e63;
            --heat-3: #155e75;
            --heat-4: #0369a1;
            --heat-5: #0284c7;
            --heat-6: #38bdf8;
            --heat-7: #7dd3fc;
            --heat-8: #bae6fd;
            --heat-9: #e0f2fe;
            
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
            
            --heat-1: #000814;
            --heat-2: #001d3d;
            --heat-3: #003566;
            --heat-4: #0a4a7a;
            --heat-5: #1a5a8a;
            --heat-6: #2a6a9a;
            --heat-7: #3a7aaa;
            --heat-8: #4a8aba;
            --heat-9: #5a9aca;
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

        .double-tap-hint {
            position: fixed;
            bottom: 120px;
            left: 50%;
            transform: translateX(-50%);
            background: var(--card-bg);
            color: var(--text-primary);
            padding: 8px 16px;
            border-radius: 40px;
            font-size: 0.7rem;
            box-shadow: var(--shadow-lg);
            border: 1px solid var(--border-light);
            z-index: 9999;
            opacity: 0;
            transition: opacity 0.3s ease;
            pointer-events: none;
            white-space: nowrap;
        }

        .double-tap-hint.show {
            opacity: 0.9;
        }

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

        .lang-toggle {
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
            transition: all 0.2s ease;
            white-space: nowrap;
        }

        .lang-toggle:active {
            transform: scale(0.92);
        }

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

        .heatmap-legend {
            display: flex;
            align-items: center;
            justify-content: flex-end;
            gap: 12px;
            margin-bottom: 8px;
            font-size: 0.65rem;
            color: var(--text-tertiary);
        }

        .legend-colors {
            display: flex;
            gap: 2px;
        }

        .legend-color {
            width: 16px;
            height: 16px;
            border-radius: 4px;
        }

        .legend-color.low { background: var(--heat-1); }
        .legend-color.mid-low { background: var(--heat-3); }
        .legend-color.mid { background: var(--heat-5); }
        .legend-color.mid-high { background: var(--heat-7); }
        .legend-color.high { background: var(--heat-9); }

        .legend-label {
            display: flex;
            gap: 8px;
        }

        .legend-label span.low { color: var(--heat-1); font-weight: bold; }
        .legend-label span.high { color: var(--heat-9); font-weight: bold; }

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

        body.night-mode .save-chart-btn {
            color: #ffffff;
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

        .table-container {
            overflow-x: auto;
            -webkit-overflow-scrolling: touch;
            margin: 0 -4px;
            padding: 0 4px;
        }

        .member-table {
            width: 100%;
            border-collapse: collapse;
            min-width: 750px;
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
        .member-table th:nth-child(2) { width: 180px; }
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
            min-height: 45px;
            position: relative;
        }

        .name-cn, .name-en {
            transition: all 0.3s ease;
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
            white-space: normal;
            line-height: 1.3;
            word-break: break-word;
        }

        .name-cell.lang-switching .name-cn {
            transform: translateY(100%);
            opacity: 0;
        }

        .name-cell.lang-switching .name-en {
            transform: translateY(-100%);
            opacity: 0;
        }

        body.lang-zh .name-cell {
            display: flex;
            flex-direction: column;
        }

        body.lang-zh .name-cell .name-cn {
            order: 1;
            transform: translateY(0);
            opacity: 1;
        }

        body.lang-zh .name-cell .name-en {
            order: 2;
            transform: translateY(0);
            opacity: 0.8;
        }

        body.lang-en .name-cell {
            display: flex;
            flex-direction: column;
        }

        body.lang-en .name-cell .name-en {
            order: 1;
            font-weight: 600;
            font-size: 0.8rem;
            color: var(--text-primary);
            opacity: 1;
            transform: translateY(0);
        }

        body.lang-en .name-cell .name-cn {
            order: 2;
            font-size: 0.65rem;
            color: var(--text-tertiary);
            font-weight: 400;
            opacity: 0.8;
            transform: translateY(0);
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        body.lang-ms .name-cell {
            display: flex;
            flex-direction: column;
        }

        body.lang-ms .name-cell .name-en {
            order: 1;
            font-weight: 600;
            font-size: 0.8rem;
            color: var(--text-primary);
            opacity: 1;
            transform: translateY(0);
        }

        body.lang-ms .name-cell .name-cn {
            order: 2;
            font-size: 0.65rem;
            color: var(--text-tertiary);
            font-weight: 400;
            opacity: 0.8;
            transform: translateY(0);
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .info-cell {
            font-size: 0.7rem;
            color: var(--text-secondary);
        }

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
            transition: transform 0.1s ease;
        }

        .score-item:hover {
            transform: scale(1.1);
            z-index: 10;
        }

        .score-item.has-score {
            background: var(--score-highlight);
            color: var(--score-highlight-text);
        }

        .score-date { opacity: 0.7; margin-right: 1px; }
        .score-value { font-weight: 600; }

        .total-heat {
            font-weight: 700;
            font-size: 0.9rem;
            padding: 2px 6px;
            border-radius: 12px;
            display: inline-block;
            min-width: 35px;
            text-align: center;
            color: #ffffff !important;
            text-shadow: 0 1px 2px rgba(0,0,0,0.3);
        }

        body.night-mode .total-heat {
            color: #ffffff !important;
            text-shadow: 0 1px 3px rgba(0,0,0,0.5);
        }

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
            z-index: 2000;
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

        body.night-mode .toast-icon {
            color: #ffffff;
        }

        .notification-toast {
            position: fixed;
            top: 20px;
            right: 20px;
            background: var(--card-bg);
            border-radius: 16px;
            padding: 16px;
            box-shadow: var(--shadow-lg);
            border: 1px solid var(--border-light);
            z-index: 3000;
            max-width: 280px;
            display: none;
            animation: slideIn 0.3s ease;
        }

        .notification-toast.show {
            display: block;
        }

        @keyframes slideIn {
            from {
                opacity: 0;
                transform: translateX(20px);
            }
            to {
                opacity: 1;
                transform: translateX(0);
            }
        }

        .toast-close {
            position: absolute;
            top: 8px;
            right: 8px;
            cursor: pointer;
            color: var(--text-tertiary);
            font-size: 0.8rem;
            padding: 4px;
            border-radius: 50%;
            width: 20px;
            height: 20px;
            display: flex;
            align-items: center;
            justify-content: center;
        }

        .toast-close:hover {
            background: var(--border-subtle);
        }

        .toast-title {
            display: flex;
            align-items: center;
            gap: 8px;
            margin-bottom: 8px;
            font-weight: 600;
        }

        .reminder-btn {
            position: fixed;
            bottom: 20px;
            right: 20px;
            background: var(--star-primary);
            color: #1a2b3c;
            padding: 14px 24px;
            border-radius: 60px;
            display: flex;
            align-items: center;
            gap: 12px;
            box-shadow: 0 4px 15px rgba(234, 179, 8, 0.3);
            cursor: pointer;
            font-weight: 600;
            border: 2px solid var(--border-light);
            z-index: 1000;
            font-size: 1rem;
            transition: all 0.3s ease;
            animation: pulse 2s infinite;
        }

        .reminder-btn:hover {
            transform: scale(1.05);
            box-shadow: 0 6px 20px rgba(234, 179, 8, 0.4);
        }

        .reminder-btn:active {
            transform: scale(0.95);
        }

        @keyframes pulse {
            0% {
                box-shadow: 0 4px 15px rgba(234, 179, 8, 0.3);
            }
            50% {
                box-shadow: 0 8px 25px rgba(234, 179, 8, 0.6);
            }
            100% {
                box-shadow: 0 4px 15px rgba(234, 179, 8, 0.3);
            }
        }

        .bell-icon {
            font-size: 1.4rem;
            animation: ring 3s infinite;
        }

        @keyframes ring {
            0%, 100% { transform: rotate(0); }
            10%, 30%, 50%, 70%, 90% { transform: rotate(10deg); }
            20%, 40%, 60%, 80% { transform: rotate(-10deg); }
        }

        .reminder-popup {
            position: fixed;
            bottom: 100px;
            right: 20px;
            background: var(--card-bg);
            border-radius: 24px;
            box-shadow: var(--shadow-lg);
            border: 1px solid var(--border-light);
            width: 300px;
            z-index: 2000;
            display: none;
            animation: slideUp 0.4s ease;
        }

        .reminder-popup.show {
            display: block;
        }

        @keyframes slideUp {
            from {
                opacity: 0;
                transform: translateY(20px);
            }
            to {
                opacity: 1;
                transform: translateY(0);
            }
        }

        .popup-header {
            padding: 16px;
            border-bottom: 1px solid var(--border-subtle);
            display: flex;
            align-items: center;
            gap: 8px;
            position: relative;
        }

        .popup-icon {
            font-size: 1.4rem;
        }

        .popup-title {
            font-size: 1.1rem;
            font-weight: 600;
            color: var(--text-primary);
        }

        .popup-close {
            position: absolute;
            top: 16px;
            right: 16px;
            cursor: pointer;
            font-size: 1rem;
            color: var(--text-tertiary);
            padding: 4px;
            border-radius: 50%;
            width: 24px;
            height: 24px;
            display: flex;
            align-items: center;
            justify-content: center;
            transition: background 0.2s;
        }

        .popup-close:hover {
            background: var(--border-subtle);
        }

        .popup-content {
            padding: 16px;
        }

        .time-item {
            display: flex;
            align-items: center;
            padding: 12px 0;
            border-bottom: 1px solid var(--border-subtle);
        }

        .time-item:last-child {
            border-bottom: none;
        }

        .time-icon {
            font-size: 1.2rem;
            width: 32px;
        }

        .time-label {
            font-size: 1rem;
            font-weight: 600;
            color: var(--text-primary);
            width: 80px;
        }

        .time-desc {
            font-size: 0.85rem;
            color: var(--text-tertiary);
        }

        .popup-footer {
            padding: 16px;
            border-top: 1px solid var(--border-subtle);
        }

        .popup-btn {
            width: 100%;
            background: var(--star-primary);
            color: #1a2b3c;
            border: none;
            border-radius: 30px;
            padding: 12px;
            font-size: 1rem;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
        }

        .popup-btn:hover {
            transform: scale(1.02);
            background: #ca8a04;
        }

        .popup-btn:active {
            transform: scale(0.98);
        }

        body.night-mode .popup-btn {
            color: #ffffff;
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
<body class="lang-zh">
    <div class="container">
        <div class="double-tap-hint" id="doubleTapHint">
            <span id="hintText">👆 双击屏幕 / 按两次空格 切换深色模式</span>
        </div>

        <div class="header">
            <div class="header-top">
                <div class="title-group">
                    <span class="school-icon">🏫</span>
                    <h1>学长团分数板 · 热力图</h1>
                </div>
                <div class="action-buttons">
                    <button class="download-btn" id="downloadBtn">
                        <span>📊</span>
                        <span>下载统计</span>
                    </button>
                    <div class="lang-toggle" id="langToggle">
                        <span class="lang-left">中</span>
                        <span class="lang-separator">/</span>
                        <span class="lang-right">EN</span>
                    </div>
                    <div class="theme-toggle" id="themeToggle">
                        <span class="moon-icon">🌓</span>
                        <span>深色</span>
                    </div>
                </div>
            </div>
            <div class="meta-info">
                <span>Prefects' Scoreboard · 颜色越浅分数越高</span>
                <span class="date-badge">{datetime.now().strftime('%m/%d %H:%M')}</span>
            </div>
            <div class="search-wrapper">
                <span class="search-icon">🔍</span>
                <input type="text" id="search" placeholder="搜姓名/班级/学号...">
            </div>
        </div>

        <div class="heatmap-legend">
            <div class="legend-label">
                <span class="low">低分 █</span>
                <span> → </span>
                <span class="high">高分 █</span>
            </div>
            <div class="legend-colors">
                <div class="legend-color low" title="低分（深色）"></div>
                <div class="legend-color mid-low" title="中低分"></div>
                <div class="legend-color mid" title="中分"></div>
                <div class="legend-color mid-high" title="中高分"></div>
                <div class="legend-color high" title="高分（浅色）"></div>
            </div>
        </div>

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

        <div class="rank-grid">
            <!-- 组排名卡片区域 -->
        </div>

        <div class="groups">
            <!-- 组别详情区域 -->
        </div>

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

        <div class="reminder-btn" id="reminderBtn">
            <span class="bell-icon">🔔</span>
            <span class="reminder-text">开启提醒</span>
        </div>

        <div class="reminder-popup" id="reminderPopup">
            <div class="popup-header">
                <span class="popup-icon">⏰</span>
                <span class="popup-title">每日提醒时间</span>
                <span class="popup-close" onclick="closePopup()">✕</span>
            </div>
            <div class="popup-content">
                <div class="time-item">
                    <span class="time-icon">🌅</span>
                    <span class="time-label">早上 6:00</span>
                    <span class="time-desc">起床提醒</span>
                </div>
                <div class="time-item">
                    <span class="time-icon">🌆</span>
                    <span class="time-label">晚上 7:00</span>
                    <span class="time-desc">明天衣服提醒</span>
                </div>
                <div class="time-item">
                    <span class="time-icon">🌙</span>
                    <span class="time-label">晚上 8:15</span>
                    <span class="time-desc">再次提醒</span>
                </div>
                <div class="time-item">
                    <span class="time-icon">🌃</span>
                    <span class="time-label">晚上 10:00</span>
                    <span class="time-desc">睡前提醒</span>
                </div>
            </div>
            <div class="popup-footer">
                <button class="popup-btn" onclick="enableReminders()">知道了，开启提醒</button>
            </div>
        </div>

        <div class="notification-toast" id="notificationToast">
            <div class="toast-close" onclick="this.parentElement.classList.remove('show')">✕</div>
            <div class="toast-title">
                <span>🔔</span>
                <span id="toastMessage">通知</span>
            </div>
            <div style="font-size: 0.8rem; color: var(--text-secondary);" id="toastDetail"></div>
        </div>

        <div class="footer">
            👆 双击屏幕 / 按两次空格切换深色 · 📊下载统计 · 颜色越浅分数越高
        </div>
    </div>

    <div class="download-toast" id="downloadToast">
        <span class="toast-icon">✓</span>
        <span id="toastMessage">统计图已生成</span>
    </div>

    <script>
        // ========== OneSignal 初始化 - 用于通知权限 ==========
        window.OneSignalDeferred = window.OneSignalDeferred || [];
        OneSignalDeferred.push(async function(OneSignal) {
            await OneSignal.init({
                appId: "137d66ce-9746-4206-9861-a1c368a0b548", // OneSignal App ID
                allowLocalhostAsSecureOrigin: true,
                notifyButton: {
                    enable: true,
                    size: 'medium',
                    showCredit: false,
                    position: 'bottom-right',
                    text: {
                        'message': '开启提醒',
                        'subscribe': '知道了，开启提醒',
                        'unsubscribe': '关闭提醒'
                    },
                    colors: {
                        'circle.background': '#eab308',
                        'circle.foreground': '#1a2b3c',
                        'badge.background': '#eab308',
                        'badge.foreground': '#1a2b3c',
                        'badge.bordercolor': '#e2e8f0',
                        'button.background': '#eab308',
                        'button.foreground': '#1a2b3c',
                        'button.hover.background': '#ca8a04'
                    }
                },
                serviceWorkerParam: { scope: '/' },
                serviceWorkerPath: 'OneSignalSDKWorker.js',
                serviceWorkerUpdaterPath: 'OneSignalSDKUpdaterWorker.js',
            });
            console.log('OneSignal 初始化成功');
        });

        // ========== 首先定义所有函数 ==========
        
        // 切换深色模式
        function toggleNightMode() {
            document.body.classList.toggle('night-mode');
            showHint(document.body.classList.contains('night-mode') ? '🌙 深色模式' : '☀️ 日间模式');
            
            // 如果图表显示中，重新生成以适配深色模式
            const chartCard = document.getElementById('chartCard');
            if (chartCard && chartCard.classList.contains('show') && window.chart) {
                generateChart();
            }
        }

        // 显示提示
        function showHint(message) {
            const hint = document.getElementById('doubleTapHint');
            const hintText = document.getElementById('hintText');
            if (!hint || !hintText) return;
            
            hintText.textContent = message;
            hint.classList.add('show');
            
            clearTimeout(window.hintTimeout);
            window.hintTimeout = setTimeout(() => {
                hint.classList.remove('show');
            }, 1500);
        }

        // 显示下载提示
        function showToast(message, isSuccess = true) {
            const toast = document.getElementById('downloadToast');
            const msgEl = document.getElementById('toastMessage');
            if (!toast || !msgEl) return;
            
            msgEl.textContent = message;
            toast.classList.add('show');
            
            setTimeout(() => {
                toast.classList.remove('show');
            }, 2000);
        }

        // 显示通知提示
        function showNotification(title, detail) {
            const toast = document.getElementById('notificationToast');
            const titleEl = document.getElementById('toastMessage');
            const detailEl = document.getElementById('toastDetail');
            
            if (!toast || !titleEl || !detailEl) return;
            
            titleEl.textContent = title;
            detailEl.textContent = detail;
            toast.classList.add('show');
            
            setTimeout(() => {
                toast.classList.remove('show');
            }, 3000);
        }

        // 开启提醒 - 使用 OneSignal 请求权限
        function enableReminders() {
            OneSignalDeferred.push(function(OneSignal) {
                OneSignal.Notifications.requestPermission();
            });
            showNotification('🔔 提醒已开启', '你将在以下时间收到通知：\\n早上6:00 · 晚上7:00 · 晚上8:15 · 晚上10:00');
            closePopup();
        }

        // 显示提醒弹窗
        function showReminderPopup() {
            const popup = document.getElementById('reminderPopup');
            if (popup) popup.classList.add('show');
        }

        // 关闭弹窗
        function closePopup() {
            const popup = document.getElementById('reminderPopup');
            if (popup) popup.classList.remove('show');
        }

        // 生成统计图
        function generateChart() {
            const canvas = document.getElementById('groupChart');
            if (!canvas) return;
            
            const ctx = canvas.getContext('2d');
            const isNightMode = document.body.classList.contains('night-mode');
            const textColor = isNightMode ? '#94a3b8' : '#5a6b7a';
            const gridColor = isNightMode ? '#2d3a4d' : '#e1e8f0';
            
            // 数据 - 会被Python替换
            const groups = ['星穹组', '夜曜组', '沧澜组'];
            const scores = [1234, 1156, 1089];
            const colors = ['#eab308', '#a855f7', '#3b82f6'];
            
            if (window.chart) {
                window.chart.destroy();
            }
            
            window.chart = new Chart(ctx, {
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
                                    return `总分: ${value}`;
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
            
            // 更新统计卡片
            const statsGrid = document.getElementById('statsGrid');
            if (statsGrid) {
                statsGrid.innerHTML = groups.map((g, i) => `
                    <div class="stat-item">
                        <div class="stat-label">${g}</div>
                        <div class="stat-value">${scores[i]}</div>
                        <div class="stat-rank">第${i+1}名 · 10人</div>
                    </div>
                `).join('');
            }
        }

        // 保存图表
        async function saveChartToGallery() {
            const chartCard = document.getElementById('chartCard');
            if (!chartCard) return;
            
            try {
                showToast('📸 正在生成图片...');
                
                if (!window.chart) {
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
                    link.download = `prefects_score_${new Date().toISOString().slice(0,10)}.png`;
                    link.href = canvas.toDataURL('image/png');
                    link.click();
                    
                    showToast('✅ 已保存到相册');
                }, 500);
                
            } catch (error) {
                console.error('保存失败:', error);
                showToast('❌ 保存失败');
            }
        }

        // ========== DOM 加载完成后绑定事件 ==========
        document.addEventListener('DOMContentLoaded', function() {
            
            // 深色模式切换
            const themeToggle = document.getElementById('themeToggle');
            if (themeToggle) {
                themeToggle.addEventListener('click', function(e) {
                    toggleNightMode();
                });
            }

            // 语言切换
            const langToggle = document.getElementById('langToggle');
            if (langToggle) {
                langToggle.addEventListener('click', function(e) {
                    const body = document.body;
                    if (body.classList.contains('lang-zh')) {
                        body.classList.remove('lang-zh');
                        body.classList.add('lang-en');
                        this.querySelector('.lang-left').textContent = 'EN';
                        this.querySelector('.lang-right').textContent = 'MS';
                    } else if (body.classList.contains('lang-en')) {
                        body.classList.remove('lang-en');
                        body.classList.add('lang-ms');
                        this.querySelector('.lang-left').textContent = 'MS';
                        this.querySelector('.lang-right').textContent = '中';
                    } else {
                        body.classList.remove('lang-ms');
                        body.classList.add('lang-zh');
                        this.querySelector('.lang-left').textContent = '中';
                        this.querySelector('.lang-right').textContent = 'EN';
                    }
                    showNotification('🌐', `已切换到 ${body.classList.contains('lang-zh') ? '中文' : body.classList.contains('lang-en') ? 'English' : 'Bahasa Melayu'}`);
                });
            }

            // 下载统计按钮
            const downloadBtn = document.getElementById('downloadBtn');
            const chartCard = document.getElementById('chartCard');
            const closeChart = document.getElementById('closeChart');
            
            if (downloadBtn && chartCard) {
                downloadBtn.addEventListener('click', function(e) {
                    chartCard.classList.add('show');
                    generateChart();
                    showToast('📊 统计图已生成');
                });
            }

            // 保存图表按钮
            const saveChartBtn = document.getElementById('saveChartBtn');
            if (saveChartBtn) {
                saveChartBtn.addEventListener('click', function(e) {
                    saveChartToGallery();
                });
            }

            // 关闭图表
            if (closeChart && chartCard) {
                closeChart.addEventListener('click', function(e) {
                    chartCard.classList.remove('show');
                });
            }

            // 搜索功能
            const searchInput = document.getElementById('search');
            if (searchInput) {
                searchInput.addEventListener('input', function(e) {
                    const searchTerm = e.target.value.toLowerCase().trim();
                    const allRows = document.querySelectorAll('tbody tr');
                    
                    allRows.forEach(row => {
                        const searchText = row.getAttribute('data-search')?.toLowerCase() || '';
                        row.style.display = searchText.includes(searchTerm) ? '' : 'none';
                    });
                });
            }

            // 提醒按钮
            const reminderBtn = document.getElementById('reminderBtn');
            if (reminderBtn) {
                reminderBtn.addEventListener('click', function(e) {
                    showReminderPopup();
                });
            }

            // 点击外部关闭弹窗
            document.addEventListener('click', function(e) {
                const popup = document.getElementById('reminderPopup');
                const btn = document.getElementById('reminderBtn');
                if (popup && btn) {
                    if (!popup.contains(e.target) && !btn.contains(e.target)) {
                        popup.classList.remove('show');
                    }
                }
            });

            // 双击/双空格监听
            let lastTap = 0;
            let lastSpaceTime = 0;
            let spaceCount = 0;

            document.addEventListener('touchstart', (e) => {
                const currentTime = new Date().getTime();
                const tapLength = currentTime - lastTap;
                
                if (tapLength < 200 && tapLength > 0) {
                    toggleNightMode();
                    e.preventDefault();
                }
                lastTap = currentTime;
            });

            document.addEventListener('keydown', (e) => {
                if (e.code === 'Space') {
                    e.preventDefault();
                    
                    const currentTime = new Date().getTime();
                    
                    if (currentTime - lastSpaceTime < 500) {
                        spaceCount++;
                        if (spaceCount === 2) {
                            toggleNightMode();
                            spaceCount = 0;
                        }
                    } else {
                        spaceCount = 1;
                    }
                    
                    lastSpaceTime = currentTime;
                }
            });
        });

        // 检查系统主题偏好
        if (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) {
            document.body.classList.add('night-mode');
        }
    </script>
</body>
</html>'''

# 添加联课活动数据和语言数据
html = html.replace("{languages_json}", languages_json)
html = html.replace("{cca_json}", cca_json)

# 添加组排名卡片
rank_icons = {1: "🥇", 2: "🥈", 3: "🥉"}
group_ids = {"星穹组": "group-xingqiong", "夜曜组": "group-yeyao", "沧澜组": "group-canglan"}
group_list = []
total_list = []

rank_cards_html = ""
for i, (g, total) in enumerate(sorted_groups, 1):
    group_id = group_ids[g]
    group_list.append(g)
    total_list.append(int(total))
    
    rank_cards_html += f'''
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

# 添加组别详情
group_emojis = {"星穹组": "✨", "夜曜组": "🌙", "沧澜组": "🌊"}
groups_html = ""
for group_name in ["星穹组", "夜曜组", "沧澜组"]:
    members = group_data[group_name]
    rank = group_rank[group_name]
    group_id = group_ids[group_name]
    avg_score = group_averages[group_name]
    
    # 获取该组的最高分和最低分
    group_max = group_max_scores[group_name]
    group_min = group_min_scores[group_name]
    
    group_html = f'''
            <div class="group-card" data-group="{group_name}" id="{group_id}">
                <div class="group-header">
                    <div class="group-title-wrapper">
                        <span class="group-emoji">{group_emojis[group_name]}</span>
                        <span class="group-title">{group_name}</span>
                    </div>
                    <div class="group-stats">
                        <span class="group-avg">平均 {int(avg_score)}</span>
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
        
        # 完整显示英文名
        name_en_full = member['name_en']
        
        # 计算热力图颜色 - 越浅越高分
        total_score = member['total']
        if group_max > group_min:
            relative_score = (total_score - group_min) / (group_max - group_min)
            heat_level = min(9, max(1, int(relative_score * 9) + 1))
        else:
            heat_level = 5
        
        total_cell = f'<span class="total-heat" style="background-color: var(--heat-{heat_level});">{int(total_score)}</span>'
        
        group_html += f'''
                        <tr data-search="{member['name_cn']} {member['name_en']} {member['class']} {member['student_id']}">
                            <td>{member['order']}</td>
                            <td class="name-cell">
                                <div class="name-cn">{member['name_cn']}</div>
                                <div class="name-en">{name_en_full}</div>
                            </td>
                            <td class="info-cell">{member['class']}</td>
                            <td class="info-cell">{member['student_id']}</td>
                            <td><div class="score-tags">{score_tags}</div></td>
                            <td>{total_cell}</td>
                            <td><span class="{member['reward_class']}">{member['reward_status']}</span></td>
                        </tr>
'''
    group_html += '''
                        </tbody>
                    </table>
                </div>
            </div>
'''
    groups_html += group_html

# 添加统计数据JSON
stats_json = json.dumps(stats_data, ensure_ascii=False)

# 替换HTML中的占位符
html = html.replace('<!-- 组排名卡片区域 -->', rank_cards_html)
html = html.replace('<!-- 组别详情区域 -->', groups_html)

# 替换JavaScript中的数据
html = html.replace('const groups = [\'星穹组\', \'夜曜组\', \'沧澜组\'];', f'const groups = {json.dumps(group_list)};')
html = html.replace('const scores = [1234, 1156, 1089];', f'const scores = {json.dumps(total_list)};')

# 替换日期
html = html.replace('{datetime.now().strftime(\'%m/%d %H:%M\')}', datetime.now().strftime('%m/%d %H:%M'))

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
