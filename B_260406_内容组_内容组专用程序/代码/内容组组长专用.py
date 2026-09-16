import os
import re
import json
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
import random

# ====================== 系统配置 ======================
MASTER_KEY = "66889900"
USER_FILE = "member_list.json"
SETTING_FILE = "settings.json"
CHECK_FOLDER = "待检查文件"
DRAMALIST_FILE = "广播剧列表.xlsx"
FINAL_TRANSFER_FILE = "可移交技术组列表.txt"
MONITOR_REPORT_FILE = "组长监控综合报告.txt"  # 组长可视化报告
# ======================================================

# 解决中文显示问题
plt.rcParams["font.sans-serif"] = ["SimHei", "Microsoft YaHei"]
plt.rcParams["axes.unicode_minus"] = False

# ------------------------------
# 组员ID管理（不变）
# ------------------------------
def generate_member_list():
    print("\n==== 组长：生成组员8位ID ====")
    names = input("输入组员姓名（空格分隔）：").strip().split()
    data = {''.join(random.choices("0123456789", k=8)): name for name in names}
    with open(USER_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    for uid, name in data.items():
        print(f"{uid} -> {name}")
    input("\n回车返回")

def get_name(uid):
    if not os.path.exists(USER_FILE):
        return None
    with open(USER_FILE, "r", encoding="utf-8") as f:
        return json.load(f).get(uid)

# ------------------------------
# DDL管理（不变，新增整体DDL统计）
# ------------------------------
def load_settings():
    if not os.path.exists(SETTING_FILE):
        return {"ddl": {}}
    with open(SETTING_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def save_settings(s):
    with open(SETTING_FILE, "w", encoding="utf-8") as f:
        json.dump(s, f, ensure_ascii=False, indent=2)

def set_ddl():
    s = load_settings()
    print("\n==== 设置DDL ====")
    drama = input("广播剧名称：")
    ddl_str = input("DDL(YYYY-MM-DD HH:MM)：")
    try:
        datetime.strptime(ddl_str, "%Y-%m-%d %H:%M")
        s["ddl"][drama] = ddl_str
        save_settings(s)
        print("设置成功")
    except:
        print("格式错误（正确格式：YYYY-MM-DD HH:MM）")
    input("返回")

# ------------------------------
# 广播剧&剧集读取（不变）
# ------------------------------
def get_all_dramas():
    if not os.path.exists(DRAMALIST_FILE):
        return []
    return list(pd.read_excel(DRAMALIST_FILE).columns)

def get_ep_list(drama):
    if not os.path.exists(DRAMALIST_FILE):
        return []
    df = pd.read_excel(DRAMALIST_FILE)
    return [str(x).strip() for x in df[drama] if pd.notna(x)] if drama in df.columns else []

def is_standard(filename):
    return re.match(r"^.+_.+_\d+_V\d+\.txt$", filename) is not None

# ------------------------------
# 组员：自动命名（不变）
# ------------------------------
def member_tool():
    print("\n==== 组员：文件规范化 ====")
    uid = input("输入8位ID：").strip()
    name = get_name(uid)
    if not name:
        print("ID无效")
        input("返回")
        return

    print(f"用户：{name}")
    dramas = get_all_dramas()
    if not dramas:
        print("无广播剧列表")
        input("返回")
        return

    print("可选择：", dramas)
    drama = input("广播剧：")
    eps = get_ep_list(drama)
    if not eps:
        print("无剧集")
        input("返回")
        return

    os.makedirs(CHECK_FOLDER, exist_ok=True)
    txts = [f for f in os.listdir(CHECK_FOLDER) if f.endswith(".txt") and not is_standard(f)]
    if not txts:
        print("无待规范文件")
        input("返回")
        return

    fname = txts[0]
    base = os.path.splitext(fname)[0]
    match = None
    for e in eps:
        a = base.lower().replace(" ","").replace("，","")
        b = e.lower().replace(" ","").replace("，","")
        if a in b or b in a:
            match = e
            break
    if not match:
        print("匹配失败")
        input("返回")
        return

    today = datetime.now().strftime("%Y%m%d")
    ver = input("版本号：")
    new = f"{match}_{name}_{today}_V{ver}.txt"
    os.rename(os.path.join(CHECK_FOLDER, fname), os.path.join(CHECK_FOLDER, new))
    print(f"已规范化：{new}")
    input("返回主页")

# ------------------------------
# 【核心升级】组长：可视化综合报告（含所有你要的指标）
# ------------------------------
def run_monitor():
    dramas = get_all_dramas()
    members = list(json.load(open(USER_FILE,"r",encoding="utf-8")).values()) if os.path.exists(USER_FILE) else []
    now = datetime.now()
    settings = load_settings()
    all_ddl = settings.get("ddl", {})

    # 基础数据统计
    os.makedirs(CHECK_FOLDER, exist_ok=True)
    all_files = [f for f in os.listdir(CHECK_FOLDER) if f.endswith(".txt")]
    std_files = [f for f in all_files if is_standard(f)]  # 规范命名的文件（全部可参与清洗）
    
    # 1. 个人提交数量统计
    user_submit = {u: sum(1 for f in std_files if f"_{u}_" in f) for u in members}
    
    # 2. 整体进度统计（总应提交数 = 所有广播剧的剧集数总和；已提交数 = 规范文件数）
    total_ep_count = sum(len(get_ep_list(drama)) for drama in dramas)
    finished_ep_count = len(std_files)
    overall_progress = round(finished_ep_count / total_ep_count * 100, 1) if total_ep_count > 0 else 0

    # 3. 准确率统计（个人+整体，只看内容，基于已清洗的文件）
    # 先统计所有已清洗、无错误的文件（读取所有清洗报告）
    cleaned_correct = 0
    cleaned_total = 0
    user_correct = {u: 0 for u in members}
    user_cleaned_total = {u: 0 for u in members}

    # 读取所有广播剧的清洗报告，统计准确率
    for drama in dramas:
        clean_report = f"【{drama}】全部通过.txt"
        error_excel = f"【{drama}】错误处理表.xlsx"
        
        # 全部通过的情况
        if os.path.exists(clean_report):
            drama_files = [f for f in std_files if any(f.startswith(e+"_") for e in get_ep_list(drama))]
            cleaned_total += len(drama_files)
            cleaned_correct += len(drama_files)
            # 分配到个人
            for f in drama_files:
                for u in members:
                    if f"_{u}_" in f:
                        user_correct[u] += 1
                        user_cleaned_total[u] += 1
        # 有错误但已全部修正的情况
        elif os.path.exists(error_excel):
            df = pd.read_excel(error_excel)
            if "是否已修正（请打√）" in df.columns:
                marks = [str(v).strip() for v in df["是否已修正（请打√）"].dropna()]
                if all(m in ("√", "v", "V", "✓") for m in marks) and len(marks) == len(df):
                    drama_files = [f for f in std_files if any(f.startswith(e+"_") for e in get_ep_list(drama))]
                    cleaned_total += len(drama_files)
                    cleaned_correct += len(drama_files)
                    # 分配到个人
                    for f in drama_files:
                        for u in members:
                            if f"_{u}_" in f:
                                user_correct[u] += 1
                                user_cleaned_total[u] += 1

    # 计算准确率（避免除零）
    overall_accuracy = round(cleaned_correct / cleaned_total * 100, 1) if cleaned_total > 0 else 0
    user_accuracy = {u: round(user_correct[u]/user_cleaned_total[u]*100,1) if user_cleaned_total[u]>0 else 0 for u in members}

    # 4. DDL统计（整体DDL：最早逾期/最早截止；个人DDL：关联其提交的广播剧DDL）
    ddl_list = [(drama, datetime.strptime(ddl_str, "%Y-%m-%d %H:%M")) for drama, ddl_str in all_ddl.items()]
    overall_ddl_warning = []
    if ddl_list:
        ddl_list.sort(key=lambda x: x[1])
        earliest_ddl = ddl_list[0]
        latest_ddl = ddl_list[-1]
        # 整体DDL预警
        for drama, ddl_time in ddl_list:
            left = ddl_time - now
            if left < timedelta(0):
                overall_ddl_warning.append(f"{drama}（已逾期）")
            elif left < timedelta(hours=24):
                overall_ddl_warning.append(f"{drama}（即将截止，剩余{round(left.total_seconds()/3600,1)}小时）")

    # 5. 个性化提醒
    personal_reminder = []
    for u in members:
        # 未提交提醒
        if user_submit[u] == 0:
            personal_reminder.append(f"{u}：未提交任何文件")
        # 提交量过少提醒（低于平均提交量）
        avg_submit = sum(user_submit.values())/len(members) if members else 0
        if user_submit[u] > 0 and user_submit[u] < avg_submit/2:
            personal_reminder.append(f"{u}：提交量过少（当前{user_submit[u]}份，平均{round(avg_submit,1)}份）")
        # 准确率过低提醒
        if user_cleaned_total[u] > 0 and user_accuracy[u] < 80:
            personal_reminder.append(f"{u}：数据准确率过低（{user_accuracy[u]}%，低于80%）")

    # ------------------------------
    # 生成可视化图表（4合一综合图，组长可直接保存）
    # ------------------------------
    fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(16, 12))
    fig.suptitle(f"内容组监控综合报告（{now.strftime('%Y-%m-%d %H:%M')}）", fontsize=16, fontweight="bold")

    # 图1：个人提交数量柱状图
    names = list(user_submit.keys())
    submits = list(user_submit.values())
    ax1.bar(names, submits, color="#3498db")
    ax1.set_title("个人提交数量统计", fontsize=12, fontweight="bold")
    ax1.set_ylabel("提交份数")
    ax1.tick_params(axis="x", rotation=45)
    # 标注数值
    for i, v in enumerate(submits):
        ax1.text(i, v+0.1, str(v), ha="center", va="bottom")

    # 图2：整体进度饼图
    progress_labels = ["已提交", "未提交"]
    progress_data = [finished_ep_count, total_ep_count - finished_ep_count] if total_ep_count > 0 else [1, 0]
    ax2.pie(progress_data, labels=progress_labels, autopct="%1.1f%%", colors=["#2ecc71", "#e74c3c"])
    ax2.set_title(f"整体进度（总剧集数：{total_ep_count}）", fontsize=12, fontweight="bold")

    # 图3：个人准确率柱状图
    acc_values = list(user_accuracy.values())
    ax3.bar(names, acc_values, color="#f39c12")
    ax3.set_title("个人数据准确率统计", fontsize=12, fontweight="bold")
    ax3.set_ylabel("准确率（%）")
    ax3.set_ylim(0, 100)
    ax3.tick_params(axis="x", rotation=45)
    # 标注数值
    for i, v in enumerate(acc_values):
        ax3.text(i, v+1, f"{v}%", ha="center", va="bottom")

    # 图4：DDL预警（柱状图，显示剩余时间）
    if ddl_list:
        drama_names = [d[0] for d in ddl_list]
        left_hours = [(d[1] - now).total_seconds()/3600 for d in ddl_list]
        colors = ["#e74c3c" if h < 0 else "#f39c12" if h < 24 else "#2ecc71" for h in left_hours]
        ax4.bar(drama_names, left_hours, color=colors)
        ax4.set_title("各广播剧DDL剩余时间", fontsize=14, fontweight="bold")
        ax4.set_ylabel("剩余时间（小时）", fontsize=12)
        ax4.tick_params(axis="x", rotation=45, labelsize=11)
        # 标注逾期/即将截止 —— 字号加大、加粗、高亮
        for i, (h, d) in enumerate(zip(left_hours, drama_names)):
            if h < 0:
                ax4.text(i, -1, "已逾期", ha="center", va="top", 
                         color="red", fontsize=16, fontweight="bold")
            elif h < 24:
                ax4.text(i, h+1, "即将截止", ha="center", va="bottom", 
                         color="darkorange", fontsize=14, fontweight="bold")
    else:
        ax4.text(0.5, 0.5, "未设置任何DDL", ha="center", va="center", 
                 transform=ax4.transAxes, fontsize=14, fontweight="bold")
        ax4.set_title("各广播剧DDL剩余时间", fontsize=14, fontweight="bold")

    plt.tight_layout()
    plt.savefig("组长监控综合图表.png", dpi=100, bbox_inches="tight")
    plt.close()

    # ------------------------------
    # 生成文字版综合报告（可打印、可转发）
    # ------------------------------
    with open(MONITOR_REPORT_FILE, "w", encoding="utf-8") as f:
        f.write(f"内容组监控综合报告\n")
        f.write(f"报告生成时间：{now.strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write("="*50 + "\n\n")

        # 整体概况
        f.write("一、整体概况\n")
        f.write(f"  总广播剧数：{len(dramas)} 部\n")
        f.write(f"  总剧集数：{total_ep_count} 集\n")
        f.write(f"  已提交规范文件：{finished_ep_count} 份\n")
        f.write(f"  整体进度：{overall_progress}%\n")
        f.write(f"  整体数据准确率：{overall_accuracy}%\n\n")

        # 个人提交统计
        f.write("二、个人提交统计\n")
        f.write(f"{'姓名':<8}{'提交份数':<10}\n")
        f.write("-"*20 + "\n")
        for u, c in user_submit.items():
            f.write(f"{u:<8}{c:<10}\n")
        f.write("\n")

        # 个人准确率统计
        f.write("三、个人数据准确率统计\n")
        f.write(f"{'姓名':<8}{'已清洗文件数':<12}{'正确文件数':<10}{'准确率':<10}\n")
        f.write("-"*40 + "\n")
        for u in members:
            f.write(f"{u:<8}{user_cleaned_total[u]:<12}{user_correct[u]:<10}{user_accuracy[u]:<10}%\n")
        f.write("\n")

        # DDL统计
        f.write("四、DDL统计与预警\n")
        if ddl_list:
            f.write(f"  最早截止广播剧：{earliest_ddl[0]}（{earliest_ddl[1].strftime('%Y-%m-%d %H:%M')}）\n")
            f.write(f"  最晚截止广播剧：{latest_ddl[0]}（{latest_ddl[1].strftime('%Y-%m-%d %H:%M')}）\n")
            f.write("  DDL预警：\n")
            if overall_ddl_warning:
                for warn in overall_ddl_warning:
                    f.write(f"    - {warn}\n")
            else:
                f.write("    - 无预警，所有DDL正常\n")
        else:
            f.write("  未设置任何DDL\n")
        f.write("\n")

        # 个性化提醒
        f.write("五、个性化提醒\n")
        if personal_reminder:
            for remind in personal_reminder:
                f.write(f"  - {remind}\n")
        else:
            f.write("  - 无异常提醒，所有组员提交、准确率均正常\n")
        f.write("\n")

        # 移交技术组提示
        f.write("六、可移交技术组广播剧\n")
        if os.path.exists(FINAL_TRANSFER_FILE):
            with open(FINAL_TRANSFER_FILE, "r", encoding="utf-8") as tf:
                transfer_dramas = [line.strip() for line in tf if line.strip()]
            if transfer_dramas:
                f.write("  " + "、".join(transfer_dramas) + "\n")
            else:
                f.write("  暂无可移交技术组的广播剧\n")
        else:
            f.write("  暂无可移交技术组的广播剧\n")

    # ------------------------------
    # 控制台显示（精简版，方便组长快速查看）
    # ------------------------------
    os.system("cls")
    print("==================================================")
    print("                  组长监控综合仪表盘")
    print("==================================================")
    print(f"报告生成时间：{now.strftime('%Y-%m-%d %H:%M')}")
    print("="*50)
    print(f"整体概况：")
    print(f"  总广播剧：{len(dramas)}部 | 总剧集：{total_ep_count}集")
    print(f"  已提交：{finished_ep_count}份 | 整体进度：{overall_progress}%")
    print(f"  整体准确率：{overall_accuracy}%")
    print("="*50)

    print("\n【个人提交数量】")
    for u, c in user_submit.items():
        print(f"{u:8s} : {c} 份")

    print("\n【个人准确率】")
    for u, acc in user_accuracy.items():
        print(f"{u:8s} : {acc}%")

    print("\n【DDL预警】")
    if overall_ddl_warning:
        for warn in overall_ddl_warning:
            print(f"  - {warn}")
    else:
        print("  无预警")

    print("\n【个性化提醒】")
    if personal_reminder:
        for remind in personal_reminder:
            print(f"  - {remind}")
    else:
        print("  无异常提醒")

    print("\n" + "="*50)
    print("已生成以下文件，可直接保存、转发：")
    print("  1. 组长监控综合图表.png（可视化图表）")
    print(f"  2. {MONITOR_REPORT_FILE}（文字版报告）")
    input("\n回车返回")

# ------------------------------
# 数据清洗 + Excel错误表 + 自动识别打勾（不变，适配准确率统计）
# ------------------------------
def run_clean():
    dramas = get_all_dramas()
    print("\n==== 数据内容清洗 & 移交列表生成 ====")
    print("可清洗广播剧：", dramas)
    target = input("输入要清洗的广播剧：")
    if target not in dramas:
        print("无效广播剧名称")
        return

    eps = get_ep_list(target)
    std_files = [f for f in os.listdir(CHECK_FOLDER) if is_standard(f)]
    drama_files = [f for f in std_files if any(f.startswith(e+"_") for e in eps)]
    total = len(drama_files)
    if total == 0:
        print("该广播剧无已规范文件，无法清洗")
        return

    error_rows = []
    correct_count = 0

    for f in drama_files:
        path = os.path.join(CHECK_FOLDER, f)
        try:
            with open(path, "r", encoding="utf-8") as fp:
                lines = [l.strip() for l in fp if l.strip()]
        except:
            error_rows.append([f, "未知", "文件读取失败（编码错误或文件损坏）"])
            continue

        errs = []
        # 严格校验4个字段（剧集名、URL、非常驻角色、其他配音演员）
        if len(lines) < 4:
            errs.append("字段不足（需包含：剧集名、URL、非常驻角色、其他配音演员）")
        else:
            # URL校验（必须以http/https开头，长度≥10）
            if not lines[1].startswith(("http://", "https://")):
                errs.append("URL格式错误（需以http://或https://开头）")
            if len(lines[1]) < 10:
                errs.append("URL过短（疑似输入错误）")
            # 非空校验
            if not lines[2]:
                errs.append("非常驻角色为空")
            if not lines[3]:
                errs.append("其他配音演员为空")

        if not errs:
            correct_count += 1
        else:
            # 匹配收集人
            user = "未知"
            for u in json.load(open(USER_FILE,"r",encoding="utf-8")).values():
                if f"_{u}_" in f:
                    user = u
            error_rows.append([f, user, " | ".join(errs)])

    # 输出结果（无错误/有错误）
    if not error_rows:
        print(f"\n✅ {target} 全部数据通过校验，无错误！")
        with open(f"【{target}】全部通过.txt", "w", encoding="utf-8") as f:
            f.write(f"{target} 全部数据校验通过\n")
            f.write(f"校验时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}\n")
            f.write(f"总文件数：{total} | 正确文件数：{correct_count} | 准确率：100.0%")
    else:
        # 生成Excel错误表（组长打勾用）
        df = pd.DataFrame(error_rows, columns=["文件名", "收集人", "错误原因"])
        df["是否已修正（请打√）"] = ""
        df.to_excel(f"【{target}】错误处理表.xlsx", index=False)
        print(f"\n⚠️  {target} 存在{len(error_rows)}个错误文件")
        print(f"已生成错误处理表：【{target}】错误处理表.xlsx")
        print("提示：组长只需打开Excel在「是否已修正」列打√，无需操作程序，下次运行监控会自动识别")

    # 自动读取Excel打勾状态，判断是否全部完成
    all_done = False
    excel_path = f"【{target}】错误处理表.xlsx"
    if os.path.exists(excel_path):
        df_check = pd.read_excel(excel_path)
        if "是否已修正（请打√）" in df_check.columns:
            # 过滤空值，判断所有行是否都打了√
            marks = [str(v).strip() for v in df_check["是否已修正（请打√）"].dropna()]
            if all(m in ("√", "v", "V", "✓") for m in marks) and len(marks) == len(df_check):
                all_done = True
    else:
        # 无错误表 = 全部通过 = 已完成
        all_done = True

    # 已全部完成，更新移交列表并提醒
    if all_done:
        print(f"\n✅ 提醒：{target} 已全部修正完成，可移交技术组！")
        # 更新移交列表
        transfer_list = []
        if os.path.exists(FINAL_TRANSFER_FILE):
            with open(FINAL_TRANSFER_FILE, "r", encoding="utf-8") as f:
                transfer_list = [line.strip() for line in f if line.strip()]
        if target not in transfer_list:
            transfer_list.append(target)
        with open(FINAL_TRANSFER_FILE, "w", encoding="utf-8") as f:
            f.write("\n".join(transfer_list))
        print(f"已更新可移交列表：{FINAL_TRANSFER_FILE}")

    input("\n回车返回主页")

# ------------------------------
# 主菜单（不变）
# ------------------------------
def main():
    while True:
        os.system("cls")
        print("=============================================================================")
        print("                      内容组信息收集自动化系统 V Ultimate Final")
        print("=============================================================================")
        print(" 1  组员：文件自动规范化")
        print(" 2  组长：监控综合报告（可视化+文字版）")
        print(" 3  组长：数据清洗 & 错误表 & 生成技术组移交列表")
        print(" 4  组长：生成组员ID")
        print(" 5  组长：设置DDL")
        print(" 0  退出")
        print("=============================================================================")
        c = input("请选择功能：")
        if c == "1":
            member_tool()
        elif c == "2":
            if input("组长密钥：") == MASTER_KEY:
                run_monitor()
        elif c == "3":
            if input("组长密钥：") == MASTER_KEY:
                run_clean()
        elif c == "4":
            if input("组长密钥：") == MASTER_KEY:
                generate_member_list()
        elif c == "5":
            if input("组长密钥：") == MASTER_KEY:
                set_ddl()
        elif c == "0":
            break

if __name__ == "__main__":
    main()