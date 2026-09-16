import os
import json
import shutil
import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path

# 配置常量
CONFIG_FILE = "member_config.json"  # 存储圈名和任务绑定信息
BASE_DIR = Path.cwd()  # 程序运行目录
# 各阶段对应的文件和文件夹配置
STAGE_CONFIG = {
    "信息收集": {
        "excel_file": "广播剧剧集.xlsx",
        "folder": "信息待检查文件夹",
        "name_prefix": "",
        "name_suffix": "信息",
        "file_ext": ".csv"
    },
    "重点剧集审核": {
        "excel_file": "审核表.xlsx",
        "folder": "审核待检查文件夹",
        "name_prefix": "审核",
        "name_suffix": "",
        "file_ext": ".txt"
    },
    "人工抽检复核": {
        "excel_file": "抽检表.xlsx",
        "folder": "抽检待检查文件夹",
        "name_prefix": "抽检",
        "name_suffix": "",
        "file_ext": ".txt"
    },
    "兜底阶段": {
        "excel_file": "需人工转接.xlsx",
        "folder": "需人工转接待检查文件夹",
        "name_prefix": "抽检",
        "name_suffix": "",
        "file_ext": ".txt"
    }
}

def load_config():
    """加载配置文件，返回圈名和任务信息"""
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"cn": "", "tasks": {stage: [] for stage in STAGE_CONFIG.keys()}}

def save_config(config):
    """保存配置到文件"""
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)

def verify_cn():
    """验证圈名，确保两次输入一致并保存"""
    print("欢迎使用该程序，该程序主要用于认领任务，确认ddl，以及最后的文件命名规范化。现在请告诉我您的cn，方便署名和持续使用。")
    while True:
        cn1 = input("请输入您的圈名，文件命名需要用到：").strip()
        print("确认后不可更改，请您再输入一遍。")
        cn2 = input().strip()
        if cn1 == cn2 and cn1:
            config = load_config()
            config["cn"] = cn1
            save_config(config)
            print(f"欢迎{cn1}加入逐声计划!")
            return cn1
        else:
            print("检测到输入有出入，请确认您的cn")

def normalize_task_name(task_name):
    """标准化任务名（去除符号、空格、大小写）"""
    if not task_name:
        return ""
    # 去除特殊符号和空格，转大写
    cleaned = ''.join([c for c in task_name if c.isalnum() or '\u4e00' <= c <= '\u9fff']).upper()
    return cleaned

def get_task_list(stage):
    """读取对应阶段的Excel，生成可认领的任务列表（修复返回值为3个）"""
    excel_path = STAGE_CONFIG[stage]["excel_file"]
    if not os.path.exists(excel_path):
        # 修复：返回3个值（空列表、空字典、空字典）
        return [], {}, {}  
    
    try:
        df = pd.read_excel(excel_path, header=None)
        task_names = []
        task_ddls = {}
        task_name_map = {}  # 标准化名→原任务名
        # 解析Excel结构：第一行是任务名（每两列一个任务）
        row1 = df.iloc[0].tolist()  # 第一行：任务名
        col_idx = 0
        while col_idx < len(row1):
            task_name = str(row1[col_idx]).strip()
            if task_name and task_name != "nan":
                task_names.append(task_name)
                norm_name = normalize_task_name(task_name)
                task_name_map[norm_name] = task_name  # 存储映射关系
                # 读取该任务的ddl列（当前任务列+1）
                ddl_col = col_idx + 1
                if ddl_col < df.shape[1]:
                    # 获取该列所有非空的ddl值（第三行开始）
                    ddl_vals = df.iloc[2:, ddl_col].dropna().tolist()
                    if ddl_vals:
                        # 取第一个非空ddl作为任务ddl
                        task_ddls[task_name] = str(ddl_vals[0])
            col_idx += 2
        # 修复：返回3个值
        return task_names, task_ddls, task_name_map
    except Exception as e:
        print(f"读取任务列表失败：{e}")
        # 修复：异常时也返回3个值
        return [], {}, {}

def get_total_task_count(stage):
    """获取指定阶段可认领的总任务数量（信息收集阶段专用）"""
    excel_path = STAGE_CONFIG[stage]["excel_file"]
    if not os.path.exists(excel_path):
        return 0
    
    try:
        df = pd.read_excel(excel_path, header=None)
        row1 = df.iloc[0].tolist()
        task_count = 0
        col_idx = 0
        while col_idx < len(row1):
            task_name = str(row1[col_idx]).strip()
            if task_name and task_name != "nan":
                task_count += 1
            col_idx += 2
        return task_count
    except Exception as e:
        print(f"统计总任务数失败：{e}")
        return 0

def claim_task(stage, cn):
    """认领指定阶段的任务（修复重复认领+优化匹配逻辑）"""
    # 关键修复：每次认领前重新加载最新配置，确保读取到已认领的任务
    config = load_config()  # 移到循环外，每次认领任务前强制刷新配置
    task_names, task_ddls, task_name_map = get_task_list(stage)
    
    if not task_names:
        print("当前没有任务可领取")
        return
    
    print(f"\n可认领的{stage}任务列表：")
    for i, task in enumerate(task_names, 1):
        print(f"{i}. {task}")
    
    # 确认认领数量
    while True:
        try:
            task_count = input("请输入您要认领的任务数量：").strip()
            if not task_count.isdigit():
                print("请输入有效的数字！")
                continue
            task_count = int(task_count)
            if task_count <= 0:
                print("任务数量必须大于0！")
                continue
            break
        except:
            print("输入格式错误，请重新输入！")
    
    # 多任务确认
    if task_count > 1:
        confirm = input(f"好的，现在你确认你要认领{task_count}个任务，请确保您的时间充裕，如若权衡后觉得时间较紧张，请输入q退出，然后与组长自行商量，确认认领请输入OK：").strip()
        if confirm.lower() == "q":
            print("已退出任务认领，请与组长沟通后重新操作。")
            return
        elif confirm.upper() != "OK":
            print("输入错误，返回任务认领界面！")
            claim_task(stage, cn)
            return
    
    # 逐个认领任务
    claimed_tasks = []
    for i in range(1, task_count + 1):
        while True:
            input_name = input(f"请告诉我您认领的第{i}个任务是：").strip()
            # 标准化输入的任务名
            norm_input = normalize_task_name(input_name)
            # 匹配标准化后的任务名
            if norm_input in task_name_map:
                task_name = task_name_map[norm_input]  # 获取原始任务名
                
                # 关键修复：实时检查已认领任务（从最新配置中读取）
                current_config = load_config()  # 每次检查前刷新配置
                existing_tasks = [t["task_name"] for t in current_config["tasks"][stage]]
                if task_name in existing_tasks:
                    print(f"您已认领过【{task_name}】任务，请勿重复认领！")
                    # 直接跳出当前任务的输入循环，不再重试
                    break
                
                # 绑定任务和ddl
                ddl = task_ddls.get(task_name, "")
                # 信息收集阶段：total存储已认领任务数（1个任务算1），其他阶段存储剧集数
                if stage == "信息收集":
                    total = 1  # 每个任务计为1
                else:
                    total = get_task_total(stage, task_name)
                
                claimed_tasks.append({
                    "task_name": task_name,
                    "ddl": ddl,
                    "submitted": 0,  # 已提交数量
                    "total": total  # 总数量（信息收集=1，其他=剧集数）
                })
                print(f"成功认领任务：{task_name}（DDL：{ddl}）")
                break
            else:
                print(f"任务名【{input_name}】不在可认领列表中，请重新输入！")
                print(f"可认领的任务名参考：{', '.join(task_names)}")
    
    # 更新配置
    if claimed_tasks:  # 只有新增任务时才更新
        config = load_config()  # 再次刷新最新配置
        config["tasks"][stage].extend(claimed_tasks)
        save_config(config)
        print(f"\n本次共成功认领{len(claimed_tasks)}个{stage}任务，请及时告知组长！")
    else:
        print("\n未认领任何新任务！")

def get_task_total(stage, task_name):
    """获取指定任务的总剧集数（非信息收集阶段使用）"""
    excel_path = STAGE_CONFIG[stage]["excel_file"]
    if not os.path.exists(excel_path):
        return 0
    
    try:
        df = pd.read_excel(excel_path, header=None)
        row1 = df.iloc[0].tolist()
        col_idx = 0
        while col_idx < len(row1):
            if str(row1[col_idx]).strip() == task_name:
                # 统计该任务列（完整剧集名列）的非空行数（第三行开始）
                total = df.iloc[2:, col_idx].dropna().shape[0]
                return total
            col_idx += 2
        return 0
    except Exception as e:
        print(f"统计任务总数失败：{e}")
        return 0

def check_progress(stage, cn):
    """查询指定阶段的任务进度和DDL（修复DDL格式解析）"""
    config = load_config()
    tasks = config["tasks"][stage]
    
    if not tasks:
        print(f"您暂时没有领取{stage}流程的任务哦")
        return
    
    print(f"\n=== {stage}任务进度查询 ===")
    today = datetime.now()
    
    # 信息收集阶段：统计任务级进度
    if stage == "信息收集":
        total_claimed = len(tasks)
        total_submitted = sum([1 for t in tasks if t["submitted"] > 0])
        print(f"总认领任务数：{total_claimed} | 已提交任务数：{total_submitted} | 剩余任务数：{total_claimed - total_submitted}")
        print("-" * 30)
    
    for task in tasks:
        task_name = task["task_name"]
        ddl_str = task["ddl"]
        submitted = task["submitted"]
        total = task["total"]
        
        # 计算剩余天数
        remaining_days = "未知"
        if ddl_str:
            try:
                # 先尝试解析带时间的格式
                ddl_date = datetime.strptime(ddl_str, "%Y-%m-%d %H:%M:%S")
            except:
                try:
                    # 再尝试解析无时间的格式
                    ddl_date = datetime.strptime(ddl_str, "%Y/%m/%d")
                except:
                    remaining_days = "格式错误"
                else:
                    remaining_days = (ddl_date - today).days
                    remaining_days = "已逾期" if remaining_days < 0 else f"{remaining_days}天"
            else:
                remaining_days = (ddl_date - today).days
                remaining_days = "已逾期" if remaining_days < 0 else f"{remaining_days}天"
        
        print(f"任务名：{task_name}")
        if stage == "信息收集":
            print(f"任务状态：{'已提交' if submitted > 0 else '未提交'}")
        else:
            print(f"总剧集数：{total} | 已提交：{submitted} | 剩余：{total - submitted}")
        print(f"DDL：{ddl_str} | 距离DDL：{remaining_days}")
        print("-" * 30)

def rename_files(stage, cn):
    """规范化指定阶段的文件命名（信息收集按任务名匹配，其他按剧集名）"""
    config = load_config()
    tasks = config["tasks"][stage]
    
    if not tasks:
        print(f"您暂时没有领取{stage}流程的任务哦，无法进行文件重命名！")
        return
    
    # 获取配置
    stage_info = STAGE_CONFIG[stage]
    # 强制使用绝对路径，避免中文/空格路径问题
    folder_path = Path(os.path.abspath(stage_info["folder"]))
    excel_path = os.path.abspath(stage_info["excel_file"])
    
    # 调试：打印关键路径（精确到绝对路径）
    print(f"\n【调试信息】")
    print(f"待检查文件夹绝对路径：{folder_path}")
    print(f"Excel文件绝对路径：{excel_path}")
    print(f"该阶段配置的文件后缀：{stage_info['file_ext']}")
    print(f"已认领的任务：{[t['task_name'] for t in tasks]}")
    
    # 检查文件夹是否存在，不存在则创建
    if not folder_path.exists():
        folder_path.mkdir(parents=True, exist_ok=True)
        print(f"文件夹不存在，已自动创建：{folder_path}")
    
    if not os.path.exists(excel_path):
        print(f"未找到Excel文件：{excel_path}，请确认文件在程序同目录！")
        return
    
    # 读取匹配所需数据
    df = pd.read_excel(excel_path, header=None)
    match_data = []
    row1 = df.iloc[0].tolist()
    col_idx = 0
    
    if stage == "信息收集":
        # 信息收集：读取已认领的任务名（用于匹配）
        for task in tasks:
            match_data.append(task["task_name"])
    else:
        # 其他阶段：读取已认领任务对应的剧集名
        for task in tasks:
            task_name = task["task_name"]
            while col_idx < len(row1):
                if str(row1[col_idx]).strip() == task_name:
                    episodes = df.iloc[2:, col_idx].dropna().tolist()
                    for ep in episodes:
                        ep_full = str(ep).strip()
                        if ep_full and ep_full != "nan":
                            match_data.append(ep_full)
                    break
                col_idx += 2
            col_idx = 0
    
    if not match_data:
        print(f"未找到可匹配的{'任务名' if stage == '信息收集' else '剧集名'}！")
        return
    
    # 获取版本号
    while True:
        version = input("\n请输入版本号（数字即可，程序会自动添加V前缀）：").strip()
        if version.isdigit():
            version = f"V{version}"
            break
        else:
            print("版本号必须为数字，请重新输入！")
    
    # 遍历文件夹中的所有文件（匹配.txt和.csv）
    today = datetime.now().strftime("%Y%m%d")
    renamed_count = 0
    target_files = list(folder_path.glob("*.txt")) + list(folder_path.glob("*.csv"))
    print(f"\n【调试】待处理文件列表（.txt+.csv）：")
    for f in target_files:
        print(f"  - {f.name} | 路径：{f}")
    
    if not target_files:
        print(f"待检查文件夹中无.txt/.csv文件，请确认文件位置！")
        return
    
    # 清洗文本函数
    def clean_text(text):
        text = text.replace("：", ":").replace("　", " ").replace("，", ",").replace("。", ".")
        return ''.join([c for c in text if c.isalnum() or '\u4e00' <= c <= '\u9fff']).upper()
    
    for file in target_files:
        # 获取文件名（去后缀）
        file_name_key = file.stem.strip()
        if not file_name_key:
            continue
        
        matched_item = None
        clean_key = clean_text(file_name_key)
        
        # 匹配逻辑
        if stage == "信息收集":
            # 信息收集：匹配任务名
            for task_name in match_data:
                clean_task = clean_text(task_name)
                if clean_key in clean_task or clean_task in clean_key:
                    matched_item = task_name
                    break
        else:
            # 其他阶段：匹配剧集名
            for ep_name in match_data:
                clean_ep = clean_text(ep_name)
                if clean_key in clean_ep or clean_ep in clean_key:
                    matched_item = ep_name
                    break
        
        if matched_item:
            # 命名规范：
            if stage == "信息收集":
                # 信息收集：任务名_操作人_日期_版本号.csv
                new_name = f"{matched_item}_{cn}_{today}_{version}.csv"
            else:
                # 其他阶段：前缀+剧集名+后缀_操作人_日期_版本号.原后缀
                file_ext = file.suffix
                new_name = f"{stage_info['name_prefix']}{matched_item}{stage_info['name_suffix']}_{cn}_{today}_{version}{file_ext}"
            
            new_path = folder_path / new_name
            
            try:
                # 先删除已存在的同名文件
                if new_path.exists():
                    new_path.unlink()
                # 重命名文件
                file.rename(new_path)
                print(f"\n重命名成功：")
                print(f"  原文件：{file.name}")
                print(f"  新文件：{new_name}")
                renamed_count += 1
                
                # 更新已提交数量
                config = load_config()
                for task in config["tasks"][stage]:
                    if stage == "信息收集":
                        # 信息收集：匹配到任务名则标记为已提交
                        if clean_text(task["task_name"]) == clean_text(matched_item):
                            task["submitted"] = 1  # 任务提交状态置为1
                            break
                    else:
                        # 其他阶段：匹配到剧集名则计数+1
                        episodes = get_task_episodes(stage, task["task_name"])
                        if matched_item in [str(ep) for ep in episodes]:
                            task["submitted"] += 1
                            break
                save_config(config)
            except Exception as e:
                print(f"\n重命名失败【{file.name}】：{str(e)}")
                print(f"  解决：关闭文件/检查文件夹权限")
        else:
            print(f"\n匹配失败：{file.name}")
            print(f"  清洗后的关键词：{clean_key}")
            print(f"  可匹配的{('任务名' if stage == '信息收集' else '剧集名')}（清洗后）：{[clean_text(item) for item in match_data]}")
    
    print(f"\n=== 重命名结果 ===")
    print(f"待处理文件数：{len(target_files)} | 成功重命名：{renamed_count}")
    print(f"请将重命名后的文件移出【{folder_path.name}】，避免重复处理！")

def get_task_episodes(stage, task_name):
    """获取指定任务的所有剧集名"""
    excel_path = STAGE_CONFIG[stage]["excel_file"]
    if not os.path.exists(excel_path):
        return []
    
    try:
        df = pd.read_excel(excel_path, header=None)
        row1 = df.iloc[0].tolist()
        col_idx = 0
        while col_idx < len(row1):
            if str(row1[col_idx]).strip() == task_name:
                return df.iloc[2:, col_idx].dropna().tolist()
            col_idx += 2
        return []
    except Exception as e:
        print(f"获取剧集名失败：{e}")
        return []

def check_ddl_reminder(cn):
    """检查是否有DDL剩余4天内的任务，生成提醒"""
    config = load_config()
    reminders = []
    today = datetime.now()
    
    for stage, tasks in config["tasks"].items():
        for task in tasks:
            ddl_str = task["ddl"]
            if ddl_str:
                try:
                    # 兼容两种DDL格式
                    try:
                        ddl_date = datetime.strptime(ddl_str, "%Y-%m-%d %H:%M:%S")
                    except:
                        ddl_date = datetime.strptime(ddl_str, "%Y/%m/%d")
                    remaining_days = (ddl_date - today).days
                    if 0 <= remaining_days <= 4:
                        reminders.append(f"{stage}任务【{task['task_name']}】距离ddl还剩{remaining_days}天")
                except:
                    pass
    
    if reminders:
        return "\n".join(reminders)
    return ""

def stage_menu(stage, cn):
    """各阶段的子菜单"""
    while True:
        print(f"\n=== {stage}操作菜单 ===")
        print("1. 认领任务")
        print("2. 查询任务进度及DDL")
        print("3. 文件命名规范化")
        print("0. 返回主菜单")
        
        choice = input("请输入操作编号：").strip()
        if choice == "1":
            claim_task(stage, cn)
        elif choice == "2":
            check_progress(stage, cn)
        elif choice == "3":
            rename_files(stage, cn)
        elif choice == "0":
            break
        else:
            print("输入错误，请选择有效的操作编号！")

def main():
    """主程序入口"""
    # 加载配置
    config = load_config()
    cn = config["cn"]
    
    # 首次使用：验证圈名
    if not cn:
        cn = verify_cn()
        config = load_config()  # 重新加载配置
    
    # 非首次使用：欢迎语+DDL提醒
    else:
        reminder = check_ddl_reminder(cn)
        welcome_msg = f"欢迎{cn}回来，请输入回车进入程序!"
        if reminder:
            welcome_msg += f"\n请注意，{reminder}，请合理安排时间，非常感谢您"
        input(welcome_msg)
    
    # 主菜单
    while True:
        print("\n=== 逐声计划组员端主菜单 ===")
        print("1. 信息收集")
        print("2. 重点剧集审核")
        print("3. 人工抽检复核")
        print("4. 兜底阶段")
        print("0. 退出程序")
        
        choice = input("请选择您要操作的阶段：").strip()
        if choice == "1":
            stage_menu("信息收集", cn)
        elif choice == "2":
            stage_menu("重点剧集审核", cn)
        elif choice == "3":
            stage_menu("人工抽检复核", cn)
        elif choice == "4":
            stage_menu("兜底阶段", cn)
        elif choice == "0":
            print("感谢使用，程序已退出！")
            break
        else:
            print("输入错误，请选择有效的阶段编号！")

if __name__ == "__main__":
    # 创建必要的文件夹（首次运行时）
    for stage in STAGE_CONFIG.values():
        folder = BASE_DIR / stage["folder"]
        if not folder.exists():
            folder.mkdir(exist_ok=True)
            print(f"已创建文件夹：{folder}")
    
    main()