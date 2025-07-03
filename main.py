import os
import google.generativeai as genai
import json
from datetime import date, timedelta
import random
import openpyxl
from openpyxl.styles import Alignment, Font

# --- 組態設定 ---
# 初始化 Gemini 模型
genai.configure(api_key=os.environ['GEMINI_API_KEY'])
model = genai.GenerativeModel('gemini-1.5-pro-latest')

# 員工與班別設定 (擴展至 100 人)
EMPLOYEES = [f"員工{i}" for i in range(1, 101)]
SHIFT_TYPES = ["早班", "午班", "晚班", "大夜班"]

# 每日各班別所需人力 (按比例增加)
SHIFTS_PER_DAY_REQUIREMENTS = {
    "weekday": {"早班": 20, "午班": 20, "晚班": 15, "大夜班": 10},
    "weekend": {"早班": 25, "午班": 25, "晚班": 20, "大夜班": 15}
}

# 員工個人班別限制 (擴展為更複雜的範例)
# 挑選一些員工並給予他們更真實的限制
constraint_employees = random.sample(EMPLOYEES, 15)
EMPLOYEE_CONSTRAINTS = {
    # 只能上早班或午班 (例如：學生、家庭主婦)
    constraint_employees[0]: ["早班", "午班"],
    constraint_employees[1]: ["早班", "午班"],
    constraint_employees[2]: ["早班", "午班"],
    constraint_employees[3]: ["早班", "午班"],
    
    # 只能上晚班或大夜班 (例如：夜校生)
    constraint_employees[4]: ["晚班", "大夜班"],
    constraint_employees[5]: ["晚班", "大夜班"],

    # 不能上大夜班 (例如：健康因素、通勤問題)
    constraint_employees[6]: ["早班", "午班", "晚班"],
    constraint_employees[7]: ["早班", "午班", "晚班"],
    constraint_employees[8]: ["早班", "午班", "晚班"],
    constraint_employees[9]: ["早班", "午班", "晚班"],

    # 只能上特定班別
    constraint_employees[10]: ["早班"],
    constraint_employees[11]: ["晚班"],

    # 只能上非週末班的班別 (此處為簡化，實際應在演算法中處理日期)
    # 這裡我們用班別類型來模擬，例如某些人不能上需求量大的班
    constraint_employees[12]: ["早班", "午班", "晚班"],

    # 跨班別的特殊組合
    constraint_employees[13]: ["早班", "大夜班"],
    constraint_employees[14]: ["午班", "晚班"],
}
print(f"已為 {len(EMPLOYEE_CONSTRAINTS)} 位員工設定特殊班別限制。")

# --- 第一步: 使用 LLM 解析自然語言的排班請求 ---

def parse_requests_with_llm(raw_requests_text):
    """
    使用 Gemini 模型將非結構化的文字請求轉換為結構化的 JSON。
    """
    prompt = f"""
    您是一位專業的 HR 排班助理。請仔細閱讀以下來自員工的排班請求文字，並將其轉換為一個結構化的 JSON 物件。

    JSON 物件必須包含一個 'leave_requests' 陣列。
    - 'leave_requests' 陣列中每個物件都應有 'employee' (員工姓名) 和 'date' (請假日期，格式為 YYYY-MM-DD)。

    請嚴格按照此格式輸出，不要有任何額外的說明文字。

    員工請求文字如下：
    ---
    {raw_requests_text}
    ---
    """
    
    print("--- 步驟 1: 正在呼叫 LLM 解析請假需求... ---")
    response = model.generate_content(prompt)
    
    # 清理並解析 LLM 的回應
    try:
        clean_response = response.text.strip().replace("```json", "").replace("```", "")
        print(f"LLM 原始回應 (清理後): {clean_response}")
        return json.loads(clean_response)
    except (json.JSONDecodeError, AttributeError) as e:
        print(f"無法解析 LLM 的回應: {e}")
        print(f"LLM原始輸出: {response.text}")
        return {"leave_requests": []}

# --- 第二步: 使用 Python 演算法進行排班 ---

def create_schedule(start_date, end_date, leave_requests):
    """
    核心排班演算法
    - 處理硬性限制：請假、每日人力需求
    - 處理軟性限制（公平性）：盡量均分週末班、大夜班
    """
    print("\n--- 步驟 2: 正在執行 Python 排班演算法... ---")
    
    schedule = {}
    # 追蹤每個員工的班別統計，用於實現公平性
    stats = {emp: {"total_shifts": 0, "weekend_shifts": 0, "大夜班": 0} for emp in EMPLOYEES}
    
    leave_map = {}
    for req in leave_requests.get('leave_requests', []):
        if req['employee'] in EMPLOYEES:
            leave_map.setdefault(req['date'], set()).add(req['employee'])

    delta = end_date - start_date
    for i in range(delta.days + 1):
        current_date = start_date + timedelta(days=i)
        date_str = current_date.strftime("%Y-%m-%d")
        schedule[date_str] = {}
        
        day_type = "weekend" if current_date.weekday() >= 5 else "weekday"
        requirements = SHIFTS_PER_DAY_REQUIREMENTS[day_type]
        
        on_leave_today = leave_map.get(date_str, set())
        
        for shift, count in requirements.items():
            schedule[date_str][shift] = []
            
            # 預先分配當天已排班的人員，避免重複排班
            assigned_today_flat = {emp for s_list in schedule[date_str].values() for emp in s_list}

            for _ in range(count):
                # 找出最適合上此班的員工
                # 演算法：從可用的人中，找一個「分數」最低的
                # 分數越低越優先 (例如: 週末班最少的人優先上週末班)
                
                available_employees = [
                    emp for emp in EMPLOYEES 
                    if emp not in on_leave_today and emp not in assigned_today_flat
                ]
                
                # 過濾掉不符合班別限制的員工
                available_employees = [
                    emp for emp in available_employees
                    if shift in EMPLOYEE_CONSTRAINTS.get(emp, SHIFT_TYPES)
                ]

                if not available_employees:
                    schedule[date_str][shift].append("!!人力不足!!")
                    continue

                def calculate_fairness_score(emp):
                    score = stats[emp]['total_shifts'] * 0.1 # 基礎分數
                    if day_type == 'weekend':
                        score += stats[emp]['weekend_shifts'] * 1.0 # 週末班權重高
                    if shift == '大夜班':
                        score += stats[emp]['大夜班'] * 1.5 # 大夜班權重更高
                    return score

                best_employee = min(available_employees, key=calculate_fairness_score)
                
                # 分配班表並更新統計數據
                schedule[date_str][shift].append(best_employee)
                assigned_today_flat.add(best_employee) # 將新排入的員工加入今日已排班清單
                stats[best_employee]['total_shifts'] += 1
                if day_type == 'weekend':
                    stats[best_employee]['weekend_shifts'] += 1
                if shift == '大夜班':
                    stats[best_employee]['大夜班'] += 1

    print("排班演算法完成！")
    return schedule, stats, leave_map

# --- 第三步: 使用 LLM 生成摘要報告與通知 ---

def summarize_with_llm(schedule, stats):
    """
    將排班結果交給 LLM，生成易讀的報告。
    """
    print("\n--- 步驟 3: 正在呼叫 LLM 生成摘要報告與通知... ---")
    
    schedule_json = json.dumps(schedule, indent=2, ensure_ascii=False)
    stats_json = json.dumps(stats, indent=2, ensure_ascii=False)

    prompt = f"""
    您是一位細心且善於溝通的 HR 經理。
    這是一份剛剛產出的員工班表草案，以及相關的公平性統計數據。

    請根據這些資料，完成以下兩項任務：

    1.  **生成摘要報告**:
        - 簡要說明這份班表的涵蓋期間。
        - 根據 'stats' 資料，進行公平性檢查。指出週末班或大夜班次數最多和最少的員工，並評估排班是否大致公平。
        - 檢查 'schedule' 中是否有 '!!人力不足!!' 的標記，並提出警告。

    2.  **草擬個人化通知 (範例)**:
        - 請為「員工1」和「員工2」這兩位員工，草擬一則簡短溫馨的排班通知訊息。訊息中需清楚列出他們各自的上班日期與班別。

    請用繁體中文、專業且友善的語氣來撰寫。

    班表資料 (JSON):
    {schedule_json}

    公平性統計資料 (JSON):
    {stats_json}
    """
    
    response = model.generate_content(prompt)
    return response.text

# --- 第四步: 將班表匯出至 Excel ---

def export_to_excel(schedule, employees, shift_types, leave_map, filename="shift_schedule.xlsx"):
    """
    將排班結果匯出成一個格式化的 Excel 檔案。
    - 包含一個所有人的總表，並標示請假。
    - 為每位員工建立一個個人班表。
    """
    print(f"\n--- 步驟 4: 正在將班表匯出至 Excel ({filename})... ---")
    
    wb = openpyxl.Workbook()
    center_alignment = Alignment(horizontal='center', vertical='center')
    bold_font = Font(bold=True)
    leave_font = Font(color="FF0000") # 紅色字體標示休假

    # --- 建立總表 ---
    summary_sheet = wb.active
    summary_sheet.title = "班表總表"

    # 建立表頭 (日期)
    dates = sorted(schedule.keys())
    summary_sheet.cell(row=1, column=1, value="員工").font = bold_font
    for col, date in enumerate(dates, start=2):
        cell = summary_sheet.cell(row=1, column=col, value=date)
        cell.font = bold_font
        cell.alignment = center_alignment

    # 建立員工列表與班別資料
    employee_row_map = {emp: i for i, emp in enumerate(employees, start=2)}
    for emp, row_idx in employee_row_map.items():
        summary_sheet.cell(row=row_idx, column=1, value=emp)

    # 填充班表內容
    for date_idx, date in enumerate(dates, start=2):
        on_leave_today = leave_map.get(date, set())
        for emp, row_idx in employee_row_map.items():
            cell = summary_sheet.cell(row=row_idx, column=date_idx)
            if emp in on_leave_today:
                cell.value = "休"
                cell.font = leave_font
                cell.alignment = center_alignment
            else:
                # 找到該員工當天的班別
                assigned_shift = None
                for shift, assigned_employees in schedule[date].items():
                    if emp in assigned_employees:
                        assigned_shift = shift
                        break
                if assigned_shift:
                    cell.value = assigned_shift
                    cell.alignment = center_alignment

    # --- 為每位員工建立個人班表 ---
    for emp in employees:
        emp_sheet = wb.create_sheet(title=f"{emp}個人班表")
        
        # 建立表頭
        emp_sheet.cell(row=1, column=1, value="班別").font = bold_font
        for col, date in enumerate(dates, start=2):
            cell = emp_sheet.cell(row=1, column=col, value=date)
            cell.font = bold_font
            cell.alignment = center_alignment

        # 建立班別列表
        shift_row_map = {shift: i for i, shift in enumerate(shift_types, start=2)}
        for shift, row_idx in shift_row_map.items():
            emp_sheet.cell(row=row_idx, column=1, value=shift)

        # 填充打勾符號
        for date_idx, date in enumerate(dates, start=2):
            on_leave_today = leave_map.get(date, set())
            if emp in on_leave_today:
                 for row_idx in shift_row_map.values():
                    cell = emp_sheet.cell(row=row_idx, column=date_idx, value="休")
                    cell.font = leave_font
                    cell.alignment = center_alignment
            else:
                for shift, assigned_employees in schedule[date].items():
                    if emp in assigned_employees:
                        row_idx = shift_row_map[shift]
                        cell = emp_sheet.cell(row=row_idx, column=date_idx, value="✓")
                        cell.alignment = center_alignment
    
    wb.save(filename)
    print("Excel 檔案匯出成功！")


# --- 主程式執行流程 ---

if __name__ == "__main__":
    # 模擬的員工自然語言請求 (更新為新員工範例)
    raw_requests = """
    To HR:
    我是員工1，我因為家裡有事，想要在 2025年7月15日 請假一天，謝謝。
    
    員工2:
    你好，我想請 2025-07-18，再麻煩了。
    
    員工3:
    老闆，7/15 我沒辦法上班喔。
    """

    # 第一步
    parsed_data = parse_requests_with_llm(raw_requests)

    # 第二步
    start_date = date(2025, 7, 14)
    end_date = date(2025, 7, 20)
    final_schedule, final_stats, leave_map = create_schedule(start_date, end_date, parsed_data)

    print("\n--- 最終班表 (Python 物件) ---")
    print(json.dumps(final_schedule, indent=2, ensure_ascii=False))
    print("\n--- 公平性統計 (Python 物件) ---")
    print(json.dumps(final_stats, indent=2, ensure_ascii=False))

    # 第三步
    summary_report = summarize_with_llm(final_schedule, final_stats)
    print("\n\n" + "="*50)
    print("     由 Gemini Pro 生成的排班摘要報告與通知")
    print("="*50)
    print(summary_report)

    # 第四步
    export_to_excel(final_schedule, EMPLOYEES, SHIFT_TYPES, leave_map)
