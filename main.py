import os
import google.generativeai as genai
import json
from datetime import date, timedelta, datetime
import random
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

# --- 組態設定 ---
# 初始化 Gemini 模型
genai.configure(api_key=os.environ["GEMINI_API_KEY"])
model = genai.GenerativeModel("gemini-1.5-pro-latest")

# 員工與班別設定 (擴展至 100 人)
EMPLOYEES = [f"員工{i}" for i in range(1, 101)]
SHIFT_TYPES = ["早班", "午班", "晚班", "大夜班"]

# 每日各班別所需人力 (按比例增加)
SHIFTS_PER_DAY_REQUIREMENTS = {
    "weekday": {"早班": 20, "午班": 20, "晚班": 15, "大夜班": 10},
    "weekend": {"早班": 25, "午班": 25, "晚班": 20, "大夜班": 15},
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

# 將當次執行的隨機限制寫入檔案，以供查閱
with open("employee_constraints.txt", "w", encoding="utf-8") as f:
    f.write("當次執行所使用的員工班別限制：\n")
    f.write(json.dumps(EMPLOYEE_CONSTRAINTS, indent=2, ensure_ascii=False))
print("詳細的員工班別限制已寫入 employee_constraints.txt 檔案。")


# --- 共用函式 ---
def calculate_fairness_score(emp, stats, day_type, shift):
    """計算員工的公平性分數"""
    score = stats[emp]["total_shifts"] * 0.1  # 基礎分數
    if day_type == "weekend":
        score += stats[emp]["weekend_shifts"] * 1.0  # 週末班權重高
    if shift == "大夜班":
        score += stats[emp]["大夜班"] * 1.5  # 大夜班權重更高
    return score


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
    """
    print("\n--- 步驟 2: 正在執行 Python 排班演算法... ---")

    schedule = {}
    stats = {
        emp: {"total_shifts": 0, "weekend_shifts": 0, "大夜班": 0} for emp in EMPLOYEES
    }

    leave_map = {}
    for req in leave_requests.get("leave_requests", []):
        if req["employee"] in EMPLOYEES:
            leave_map.setdefault(req["date"], set()).add(req["employee"])

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
            assigned_today_flat = {
                emp for s_list in schedule[date_str].values() for emp in s_list
            }

            for _ in range(count):
                available_employees = [
                    emp
                    for emp in EMPLOYEES
                    if emp not in on_leave_today and emp not in assigned_today_flat
                ]

                available_employees = [
                    emp
                    for emp in available_employees
                    if shift in EMPLOYEE_CONSTRAINTS.get(emp, SHIFT_TYPES)
                ]

                if not available_employees:
                    schedule[date_str][shift].append("!!人力不足!!")
                    continue

                best_employee = min(
                    available_employees,
                    key=lambda emp: calculate_fairness_score(
                        emp, stats, day_type, shift
                    ),
                )

                schedule[date_str][shift].append(best_employee)
                assigned_today_flat.add(best_employee)
                stats[best_employee]["total_shifts"] += 1
                if day_type == "weekend":
                    stats[best_employee]["weekend_shifts"] += 1
                if shift == "大夜班":
                    stats[best_employee]["大夜班"] += 1

    print("排班演算法完成！")
    return schedule, stats, leave_map


# --- 新增步驟: 最小幅度調整機制 ---
def find_and_assign_replacement(
    schedule, stats, leave_map, emp_on_leave, leave_date_str
):
    """
    為臨時請假的員工尋找最適合的代班人選，並只做最小幅度修改。
    """
    print(f"\n--- 正在為 {leave_date_str} 的 {emp_on_leave} 尋找代班人選... ---")

    # 1. 找出請假員工原本的班別
    original_shift = None
    for shift, employees in schedule.get(leave_date_str, {}).items():
        if emp_on_leave in employees:
            original_shift = shift
            break

    if not original_shift:
        print(f"錯誤：員工 {emp_on_leave} 在 {leave_date_str} 並沒有排班。")
        return schedule, stats, leave_map, False

    # 2. 更新資料：將員工從班表移除，並加入請假地圖
    schedule[leave_date_str][original_shift].remove(emp_on_leave)
    leave_map.setdefault(leave_date_str, set()).add(emp_on_leave)

    # 3. 更新原員工的統計數據
    day_is_weekend = datetime.strptime(leave_date_str, "%Y-%m-%d").weekday() >= 5
    stats[emp_on_leave]["total_shifts"] -= 1
    if day_is_weekend:
        stats[emp_on_leave]["weekend_shifts"] -= 1
    if original_shift == "大夜班":
        stats[emp_on_leave]["大夜班"] -= 1

    # 4. 尋找代班人選
    # 4.1. 取得當天所有已排班或已請假的員工
    unavailable_today = set(leave_map.get(leave_date_str, set()))
    for shift_employees in schedule.get(leave_date_str, {}).values():
        unavailable_today.update(shift_employees)

    # 4.2. 篩選出所有可代班的候選人
    potential_replacements = [
        emp
        for emp in EMPLOYEES
        if emp not in unavailable_today
        and original_shift in EMPLOYEE_CONSTRAINTS.get(emp, SHIFT_TYPES)
    ]

    if not potential_replacements:
        print(
            f"警告：找不到任何符合資格的員工可以替補 {leave_date_str} 的 {original_shift}。"
        )
        schedule[leave_date_str][original_shift].append("!!人力不足!!")
        return schedule, stats, leave_map, False

    # 5. 根據公平性分數，找出最佳代班人
    day_type = "weekend" if day_is_weekend else "weekday"
    best_replacement = min(
        potential_replacements,
        key=lambda emp: calculate_fairness_score(emp, stats, day_type, original_shift),
    )

    # 6. 分派代班並更新其統計數據
    schedule[leave_date_str][original_shift].append(best_replacement)
    stats[best_replacement]["total_shifts"] += 1
    if day_is_weekend:
        stats[best_replacement]["weekend_shifts"] += 1
    if original_shift == "大夜班":
        stats[best_replacement]["大夜班"] += 1

    print(
        f"成功！由 {best_replacement} 替補 {emp_on_leave} 在 {leave_date_str} 的 {original_shift}。"
    )
    return schedule, stats, leave_map, True


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


def export_to_excel(
    schedule, employees, shift_types, leave_map, filename="shift_schedule.xlsx"
):
    """
    將排班結果匯出成一個格式化的 Excel 檔案。
    """
    print(f"\n--- 步驟 4: 正在將班表匯出至 Excel ({filename})... ---")

    wb = openpyxl.Workbook()
    center_alignment = Alignment(horizontal="center", vertical="center")
    bold_font = Font(bold=True)
    leave_font = Font(color="FF0000", bold=True)
    leave_fill = PatternFill(
        start_color="FFFF00", end_color="FFFF00", fill_type="solid"
    )  # 黃色底

    # --- 建立總表 ---
    summary_sheet = wb.active
    summary_sheet.title = "班表總表"

    dates = sorted(schedule.keys())
    summary_sheet.cell(row=1, column=1, value="員工").font = bold_font
    for col, date in enumerate(dates, start=2):
        cell = summary_sheet.cell(row=1, column=col, value=date)
        cell.font = bold_font
        cell.alignment = center_alignment

    employee_row_map = {emp: i for i, emp in enumerate(employees, start=2)}
    for emp, row_idx in employee_row_map.items():
        summary_sheet.cell(row=row_idx, column=1, value=emp)

    for date_idx, date in enumerate(dates, start=2):
        on_leave_today = leave_map.get(date, set())
        for emp, row_idx in employee_row_map.items():
            cell = summary_sheet.cell(row=row_idx, column=date_idx)
            cell.alignment = center_alignment
            if emp in on_leave_today:
                cell.value = "休"
                cell.font = leave_font
                cell.fill = leave_fill
            else:
                assigned_shift = None
                for shift, assigned_employees in schedule[date].items():
                    if emp in assigned_employees:
                        assigned_shift = shift
                        break
                if assigned_shift:
                    cell.value = assigned_shift

    # --- 為每位員工建立個人班表 ---
    for emp in employees:
        emp_sheet = wb.create_sheet(title=f"{emp}個人班表")

        emp_sheet.cell(row=1, column=1, value="班別").font = bold_font
        for col, date in enumerate(dates, start=2):
            cell = emp_sheet.cell(row=1, column=col, value=date)
            cell.font = bold_font
            cell.alignment = center_alignment

        shift_row_map = {shift: i for i, shift in enumerate(shift_types, start=2)}
        for shift, row_idx in shift_row_map.items():
            emp_sheet.cell(row=row_idx, column=1, value=shift)

        for date_idx, date in enumerate(dates, start=2):
            on_leave_today = leave_map.get(date, set())
            if emp in on_leave_today:
                for row_idx in shift_row_map.values():
                    cell = emp_sheet.cell(row=row_idx, column=date_idx, value="休")
                    cell.font = leave_font
                    cell.alignment = center_alignment
                    cell.fill = leave_fill
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
    # 模擬的員工自然語言請求
    raw_requests = """
    To HR:
    我是員工1，我因為家裡有事，想要在 2025年7月15日 請假一天，謝謝。
    
    員工2:
    你好，我想請 2025-07-18，再麻煩了。
    
    員工3:
    老闆，7/15 我沒辦法上班喔。
    """

    # 步驟 1
    parsed_data = parse_requests_with_llm(raw_requests)

    # 步驟 2
    start_date = date(2025, 7, 14)
    end_date = date(2025, 7, 20)
    final_schedule, final_stats, leave_map = create_schedule(
        start_date, end_date, parsed_data
    )

    print("\n--- 初版班表已生成 ---")

    # 優先產生初版 Excel 檔案
    initial_filename = "shift_schedule_initial.xlsx"
    export_to_excel(
        final_schedule, EMPLOYEES, SHIFT_TYPES, leave_map, filename=initial_filename
    )
    print(f"已產生初版班表 Excel 檔案: {initial_filename}")

    # 新增：互動式調整迴圈
    while True:
        adjust = input(
            "\n是否需要進行手動調班？ (請輸入 'y' 進行調整，或直接按 Enter 結束): "
        ).lower()
        if adjust != "y":
            break

        emp_to_replace = input("請輸入要請假的員工姓名 (例如 員工87): ")
        if emp_to_replace not in EMPLOYEES:
            print("錯誤：找不到該員工。")
            continue

        date_to_replace = input("請輸入請假日期 (格式 YYYY-MM-DD): ")
        try:
            datetime.strptime(date_to_replace, "%Y-%m-%d")
        except ValueError:
            print("錯誤：日期格式不正確。")
            continue

        final_schedule, final_stats, leave_map, success = find_and_assign_replacement(
            final_schedule, final_stats, leave_map, emp_to_replace, date_to_replace
        )
        if success:
            print(f"班表已更新。")

    print("\n--- 所有調整已完成 ---")

    # 步驟 3
    summary_report = summarize_with_llm(final_schedule, final_stats)
    print("\n\n" + "=" * 50)
    print("     由 Gemini Pro 生成的最終排班摘要報告與通知")
    print("=" * 50)
    print(summary_report)

    # 步驟 4 (最終版)
    final_filename = "shift_schedule_final.xlsx"
    export_to_excel(
        final_schedule, EMPLOYEES, SHIFT_TYPES, leave_map, filename=final_filename
    )
    print(f"已產生最終版班表 Excel 檔案: {final_filename}")
