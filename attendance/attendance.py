import openpyxl
import os
from openpyxl.styles import Alignment
import calendar
import logging

# 配置日志记录
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def load_employees_from_txt(file_path):
    """从 TXT 文件加载员工考勤数据"""
    employees = {}
    if not os.path.exists(file_path):
        logging.error(f"文件 {file_path} 不存在，请检查路径！")
        return employees
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            for line in file:
                line = line.strip()
                if not line:
                    continue
                if ":" not in line:
                    logging.warning(f"无效的员工数据格式: {line}")
                    continue
                try:
                    emp_id, attendance = line.split(':', 1)
                    emp_id = emp_id.strip()
                    attendance = attendance.strip()
                    if not emp_id.isalnum() or not attendance:
                        logging.warning(f"无效的员工数据格式: {line}")
                        continue
                    employees[emp_id] = attendance
                except ValueError:
                    logging.warning(f"无效的员工数据格式: {line}")
    except Exception as e:
        logging.error(f"加载员工数据时发生错误: {e}")
    return employees

def parse_attendance_input(input_str: str, days_in_month: int = 31):
    """解析快速输入字符串"""
    attendance = {"morning": {}, "afternoon": {}, "overtime": {}}

    def mark_attendance(day: int, period: int = None, overtime: float = None):
        """标记考勤数据"""
        if 1 <= day <= days_in_month:
            if period == 1:
                attendance["morning"][day] = "\u2713"
            elif period == 2:
                attendance["afternoon"][day] = "\u2713"
            else:
                attendance["morning"][day] = "\u2713"
                attendance["afternoon"][day] = "\u2713"

            if overtime is not None:
                attendance["overtime"][day] = overtime

    for part in input_str.split(","):
        part = part.strip()
        if not part:
            continue
        try:
            if "+" in part and "." in part:
                date_period, overtime = part.split("+")
                day, period = map(int, date_period.split("."))
                mark_attendance(day, period, float(overtime))
            elif "-" in part:
                start, end = map(int, part.split("-"))
                for day in range(start, end + 1):
                    mark_attendance(day)
            elif "." in part:
                day, period = map(int, part.split("."))
                mark_attendance(day, period)
            elif "+" in part:
                day, overtime = part.split("+")
                mark_attendance(int(day), None, float(overtime))
            else:
                mark_attendance(int(part))
        except ValueError:
            logging.warning(
                f"Invalid format: {part}. Expected formats include 'day', 'day.period', 'day+overtime', 'start-end', etc.")
            continue

    for day in range(1, days_in_month + 1):
        attendance["morning"].setdefault(day, "\u2717")
        attendance["afternoon"].setdefault(day, "\u2717")
        attendance["overtime"].setdefault(day, "")

    return attendance

def create_attendance_template(filename: str, days_in_month: int = 31):
    """创建考勤模板"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "考勤表"

    ws.cell(row=1, column=1, value="员工编号")
    for col in range(2, days_in_month + 2):
        ws.cell(row=1, column=col, value=col - 1)

    set_column_width(ws, width=3.6, exclude_first_column=True)
    wb.save(filename)
    logging.info(f"考勤模板已生成: {filename}")

def set_column_width(ws, width=3.6, exclude_first_column=False):
    """设置列宽，并将内容居中"""
    for col_idx, col_cells in enumerate(ws.columns, start=1):
        if exclude_first_column and col_idx == 1:
            continue
        col_letter = col_cells[0].column_letter
        ws.column_dimensions[col_letter].width = width

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")


def fill_attendance(filename: str, employee_data: dict, days_in_month: int = 31):
    """根据员工考勤数据填写表格"""
    # 如果文件不存在，则创建考勤模板
    if not os.path.exists(filename):
        create_attendance_template(filename, days_in_month)

    wb = openpyxl.load_workbook(filename)
    ws = wb.active

    start_row = 2

    # 检查是否已经设置了列宽（通过查看第 1 行是否有数据）
    if ws.cell(row=1, column=2).value is None:
        set_column_width(ws, width=3.6, exclude_first_column=True)  # 仅在首次填充时设置列宽

    for emp_id, attendance_str in employee_data.items():
        # 填写员工编号
        ws.cell(row=start_row, column=1, value=emp_id)

        # 合并员工编号行的 3 行，并居中
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row + 2, end_column=1)
        cell = ws.cell(row=start_row, column=1)
        cell.alignment = Alignment(horizontal="center", vertical="center")

        attendance = parse_attendance_input(attendance_str, days_in_month)
        for col in range(2, days_in_month + 2):
            day = col - 1
            ws.cell(row=start_row, column=col, value=attendance["morning"].get(day, "\u2717"))
            ws.cell(row=start_row + 1, column=col, value=attendance["afternoon"].get(day, "\u2717"))
            overtime = attendance["overtime"].get(day, "")
            if overtime:
                ws.cell(row=start_row + 2, column=col, value=overtime)

        start_row += 3

    wb.save(filename)
    logging.info(f"考勤表已更新并保存为: {filename}")


if __name__ == "__main__":
    # 动态获取月份天数
    year, month = 2024, 12
    days_in_month = calendar.monthrange(year, month)[1]

    # 加载员工数据文件
    employee_file = os.path.join(os.path.dirname(__file__), "employees.txt")
    employees = load_employees_from_txt(employee_file)

    # 输出文件名
    output_file = os.path.join(os.path.dirname(__file__), "../考勤表.xlsx")

    # 填充考勤数据
    fill_attendance(output_file, employees, days_in_month)
