import openpyxl  # 导入 openpyxl 模块，用于处理 Excel 文件
import os  # 导入 os 模块，用于操作文件和目录
from openpyxl.styles import Alignment  # 从 openpyxl.styles 导入 Alignment 类，用于设置单元格对齐方式
import calendar  # 导入 calendar 模块，用于处理日期和时间
import logging  # 导入 logging 模块，用于记录日志信息

# 配置日志记录，设置日志级别为 INFO，格式为时间 - 日志级别 - 消息内容
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


def load_employees_from_txt(file_path):
    """从 TXT 文件加载员工考勤数据
    参数:
        file_path: 文件路径，指向包含员工数据的文本文件
    返回:
        一个字典，键为员工编号，值为对应的考勤字符串"""
    employees = {}  # 初始化一个空字典，用于存储员工编号和考勤数据
    if not os.path.exists(file_path):
        # 检查文件是否存在，如果不存在则记录错误信息并返回空字典
        logging.error(f"文件 {file_path} 不存在，请检查路径！")
        return employees
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            # 尝试以只读模式打开文件，编码格式为 UTF-8
            for line in file:
                line = line.strip()  # 去除每行的首尾空格
                if not line:
                    continue  # 跳过空行
                if ":" not in line:
                    # 如果行中没有冒号分隔符，记录警告信息并跳过
                    logging.warning(f"无效的员工数据格式: {line}")
                    continue
                try:
                    emp_id, attendance = line.split(':', 1)
                    # 按冒号分割为员工编号和考勤数据
                    emp_id = emp_id.strip()  # 去除编号中的空格
                    attendance = attendance.strip()  # 去除考勤数据中的空格
                    if not emp_id.isalnum() or not attendance:
                        # 检查员工编号是否为字母数字组合，考勤数据是否为空
                        logging.warning(f"无效的员工数据格式: {line}")
                        continue
                    employees[emp_id] = attendance  # 将解析出的员工编号和考勤数据加入字典
                except ValueError:
                    # 捕获分割过程中可能发生的错误并记录警告信息
                    logging.warning(f"无效的员工数据格式: {line}")
    except Exception as e:
        # 捕获文件读取过程中发生的任何异常并记录错误信息
        logging.error(f"加载员工数据时发生错误: {e}")
    return employees  # 返回包含员工数据的字典


def parse_attendance_input(input_str: str, days_in_month: int = 31):
    """解析快速输入字符串
    参数:
        input_str: 表示考勤信息的字符串，例如 "1, 2.1, 3.2+2, 4-6"
        days_in_month: 当前月份的天数，默认为 31 天
    返回:
        一个包含早班、晚班和加班信息的嵌套字典"""
    attendance = {"morning": {}, "afternoon": {}, "overtime": {}}  # 初始化嵌套字典

    def mark_attendance(day: int, period: int = None, overtime: float = None):
        """标记考勤数据
        参数:
            day: 天数，对应日期
            period: 考勤时间段，1 表示早班，2 表示晚班，默认标记全天
            overtime: 加班时长，默认为 None"""
        if 1 <= day <= days_in_month:
            if period == 1:
                attendance["morning"][day] = "\u2713"  # 早班出勤标记
            elif period == 2:
                attendance["afternoon"][day] = "\u2713"  # 晚班出勤标记
            else:
                attendance["morning"][day] = "\u2713"
                attendance["afternoon"][day] = "\u2713"  # 全天出勤

            if overtime is not None:
                attendance["overtime"][day] = overtime  # 记录加班时间

    for part in input_str.split(","):
        part = part.strip()  # 去除每部分的首尾空格
        if not part:
            continue  # 跳过空部分
        try:
            if "+" in part and "." in part:
                # 如果部分同时包含 "+" 和 "."，表示 "日.时段+加班"
                date_period, overtime = part.split("+")
                day, period = map(int, date_period.split("."))
                mark_attendance(day, period, float(overtime))
            elif "-" in part:
                # 如果部分包含 "-"，表示 "起始日-结束日"
                start, end = map(int, part.split("-"))
                for day in range(start, end + 1):
                    mark_attendance(day)
            elif "." in part:
                # 如果部分包含 "."，表示 "日.时段"
                day, period = map(int, part.split("."))
                mark_attendance(day, period)
            elif "+" in part:
                # 如果部分包含 "+"，表示 "日+加班"
                day, overtime = part.split("+")
                mark_attendance(int(day), None, float(overtime))
            else:
                # 否则部分表示单独的日期
                mark_attendance(int(part))
        except ValueError:
            logging.warning(
                f"无效的格式: {part}。预期格式包括 'day', 'day.period', 'day+overtime', 'start-end' 等。")
            continue

    # 填充默认值，将未标记的日期标记为未出勤
    for day in range(1, days_in_month + 1):
        attendance["morning"].setdefault(day, "\u2717")
        attendance["afternoon"].setdefault(day, "\u2717")
        attendance["overtime"].setdefault(day, "")

    return attendance  # 返回解析后的考勤数据字典


def create_attendance_template(filename: str, days_in_month: int = 31):
    """创建考勤模板
    参数:
        filename: 模板保存的文件路径
        days_in_month: 当前月份的天数"""
    wb = openpyxl.Workbook()  # 创建一个新的工作簿
    ws = wb.active  # 获取活动工作表
    ws.title = "考勤表"  # 设置表格标题

    ws.cell(row=1, column=1, value="员工编号")  # 填写标题行的第一列
    for col in range(2, days_in_month + 2):
        ws.cell(row=1, column=col, value=col - 1)  # 按日期填写列标题
    for cell in ws[1]:
        cell.alignment = Alignment(horizontal="center", vertical="center")  # 首行标题进行居中处理

    set_column_width(ws, width=3.6, exclude_first_column=True)  # 设置列宽
    wb.save(filename)  # 保存工作簿
    logging.info(f"考勤模板已生成: {filename}")  # 记录模板生成的信息


def set_column_width(ws, width=3.6, exclude_first_column=False):
    """设置列宽，并将内容居中
    参数:
        ws: 工作表对象
        width: 列宽的数值
        exclude_first_column: 是否排除第一列"""
    for col_idx, col_cells in enumerate(ws.columns, start=1):
        if exclude_first_column and col_idx == 1:
            continue
        col_letter = col_cells[0].column_letter  # 获取列字母
        ws.column_dimensions[col_letter].width = width  # 设置列宽

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")  # 居中对齐


def fill_attendance(filename: str, employee_data: dict, days_in_month: int = 31):
    """根据员工考勤数据填写表格，并增加出勤统计列
    参数:
        filename: 表格文件路径
        employee_data: 包含员工编号及考勤数据的字典
        days_in_month: 当前月份的天数"""
    if not os.path.exists(filename):
        # 如果文件不存在，先创建考勤模板
        create_attendance_template(filename, days_in_month)

    wb = openpyxl.load_workbook(filename)  # 打开现有的 Excel 工作簿
    ws = wb.active  # 获取活动工作表

    extra_columns = [("出勤天数", 0), ("加班计时", 1), ("合计工数", 2)]
    base_col = days_in_month + 2  # 额外列的起始列号
    for title, offset in extra_columns:
        col_idx = base_col + offset
        ws.cell(row=1, column=col_idx, value=title)  # 增加新的统计列标题

    start_row = 2  # 设置起始行，从第二行开始填写数据

    for emp_id, attendance_str in employee_data.items():
        ws.cell(row=start_row, column=1, value=emp_id)  # 填写员工编号到第一列

        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row + 2, end_column=1)
        cell = ws.cell(row=start_row, column=1)
        cell.alignment = Alignment(horizontal="center", vertical="center")  # 将编号单元格内容居中对齐

        attendance = parse_attendance_input(attendance_str, days_in_month)  # 解析员工的考勤数据字符串

        total_morning = 0
        total_afternoon = 0
        total_overtime = 0
        for col in range(2, days_in_month + 2):
            day = col - 1  # 列号减去 1 对应日期
            morning = attendance["morning"].get(day, "\u2717")
            ws.cell(row=start_row, column=col, value=morning)  # 填写早班考勤
            if morning == "\u2713":
                total_morning += 0.5

            afternoon = attendance["afternoon"].get(day, "\u2717")
            ws.cell(row=start_row + 1, column=col, value=afternoon)  # 填写晚班考勤
            if afternoon == "\u2713":
                total_afternoon += 0.5

            overtime = attendance["overtime"].get(day, "")
            if overtime:
                ws.cell(row=start_row + 2, column=col, value=overtime)  # 填写加班数据
                total_overtime += float(overtime)

        attendance_days = total_morning + total_afternoon
        overtime_days = total_overtime / 6
        total_days = attendance_days + overtime_days

        ws.merge_cells(start_row=start_row, start_column=base_col, end_row=start_row + 1, end_column=base_col)
        attendance_cell = ws.cell(row=start_row, column=base_col, value=attendance_days)
        attendance_cell.alignment = Alignment(horizontal="center", vertical="center")  # 填写出勤天数（前两行合并，显示总出勤天数）

        overtime_days_cell = ws.cell(row=start_row + 2, column=base_col, value=overtime_days)
        overtime_days_cell.alignment = Alignment(horizontal="center", vertical="center")  # 填写加班天数（出勤天数第3行显示加班时长换算的出勤天数）
        ws.cell(row=start_row + 2, column=base_col, value=overtime_days).number_format = "0.0"

        ws.merge_cells(start_row=start_row, start_column=base_col + 1, end_row=start_row + 2, end_column=base_col + 1)
        overtime_cell = ws.cell(row=start_row, column=base_col + 1, value=total_overtime)
        overtime_cell.alignment = Alignment(horizontal="center", vertical="center")  # 填写加班计时（所有加班小时数合并）

        ws.merge_cells(start_row=start_row, start_column=base_col + 2, end_row=start_row + 2, end_column=base_col + 2)
        total_cell = ws.cell(row=start_row, column=base_col + 2, value=total_days)
        total_cell.alignment = Alignment(horizontal="center", vertical="center")  # 填写合计工数（出勤天数 + 加班天数，合并）

        start_row += 3  # 每个员工占用三行，因此起始行向下移动三行

    wb.save(filename)  # 保存更新后的工作簿
    logging.info(f"考勤表已更新并保存为: {filename}")  # 记录考勤表更新的信息


if __name__ == "__main__":
    year, month = 2024, 12  # 动态获取年份和月份对应的天数
    days_in_month = calendar.monthrange(year, month)[1]  # 获取月份的天数

    employee_file = os.path.join(os.path.dirname(__file__), "employees.txt")  # 加载员工数据文件
    employees = load_employees_from_txt(employee_file)  # 从文件中加载员工考勤数据

    output_file = os.path.join(os.path.dirname(__file__), "../考勤表.xlsx")  # 输出文件名

    fill_attendance(output_file, employees, days_in_month)  # 填充考勤数据
