import openpyxl  # 导入 openpyxl 模块，用于处理 Excel 文件
import os  # 导入 os 模块，用于操作文件和目录
import calendar  # 导入 calendar 模块，用于处理日期和时间
import logging  # 导入 logging 模块，用于记录日志信息
import json
import mysql.connector
from openpyxl.styles import Alignment, Font  # 从 openpyxl.styles 导入 Alignment,Font 类，用于设置单元格对齐方式及字体
from openpyxl.utils import get_column_letter
from openpyxl.styles import Border, Side
from openpyxl.worksheet.page import PageMargins

# 配置日志记录，设置日志级别为 INFO，格式为时间 - 日志级别 - 消息内容
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


def sort_txt_by_number(input_file, output_file):
    try:
        # 检查文件是否存在
        if not os.path.exists(input_file):
            print(f"输入文件不存在：{input_file}")
            return

        with open(input_file, 'r', encoding='utf-8') as file:
            lines = [line.strip() for line in file if line.strip()]
            '''
            # for line in file
            # 遍历 file 文件对象的每一行，逐行读取内容。
            # file 是通过 open 打开的文件对象，支持迭代访问，每次返回一行字符串（包括换行符 \n）。

            # line.strip()
            # strip() 是字符串的一个方法，用于移除字符串开头和结尾的所有空白字符（如空格、制表符 \t 和换行符 \n）。
            # 如果一行内容为 " example text \n"，调用 strip() 后会得到 "example text"。

            # if line.strip()
            # 这是一个条件过滤器，表示只有在 line.strip() 的结果为真时（非空字符串），才会将该行加入列表。
            # 空行（如只有换行符或空格）经过 strip() 会变成空字符串 ""，在布尔上下文中为 False，因此会被过滤掉。

            # [... for ... if ...]
            # 这是一个列表生成式，用于构建一个列表。
            # 它会将符合条件的 line.strip() 的结果依次加入到列表中
            '''

        sorted_lines = sorted(lines, key=lambda x: int(x.split(':')[0]))
        '''
        # 1.sorted(lines)
        # sorted是Python的内置函数，用于对可迭代对象进行排序。
        # 它会返回一个新的排序后的列表，不会修改原始列表lines。
        # 语法：sorted(iterable, key=None, reverse=False)
        #     iterable：要排序的对象，这里是lines列表。
        #     key：指定排序的规则（一个函数），这里通过lambda 表达式定义。
        #     reverse：默认False，表示升序排序。设置为True则降序
        #
        # 2.key = lambda x: int(x.split(':')[0])
        # key是一个函数，指定排序的依据。
        # 这里使用lambda x 定义了一个匿名函数，x是列表中每个元素的值。
        #
        # 3.拆解lambda x: int(x.split(':')[0])
        #     1 x.split(':')
        #
        #         split(':')是字符串的一个方法，用于将字符串按冒号: 分割成多个部分。
        #         返回值是一个列表，其中每部分是分割后的子字符串。
        #         示例："1:苹果".split(':') → ["1", "苹果"]
        #     2 x.split(':')[0]
        #         取分割后列表的第一个元素，即冒号前的部分。
        #         示例："1:苹果".split(':')[0] → "1"
        #     3 int(x.split(':')[0])
        #         将冒号前的部分从字符串转换为整数，用作排序的依据。
        #         示例：int("1") → 1
        #     4 lambda x 是 Python 中用于定义 匿名函数 的语法。它的功能类似于 def 关键字创建的普通函数，但更简洁。
        #         lambda 参数: 表达式
        #         lambda：关键字，用于声明一个匿名函数。
        #         参数：函数的输入，可以是一个或多个，多个参数用逗号分隔。
        #         表达式：函数的返回值，必须是一个单一表达式，不能包含复杂语句。
        #         匿名函数的特点
        #         没有名称：lambda 定义的函数是匿名的，常用在需要临时函数的场合。
        #         内联简洁：适合用在较简单的情况下，定义和使用往往写在同一行。
        #         自动返回：lambda 的表达式部分会自动作为返回值，无需使用 return。
        '''

        with open(output_file, 'w', encoding='utf-8') as file:
            file.write('\n'.join(sorted_lines))
        '''
        # 代码解析
        # 1.with open(output_file, 'w', encoding='utf-8') as file
        #     open(output_file, 'w', encoding='utf-8')
        #     打开（或创建）一个文件进行写操作。
        #     参数说明：
        #         output_file：指定的文件路径。
        #         'w'：表示写模式，会覆盖文件内容。如果文件不存在，会自动创建一个新文件。
        #         encoding = 'utf-8'：指定使用UTF - 8编码写入内容，保证对非ASCII字符（如中文）兼容。
        #     with 关键字
        #     with 是上下文管理器，负责在操作完成后自动关闭文件，无需手动调用 file.close()，即使在程序异常时也会安全关闭文件。
        # 2.file.write('\n'.join(sorted_lines))
        #     '\n'.join(sorted_lines)
        #         将sorted_lines列表中的元素拼接成一个字符串，每个元素之间用换行符 \n分隔。
        #         示例：
        #         sorted_lines = ["1:香蕉", "2:橙子", "3:苹果"]
        #         result = '\n'.join(sorted_lines)
        #         print(result)
        #         # 输出：
        #         # 1:香蕉
        #         # 2:橙子
        #         # 3:苹果
        #     file.write(...)
        #         将拼接后的字符串写入文件。
        #         如果sorted_lines是上例中的内容，文件的最终内容会是：
        #         1: 香蕉
        #         2: 橙子
        #         3: 苹果
        # 优点
        # 自动关闭文件：使用with 块，无论文件操作是否成功，都会自动释放资源。
        # 简洁高效：直接将排序后的列表转换为字符串并一次性写入文件。
        # 跨平台支持：指定编码为UTF - 8，兼容各种操作系统和语言环境。
        # 注意事项
        # 覆盖风险：如果output_file已存在，其内容会被覆盖。如果需要追加内容，可以将'w'替换为'a'（追加模式）。
        # 编码问题：如果内容中包含特殊字符，确保文件编码和操作系统的默认编码兼容。UTF-8是推荐选择。
        # 数据格式一致性：确保sorted_lines的内容已经按照所需的格式排序并去除了无效数据（如空行）。
        '''
        print(f"文件已成功排序并保存到 {output_file}")
    except Exception as e:
        print(f"发生错误：{e}")


# 自动获取当前脚本所在目录，并拼接文件路径
current_dir = os.path.dirname(os.path.abspath(__file__))
input_file = os.path.join(current_dir, 'input.txt')
output_file = os.path.join(current_dir, 'employees.txt')
sort_txt_by_number(input_file, output_file)
'''
    # 逐行解析
    # 1. current_dir = os.path.dirname(os.path.abspath(__file__))
    # os.path.abspath(__file__)：
    # 获取当前脚本文件的绝对路径（包含文件名）。
    # 例如，假设脚本位于 /home/user/project/script.py，则 os.path.abspath(__file__) 的结果是 /home/user/project/script.py。
    # os.path.dirname()：
    # 获取文件所在的目录路径。
    # 结合上例，os.path.dirname("/home/user/project/script.py") 的结果是 /home/user/project。
    # current_dir：
    # 最终保存了当前脚本所在的目录路径。

    # 2. input_file = os.path.join(current_dir, 'input.txt')
    # os.path.join()：
    # 将目录路径 current_dir 与文件名 input.txt 拼接成一个完整的路径。
    # 例如，如果 current_dir 是 /home/user/project，则结果是 /home/user/project/input.txt。
    # input_file：
    # 保存了待处理文件 input.txt 的完整路径。

    # 3. output_file = os.path.join(current_dir, 'employees.txt')
    # 功能与上面类似，只不过目标文件是 sorted_output.txt。
    # output_file 保存了排序后输出文件的完整路径。

    # 4. sort_txt_by_number(input_file, output_file)
    # 调用定义好的函数 sort_txt_by_number，并传入两个参数：
    # input_file：表示要读取并排序的输入文件路径。
    # output_file：表示排序后结果要保存的输出文件路径。
'''


def load_employees_from_txt(file_path):
    """从 TXT 文件加载员工考勤数据
    参数:
        file_path: 文件路径，指向包含员工数据的文本文件
    返回:
        一部字典，键为员工编号，值为对应的考勤字符串"""
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
                    # line.split(':') 是 Python 字符串的 split 方法，用于根据指定的分隔符（这里是 :）将字符串分割成一个列表。
                    # 第二个参数 1 是 maxsplit，表示最多分割 1 次，即字符串被拆分为最多 2 个部分。返回值是一个包含最多两个元素的列表
                    # emp_id, attendance 是解包赋值，将分割后的列表中第一个元素赋值给 emp_id，第二个元素赋值给 attendance
                    emp_id = emp_id.strip()  # 去除编号中的空格
                    attendance = attendance.strip()  # 去除考勤数据中的空格
                    if not emp_id.isalnum() or not attendance:
                        # 检查员工编号是否为字母数字组合，考勤数据是否为空
                        # emp_id.isalnum() 是一个字符串方法，用于检查字符串是否仅由字母和数字组成（即是否是“字母数字”字符串），且不能为空。
                        # not emp_id.isalnum() 表示字符串 emp_id 不是由纯字母和数字组成。如果 emp_id 包含空格、特殊字符（例如 @, # 等），或者为空字符串，条件为真
                        # not attendance检查变量 attendance 是否为假值。假值包括：""（空字符串）、None、False等。not attendance 表示 attendance 是空字符串或其他假值
                        logging.warning(f"无效的员工数据格式: {line}")
                        continue
                    employees[emp_id] = attendance  # 将解析出的员工编号和考勤数据加入字典
                    # 这行代码的作用是将 emp_id 作为键（Key），将 attendance 作为值（Value），存储到字典 employees 中
                    # employees：这是一个字典（dict）对象，用于存储键值对。
                    # emp_id：表示一个员工的唯一标识（如员工编号）。
                    # attendance：表示该员工的考勤信息（如“Present”或“Absent”）
                    # 这行代码的作用是将 emp_id 和 attendance 作为键值对加入到字典中。如果 emp_id 已经存在于字典中，则会更新其对应的值为新的 attendance
                except ValueError:
                    # 捕获分割过程中可能发生的错误并记录警告信息
                    logging.warning(f"无效的员工数据格式: {line}")
    except Exception as e:
        # 捕获文件读取过程中发生的任何异常并记录错误信息
        logging.error(f"加载员工数据时发生错误: {e}")

    return employees  # 返回包含员工数据的字典 employees[emp_id] = attendance


def parse_attendance_input(input_str: str, emp_ids: list, days_in_month: int = 31):
    """解析快速输入字符串，返回按员工编号组织的考勤数据字典
    参数:
        input_str: 表示考勤信息的字符串，例如 "1, 2.1, 3.2+2, 4-6"
        employee_ids: 员工编号的列表，用于将考勤信息按员工编号组织
        days_in_month: 当前月份的天数，默认为 31 天
    返回:
        一个按员工编号组织的嵌套字典, 格式如下：
        {
            emp_id: {
                "morning": {1: "✓", 2: "✗", 3: "✓", ...},
                "afternoon": {1: "✓", 2: "✓", 3: "✗", ...},
                "overtime": {1: 2.5, 2: "", 3: 1.0, ...}
            },
            ...
        }
    """

    # 创建一个空的字典，用于存储每个员工的考勤数据
    attendance_data = {emp_id: {"morning": {}, "afternoon": {}, "overtime": {}} for emp_id in emp_ids}

    def mark_attendance(emp_id, day: int, period: int = None, overtime: float = None):
        """标记某个员工的考勤数据
        参数:
            emp_id: 员工编号
            day: 天数，对应日期
            period: 考勤时间段，1 表示早班，2 表示晚班，默认标记全天
            overtime: 加班时长，默认为 None
        """
        if emp_id not in attendance_data:
            return  # 如果员工编号不在考勤数据中，则跳过

        if 1 <= day <= days_in_month:
            if period == 1:
                attendance_data[emp_id]["morning"][day] = "\u2713"  # 早班出勤标记
            elif period == 2:
                attendance_data[emp_id]["afternoon"][day] = "\u2713"  # 晚班出勤标记
            else:
                attendance_data[emp_id]["morning"][day] = "\u2713"
                attendance_data[emp_id]["afternoon"][day] = "\u2713"  # 全天出勤

            if overtime is not None:
                attendance_data[emp_id]["overtime"][day] = overtime  # 记录加班时间

    # 解析输入字符串，并标记考勤
    for part in input_str.split(","):
        part = part.strip()  # 去除每部分的首尾空格
        if not part:
            continue  # 跳过空部分
        try:
            if "+" in part and "." in part:
                # "日.时段+加班"格式
                date_period, overtime = part.split("+")
                day, period = map(int, date_period.split("."))
                overtime = float(overtime)
                for emp_id in emp_ids:
                    mark_attendance(emp_id, day, period, overtime)
            elif "-" in part:
                # "起始日-结束日"格式
                start, end = map(int, part.split("-"))
                for day in range(start, end + 1):
                    for emp_id in emp_ids:
                        mark_attendance(emp_id, day)
            elif "." in part:
                # "日.时段"格式
                day, period = map(int, part.split("."))
                for emp_id in emp_ids:
                    mark_attendance(emp_id, day, period)
            elif "+" in part:
                # "日+加班"格式
                day, overtime = part.split("+")
                overtime = float(overtime)
                for emp_id in emp_ids:
                    mark_attendance(emp_id, int(day), None, overtime)
            else:
                # 单独日期
                for emp_id in emp_ids:
                    mark_attendance(emp_id, int(part))
        except ValueError:
            logging.warning(
                f"无效的格式: {part}。预期格式包括 'day', 'day.period', 'day+overtime', 'start-end' 等。")
            continue

    # 填充默认值，将未标记的日期标记为未出勤
    for emp_id in emp_ids:
        for day in range(1, days_in_month + 1):
            attendance_data[emp_id]["morning"].setdefault(day, "\u2717")
            attendance_data[emp_id]["afternoon"].setdefault(day, "\u2717")
            attendance_data[emp_id]["overtime"].setdefault(day, "")

    return attendance_data  # 返回按员工编号组织的考勤数据字典


def load_db_config(config_file):
    if not os.path.exists(config_file):
        print(f"配置文件 {config_file} 不存在！")
        return None
    try:
        with open(config_file, 'r', encoding='utf-8') as file:
            config = json.load(file)
        if isinstance(config, dict):
            return config
        else:
            print(f"配置文件格式错误，应该是一个字典：{config}")
            return None
    except Exception as e:
        print(f"读取配置文件时出错：{e}")
        return None


def get_employee_info_from_mysql(config, emp_ids):
    """根据员工工号列表从 MySQL 获取员工信息，得到一部键为员工工号的字典"""
    if not config:
        print("数据库配置无效，无法连接数据库。")
        return {}

    if not emp_ids:
        print("员工ID列表为空，无法执行查询。")
        return {}

    try:
        # 连接数据库
        conn = mysql.connector.connect(
            host=config['host'],
            user=config['user'],
            password=config['password'],
            database=config['database']
        )

        # 自动管理游标和连接
        with conn.cursor(dictionary=True) as cursor:
            # 创建游标对象 cursor
            # 参数 dictionary=True 指定查询结果以字典形式返回，而非默认的元组形式
            # 使用 with 语句确保游标在代码块结束后自动关闭，避免资源泄露

            # 构建参数化查询
            placeholders = ", ".join(["%s"] * len(emp_ids))
            # 根据 emp_ids 的长度，动态生成 SQL 占位符字符串。
            # 假如 emp_ids = [101, 102]，生成的字符串为 "%s, %s"
            # 使用占位符 %s 实现参数化查询，避免直接拼接字符串，防止 SQL 注入风险

            query = f"SELECT emp_id, job_type, name, unit_price FROM employees WHERE emp_id IN ({placeholders})"
            # 构建查询语句，将占位符插入 WHERE emp_id IN (...) 中,保证查询语句安全且易于扩展保证查询语句安全且易于扩展

            # 执行查询
            cursor.execute(query, emp_ids)
            # query, emp_ids都为参数
            # 执行 SQL 查询，emp_ids 中的值会替换占位符 %s,通过参数化查询方式，确保数据被安全转义，防止 SQL 注入

            # 获取查询结果
            employee_data = cursor.fetchall()
            # 使用 fetchall() 获取查询结果，返回一个包含所有记录的列表

        # 关闭数据库连接
        conn.close()

        # 打印查询结果
        logging.info(f"查询结果: {employee_data}")

        # 将结果转换为字典形式
        employee_info = {
            emp['emp_id']: {'job_type': emp['job_type'], 'name': emp['name'], 'unit_price': emp['unit_price']}
            for emp in employee_data
        }
        # emp 是字典推导式中的临时变量，不需要在上下文中单独定义。
        # 它的值来自 for emp in employee_data，每次迭代时会被动态赋值为 employee_data 的当前元素。
        # 列表推导式的基本语法：[expression for item in iterable if condition]
        # 它的作用域仅限于字典推导式或循环语句中，执行完成后就会释放

        # 打印查询结果
        logging.info(f"填充后的员工信息字典: {employee_info}")

        return employee_info

    except mysql.connector.Error as err:
        print(f"数据库连接错误: {err}")
        return {}


def set_column_widths(ws):  # 设置列宽
    # 设置第1列(A), 第2列(B), 第3列(C)的列宽
    ws.column_dimensions["A"].width = 3  # 第1列(A)宽度为3
    ws.column_dimensions["B"].width = 7  # 第2列(B)宽度为7
    ws.column_dimensions["C"].width = 13  # 第3列(C)宽度为13

    # 设置第4列到第34列(D到AH)的列宽为3.6
    for col in range(4, 35):
        col_letter = get_column_letter(col)  # 自动处理列号到字母的转换
        ws.column_dimensions[col_letter].width = 3.6

    # 设置第35列到第38列(AI到AL)的列宽为6.8
    for col in range(35, 39):
        col_letter = get_column_letter(col)  # 自动处理列号到字母的转换
        ws.column_dimensions[col_letter].width = 6.8

    # 设置第39列(AM)的列宽为10.5
    ws.column_dimensions["AM"].width = 10.5

    # 设置第40列(AN)的列宽为22.5
    ws.column_dimensions["AN"].width = 22.5


def create_attendance_template(attendance_data, filename: str, days_in_month: int = 31, ):
    """创建考勤模板
    参数:
        filename: 模板保存的文件路径
        days_in_month: 当前月份的天数
        :param attendance_data: 解析txt而来的字典,用于设置模板框线
        :type filename: object"""
    wb = openpyxl.Workbook()  # 创建一个新的工作簿
    ws = wb.active  # 获取活动工作表
    ws.title = "考勤表"  # 设置工作表表格标题

    # 合并第1行并填写标题
    ws.merge_cells("A1:AN1")  # 合并第1行的 A:AN 单元格
    header_cell1 = ws.cell(row=1, column=1, value="考勤表")  # 设置合并单元格的内容为 "考勤表"
    header_cell1.font = Font(size=28, bold=True)  # 设置字体为 28 号
    ws.row_dimensions[1].height = 40  # 设置行高为 40
    header_cell1.alignment = Alignment(horizontal="center", vertical="center")

    # 合并第2行填写项目名称:威尔斯
    ws.merge_cells("A2:AK2")  # 合并第2行的 A:AK 单元格
    header_cell2 = ws.cell(row=2, column=1, value="项目名称：威尔斯 ")  # 设置合并单元格的内容为 "项目名称：威尔斯 "
    header_cell2.alignment = Alignment(horizontal="left", vertical="center")  # 设置居左对齐

    # 合并第2行填写日期
    ws.merge_cells("AL2:AN2")  # 合并第2行的 AL:AN 单元格
    header_text = "2025年1月"  # 默认固定的年月
    header_cell3 = ws.cell(row=2, column=38, value=header_text)  # 获取合并单元格的起始单元格并设置内容
    header_cell3.alignment = Alignment(horizontal="center", vertical="center")  # 设置内容居中

    # 第三行填写数据
    headers = ['序\n号', '工号', '']
    headers.extend([str(day) for day in range(1, days_in_month + 1)])  # 日期列
    headers.extend(["出勤\n天数", "加班\n计时", "合计\n工数", "人工\n单价", "合计\n工资", "签名"])  # 新增列
    for col, header in enumerate(headers, start=1):
        # enumerate 是 Python 中的一个内置函数，用于将一个可迭代对象（如列表、元组等）转换为一个枚举对象，返回一个由索引和元素组成的元组。
        # headers 是一个包含列标题的列表。
        # enumerate(headers, start=1) 表示从 1 开始为 headers 列表中的每个元素提供一个索引。col 表示列索引，header 表示 headers 列表中的当前元素（即列标题）
        cell = ws.cell(row=3, column=col, value=header)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)  # 设置单元格内容居中及自动换行
    ws.row_dimensions[3].height = 34  # 设置行高为 34

    # 第三行填写数据
    header_cell4 = ws.cell(row=3, column=3, value='       日期\n姓名')  # 获取合并单元格的起始单元格并设置内容
    header_cell4.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)  # 设置内容居中及自动换行

    # 调用设置列宽函数
    set_column_widths(ws)

    # 设置表格边框线
    # 创建边框样式（四个方向的线）
    thin_border = Border(
        left=Side(border_style="thin", color="000000"),
        right=Side(border_style="thin", color="000000"),
        top=Side(border_style="thin", color="000000"),
        bottom=Side(border_style="thin", color="000000")
    )
    # 设置要添加边框的单元格范围
    start_row = 1
    end_row = len(attendance_data) * 3 + 3
    start_column = 1
    end_column = days_in_month + 9
    # 给指定区域的所有单元格添加边框
    for row in ws.iter_rows(min_row=start_row, max_row=end_row, min_col=start_column, max_col=end_column):
        for cell in row:
            cell.border = thin_border

    # 设置页边距
    ws.page_margins = PageMargins(
        left=0.5,  # 左边距（英寸）
        right=0.5,  # 右边距（英寸）
        top=0.5,  # 上边距（英寸）
        bottom=0.5,  # 下边距（英寸）
        header=0.3,  # 页眉到页面顶部的距离（英寸）
        footer=0.3  # 页脚到页面底部的距离（英寸）
    )

    # 设置页眉和页脚
    ws.oddHeader.left.text = ""
    ws.oddHeader.center.text = ""
    ws.oddHeader.right.text = ""

    ws.oddFooter.left.text = "备注：出勤：✔ 旷工：✘"
    ws.oddFooter.center.text = "考勤员：        项目经理：   "
    ws.oddFooter.right.text = "第&P页，共&N页"

    # 设置打印方向为横向或纵向
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE  # 横向
    # ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT  # 纵向

    wb.save(filename)  # 保存工作簿
    logging.info(f"考勤模板已生成: {filename}")  # 记录模板生成的信息


def fill_attendance(filename: str, attendance_data: dict, employee_info: dict, emp_ids, days_in_month: int = 31):
    """填充考勤信息到员工数据中"""

    logging.info(f"考勤数据: {attendance_data}")
    logging.info(f"员工信息: {employee_info}")

    if not os.path.exists(filename):
        create_attendance_template(attendance_data, filename, days_in_month)

    wb = openpyxl.load_workbook(filename)
    ws = wb.active

    start_row = 4
    start_value = 1  # 自增数字的起始值

    for emp_id, attendance in attendance_data.items():
        logging.info(f"正在处理员工 {emp_id} 的考勤数据...")
        # 写入序号并合并单元格

        # 写入序号并设置单元格格式
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row + 2, end_column=1)
        cell0 = ws.cell(row=start_row, column=1, value=start_value)
        cell0.alignment = Alignment(horizontal="center", vertical="center")
        # 更新序号
        start_value += 1

        # 写入员工工号并合        并单元格数据
        ws.merge_cells(start_row=start_row, start_column=2, end_row=start_row + 2, end_column=2)
        cell1 = ws.cell(row=start_row, column=2, value=emp_id)  # 写入编号到左上角单元格
        cell1.alignment = Alignment(horizontal="center", vertical="center")  # 设置单元格内容居中

        # 填充每天的考勤数据
        total_morning, total_afternoon, total_overtime = 0, 0, 0
        for col in range(2, days_in_month + 2):
            day = col - 1
            morning = attendance["morning"].get(day, "\u2717")
            afternoon = attendance["afternoon"].get(day, "\u2717")
            overtime = attendance["overtime"].get(day, "")

            cell2 = ws.cell(row=start_row, column=col + 2, value=morning)
            cell2.alignment = Alignment(horizontal="center", vertical="center")

            cell3 = ws.cell(row=start_row + 1, column=col + 2, value=afternoon)
            cell3.alignment = Alignment(horizontal="center", vertical="center")

            if overtime:
                cell4 = ws.cell(row=start_row + 2, column=col + 2, value=overtime)
                cell4.alignment = Alignment(horizontal="center", vertical="center")

            total_morning += 0.5 if morning == "\u2713" else 0
            total_afternoon += 0.5 if afternoon == "\u2713" else 0
            total_overtime += float(overtime or 0)

        # 计算额外统计信息
        attendance_days = total_morning + total_afternoon
        overtime_days = total_overtime / 6
        total_days = attendance_days + overtime_days

        emp_id = int(emp_id)  # 将 emp_id 从字符串转换为整数
        emp_info = employee_info.get(emp_id, {})
        total_salary = total_days * float(emp_info.get('unit_price', 0))
        total_salary = round(total_salary, 2)  # 保留两位小数

        # 填充出勤天数,加班计时,合计工数,合计工资列的数据并设置格式
        ws.merge_cells(start_row=start_row, start_column=days_in_month + 4, end_row=start_row + 1,
                       end_column=days_in_month + 4)
        cell5 = ws.cell(row=start_row, column=days_in_month + 4, value=attendance_days)
        cell5.number_format = "0.00"  # 设置保留两位小数
        cell5.alignment = Alignment(horizontal="center", vertical="center")

        # 出勤天数下的加班换算天数
        ws.merge_cells(start_row=start_row + 2, start_column=days_in_month + 4, end_row=start_row + 2,
                       end_column=days_in_month + 4)
        cell6 = ws.cell(row=start_row + 2, column=days_in_month + 4, value=overtime_days)
        cell6.number_format = "0.00"  # 设置保留两位小数
        cell6.alignment = Alignment(horizontal="center", vertical="center")

        # 加班计时
        ws.merge_cells(start_row=start_row, start_column=days_in_month + 5, end_row=start_row + 2,
                       end_column=days_in_month + 5)
        cell7 = ws.cell(row=start_row, column=days_in_month + 5, value=total_overtime)
        cell7.number_format = "0.00"  # 设置保留两位小数
        cell7.alignment = Alignment(horizontal="center", vertical="center")

        # 合计工数
        ws.merge_cells(start_row=start_row, start_column=days_in_month + 6, end_row=start_row + 2,
                       end_column=days_in_month + 6)
        cell8 = ws.cell(row=start_row, column=days_in_month + 6, value=total_days)
        cell8.number_format = "0.00"  # 设置保留两位小数
        cell8.alignment = Alignment(horizontal="center", vertical="center")

        # 合计工资
        ws.merge_cells(start_row=start_row, start_column=days_in_month + 8, end_row=start_row + 2,
                       end_column=days_in_month + 8)
        cell9 = ws.cell(row=start_row, column=days_in_month + 8, value=total_salary)
        cell9.number_format = "0.00"  # 设置保留两位小数
        cell9.alignment = Alignment(horizontal="center", vertical="center")

        # 签名
        ws.merge_cells(start_row=start_row, start_column=days_in_month + 9, end_row=start_row + 2,
                       end_column=days_in_month + 9)
        cell10 = ws.cell(row=start_row, column=days_in_month + 9, value='')
        cell10.alignment = Alignment(horizontal="center", vertical="center")

        # 填充员工的基本信息（职位-姓名、人工单价）
        if emp_info:
            logging.info(f"正在填充员工信息: {emp_info}")
            # 职位-姓名
            ws.merge_cells(start_row=start_row, start_column=3, end_row=start_row + 2, end_column=3)
            cell11 = ws.cell(row=start_row, column=3,
                             value=emp_info.get('job_type', '') + '-' + emp_info.get('name', ''))
            logging.info(f"填充职位-姓名: {emp_info.get('job_type', '')} {emp_info.get('name', '')}")
            cell11.number_format = "0.00"  # 设置保留两位小数
            cell11.alignment = Alignment(horizontal="center", vertical="center")

            # 人工单价
            ws.merge_cells(start_row=start_row, start_column=days_in_month + 7, end_row=start_row + 2,
                           end_column=days_in_month + 7)
            cell12 = ws.cell(row=start_row, column=days_in_month + 7, value=emp_info.get('unit_price', 0))
            logging.info(f"填充人工单价: {emp_info.get('unit_price', 0)}")
            cell12.number_format = "0.00"  # 设置保留两位小数
            cell12.alignment = Alignment(horizontal="center", vertical="center")

        start_row += 3  # 下一位员工的行

    wb.save(filename)
    logging.info(f"考勤表已更新并保存为: {filename}")


def get_days_in_month(year, month):
    """获取指定年份和月份的天数"""
    return calendar.monthrange(year, month)[1]


def load_and_validate_employees(file_path):
    """加载并验证员工数据"""
    employees = load_employees_from_txt(file_path)
    if not employees:
        print(f"员工数据加载失败，请检查文件：{file_path}")
        exit(1)
    return employees


def load_and_validate_db_config(file_path):
    """加载并验证数据库配置"""
    db_config = load_db_config(file_path)
    if not db_config:
        print(f"数据库配置加载失败，请检查文件：{file_path}")
        exit(1)
    return db_config


def parse_all_attendance(employees, days_in_month):
    """解析所有员工的考勤数据"""
    attendance_data = {}
    for emp_id, input_str in employees.items():
        emp_attendance = parse_attendance_input(input_str, [emp_id], days_in_month)
        attendance_data[emp_id] = emp_attendance[emp_id]
    return attendance_data


def main():
    # 动态获取年份和月份
    year, month = 2024, 12
    days_in_month = get_days_in_month(year, month)

    # 文件路径
    base_dir = os.path.dirname(__file__)
    employee_file = os.path.join(base_dir, "employees.txt")
    db_config_file = os.path.join(base_dir, "db_config.json")
    output_file = os.path.join(base_dir, "考勤表.xlsx")

    # 加载员工数据和数据库配置
    employees = load_and_validate_employees(employee_file)
    db_config = load_and_validate_db_config(db_config_file)

    # 获取员工信息
    emp_ids = list(employees.keys())

    employee_info = get_employee_info_from_mysql(db_config, emp_ids)
    if not employee_info:
        print("无法获取员工信息")
        exit(1)

    # 解析考勤数据
    attendance_data = parse_all_attendance(employees, days_in_month)

    # 填充考勤表
    fill_attendance(output_file, attendance_data, employee_info, emp_ids, days_in_month)
if __name__ == "__main__":
    main()
