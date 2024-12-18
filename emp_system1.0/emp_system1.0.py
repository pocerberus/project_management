"""
本系统能正常运行，录入员工信息，考勤信息，并搜索显示员工考勤，而且考勤显示方式正确，版本标号1.0
有以下需要改进

1、主页面有管理员登录、用户登录、用户注册、用户密码找回。用户注册有输入电话号码和密保问题，可以用来找回密码。
2、管理员登录成功后页面显示信息录入、信息查询、考勤录入、考勤查询，用户登录成功后只显示考勤录入、考勤查询、信息查询
    员工信息
    1、信息录入。要有编号，姓名，工种，国籍，工价，照片，备注。
    编号为已录入员工最大编号+1，第一位录入的员工编号为1，
        根据编号、工种、国籍自动生成工号，规则如下：
            工号一共5位数
            工号第一位数字由1-7，依次代表瓦工、木工、水电、油工、男小工、女小工、翻译，根据员工信息录入时的工种生成。工号后4位各工种前1000（含）
            数字当员工是China国籍时候使用，1001-9999数字由国籍为Cambodia使用，具体数字由录入员工编号与工种对应数字一起生成。例如国籍为China，
            木工，编号1，则工号为：10001，如果国籍为Cambodia，木工，编号1，则工号为：11001.
    2、信息查询。有搜索框的查询，输入名字或者工号，则显示指定员工信息，如输入工种信息，则显示指定工种所有员工信息。如不输入，则显示所有员工信息。员工
    信息要显示员工工号、姓名、工种、国籍、工价、照片、备注。
        3、考勤录入。选择录入月份，默认是上次录入月份，输入工号。能够快速录入，数字-数字表示几号到几号整天都上班，数字.1表示该天只上午上班，下午没上
        班，数字.2表示该天下午上班，上午没上班，数字+数字，表示当天上班且加班小时为第二个数字，例如1-5,6.1，7.2,10+3表示1-5号上班，6号上午上班，下午
        没上班，7号下午上班，上午没上班，10号上班且加班3小时
    4、考勤查询。有时间输入框、工种输入框、员工信息输入框。时间输入年月，默认是上月，如不输入则显示查询所有月份。输入则显示查询输入月份。工种输入框及
    员工信息输入框与时间输入框功能类似；
        考勤查询显示。第一列显示自然序号，第二列显示工种，第三列显示工号及员工姓名。后面列横向显示日期，第一行是日期，显示整个月的日期1号到月底最后一
    天，对应日期下面如果上班显示打钩，未上班显示打叉，分上午和下午2行。加班显示在第三行，显示加班时间。查询结果可以导出为Excel表格

"""

import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
import calendar


class AttendanceSystem:
    def __init__(self, root):
        self.root = root
        self.root.title("员工考勤系统")

        # 初始化员工信息和考勤数据的 DataFrame
        self.employees_df = pd.DataFrame(columns=["emp_id", "emp_name"])
        self.attendance_df = pd.DataFrame(columns=["emp_id", "emp_name", "date", "status"])

        # 当前选择的月份
        self.current_month = datetime.now().strftime("%Y-%m")

        # 创建界面
        self.create_main_ui()

    def create_main_ui(self):
        """创建主界面"""
        # 添加月份选择框
        self.month_frame = tk.LabelFrame(self.root, text="选择月份", padx=10, pady=10)
        self.month_frame.grid(row=0, column=0, padx=10, pady=10)

        tk.Label(self.month_frame, text="月份:").grid(row=0, column=0)
        self.month_combobox = ttk.Combobox(self.month_frame, values=self.generate_months(), state="readonly")
        self.month_combobox.grid(row=0, column=1)
        self.month_combobox.set(self.current_month)

        self.change_month_btn = tk.Button(self.month_frame, text="切换", command=self.change_month)
        self.change_month_btn.grid(row=0, column=2, padx=10)

        # 添加员工信息框
        self.employee_frame = tk.LabelFrame(self.root, text="员工管理", padx=10, pady=10)
        self.employee_frame.grid(row=1, column=0, padx=10, pady=10)

        tk.Label(self.employee_frame, text="姓名:").grid(row=0, column=0)
        self.emp_name_entry = tk.Entry(self.employee_frame)
        self.emp_name_entry.grid(row=0, column=1)

        self.add_employee_btn = tk.Button(self.employee_frame, text="添加员工", command=self.add_employee)
        self.add_employee_btn.grid(row=0, column=2, padx=10)

        # 快速录入考勤
        self.quick_fill_frame = tk.LabelFrame(self.root, text="快速录入考勤", padx=10, pady=10)
        self.quick_fill_frame.grid(row=2, column=0, padx=10, pady=10)

        tk.Label(self.quick_fill_frame, text="选择员工:").grid(row=0, column=0)
        self.employee_combobox = ttk.Combobox(self.quick_fill_frame, state="readonly")
        self.employee_combobox.grid(row=0, column=1)

        tk.Label(self.quick_fill_frame, text="日期范围 (如 1-10,15):").grid(row=1, column=0)
        self.date_range_entry = tk.Entry(self.quick_fill_frame)
        self.date_range_entry.grid(row=1, column=1)

        self.record_attendance_btn = tk.Button(self.quick_fill_frame, text="录入考勤", command=self.record_attendance)
        self.record_attendance_btn.grid(row=2, column=1, pady=10)

        # 查看考勤
        self.view_attendance_frame = tk.LabelFrame(self.root, text="查看考勤", padx=10, pady=10)
        self.view_attendance_frame.grid(row=3, column=0, padx=10, pady=10)

        tk.Label(self.view_attendance_frame, text="输入员工编号或姓名:").grid(row=0, column=0)
        self.search_entry = tk.Entry(self.view_attendance_frame)
        self.search_entry.grid(row=0, column=1)

        self.view_attendance_btn = tk.Button(self.view_attendance_frame, text="查看考勤",
                                             command=self.view_employee_attendance)
        self.view_attendance_btn.grid(row=0, column=2)

        # 更新下拉框
        self.update_employee_combobox()

    def generate_months(self):
        """生成最近12个月的列表"""
        today = datetime.now()
        return [(today - relativedelta(months=i)).strftime("%Y-%m") for i in range(12)]

    def change_month(self):
        """切换月份"""
        self.current_month = self.month_combobox.get()

    def add_employee(self):
        """添加员工"""
        emp_name = self.emp_name_entry.get()
        if not emp_name.strip():
            messagebox.showerror("错误", "员工姓名不能为空！")
            return

        emp_id = len(self.employees_df) + 1
        self.employees_df = pd.concat(
            [self.employees_df, pd.DataFrame({"emp_id": [emp_id], "emp_name": [emp_name]})],
            ignore_index=True,
        )
        self.initialize_attendance(emp_id, emp_name)
        self.update_employee_combobox()
        messagebox.showinfo("成功", f"员工 {emp_name} 已添加！")
        self.emp_name_entry.delete(0, tk.END)

    def initialize_attendance(self, emp_id, emp_name):
        """初始化员工考勤记录"""
        year, month = map(int, self.current_month.split("-"))
        days_in_month = calendar.monthrange(year, month)[1]

        for day in range(1, days_in_month + 1):
            date = f"{self.current_month}-{day:02d}"
            self.attendance_df = pd.concat(
                [
                    self.attendance_df,
                    pd.DataFrame({"emp_id": [emp_id], "emp_name": [emp_name], "date": [date], "status": ["absent"]}),
                ],
                ignore_index=True,
            )

    def update_employee_combobox(self):
        """更新员工下拉框"""
        if not self.employees_df.empty:
            self.employee_combobox["values"] = self.employees_df["emp_name"].tolist()
            self.employee_combobox.current(0)
        else:
            self.employee_combobox.set("")

    def record_attendance(self):
        """录入考勤"""
        emp_name = self.employee_combobox.get()
        if not emp_name:
            messagebox.showerror("错误", "请选择员工！")
            return

        date_range = self.date_range_entry.get()
        emp_id = self.employees_df[self.employees_df["emp_name"] == emp_name]["emp_id"].iloc[0]

        try:
            days = set()
            for part in date_range.split(","):
                if "-" in part:
                    start, end = map(int, part.split("-"))
                    days.update(range(start, end + 1))
                else:
                    days.add(int(part))

            for day in days:
                date = f"{self.current_month}-{day:02d}"
                self.attendance_df.loc[
                    (self.attendance_df["emp_id"] == emp_id) & (self.attendance_df["date"] == date), "status"
                ] = "present"

            messagebox.showinfo("成功", "考勤记录已更新！")

        except ValueError:
            messagebox.showerror("错误", "日期范围格式无效！")

    def view_employee_attendance(self):
        """查看员工考勤"""
        search_query = self.search_entry.get().strip()
        filtered_df = self.employees_df

        if search_query.isdigit():
            filtered_df = filtered_df[filtered_df["emp_id"] == int(search_query)]
        elif search_query:
            filtered_df = filtered_df[filtered_df["emp_name"].str.contains(search_query)]

        if filtered_df.empty:
            messagebox.showinfo("提示", "未找到匹配的员工记录！")
            return

        view_window = tk.Toplevel(self.root)
        view_window.title("查看考勤")

        year, month = map(int, self.current_month.split("-"))
        days_in_month = calendar.monthrange(year, month)[1]
        dates = [f"{day:02d}" for day in range(1, days_in_month + 1)]

        # 日期横向显示
        for col, date in enumerate(dates, start=2):
            tk.Label(view_window, text=date, width=5).grid(row=0, column=col)

        # 显示员工编号、姓名和考勤状态
        for row, emp in filtered_df.iterrows():
            emp_id, emp_name = emp["emp_id"], emp["emp_name"]
            tk.Label(view_window, text=emp_id, width=5).grid(row=row + 1, column=0)
            tk.Label(view_window, text=emp_name, width=10).grid(row=row + 1, column=1)

            for col, date in enumerate(dates, start=2):
                status = self.attendance_df.loc[
                    (self.attendance_df["emp_id"] == emp_id) & (
                            self.attendance_df["date"] == f"{self.current_month}-{date}")
                    ]["status"].values
                symbol = "✔" if status and status[0] == "present" else "✘"
                tk.Label(view_window, text=symbol, width=5).grid(row=row + 1, column=col)


if __name__ == "__main__":
    root = tk.Tk()
    AttendanceSystem(root)
    root.mainloop()
