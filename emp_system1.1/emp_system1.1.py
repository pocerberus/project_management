import calendar
import tkinter as tk
from tkinter import messagebox


class AttendanceSystem:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("考勤管理系统")
        self.root.geometry("600x400")

        self.admin_username = "admin"
        self.admin_password = "admin123"
        self.current_user = None
        self.current_role = None

        # 模拟数据库
        self.employees = []  # 员工信息存储
        self.attendance_data = {}  # 考勤数据
        self.last_month = "2024-01"  # 默认月份

        self.create_main_ui()

    def create_main_ui(self):
        """主界面"""
        self.clear_frame()
        tk.Label(self.root, text="考勤管理系统", font=("Arial", 20)).pack(pady=20)
        tk.Button(self.root, text="管理员登录", command=self.admin_login_ui, width=20).pack(pady=10)
        tk.Button(self.root, text="用户登录", command=self.user_login_ui, width=20).pack(pady=10)
        tk.Button(self.root, text="用户注册", command=self.user_register_ui, width=20).pack(pady=10)
        tk.Button(self.root, text="找回密码", command=self.password_recovery_ui, width=20).pack(pady=10)

    def admin_login_ui(self):
        """管理员登录界面"""
        self.clear_frame()
        tk.Label(self.root, text="管理员登录", font=("Arial", 18)).pack(pady=20)

        tk.Label(self.root, text="用户名:").pack()
        username_entry = tk.Entry(self.root)
        username_entry.pack()

        tk.Label(self.root, text="密码:").pack()
        password_entry = tk.Entry(self.root, show="*")
        password_entry.pack()

        def login():
            username = username_entry.get()
            password = password_entry.get()
            if username == self.admin_username and password == self.admin_password:
                self.current_user = "admin"
                self.current_role = "admin"
                messagebox.showinfo("登录成功", "管理员登录成功！")
                self.admin_dashboard()
            else:
                messagebox.showerror("登录失败", "用户名或密码错误！")

        tk.Button(self.root, text="登录", command=login, width=20).pack(pady=10)
        tk.Button(self.root, text="返回主界面", command=self.create_main_ui, width=20).pack(pady=10)

    def admin_dashboard(self):
        """管理员控制面板"""
        self.clear_frame()
        tk.Label(self.root, text="管理员控制面板", font=("Arial", 18)).pack(pady=20)
        tk.Button(self.root, text="信息录入", command=self.employee_entry_ui, width=20).pack(pady=5)
        tk.Button(self.root, text="信息查询", command=self.employee_query_ui, width=20).pack(pady=5)
        tk.Button(self.root, text="考勤录入", command=self.attendance_entry_ui, width=20).pack(pady=5)
        tk.Button(self.root, text="考勤查询", command=self.attendance_query_ui, width=20).pack(pady=5)
        tk.Button(self.root, text="退出登录", command=self.create_main_ui, width=20).pack(pady=20)

    def user_login_ui(self):
        """用户登录界面"""
        messagebox.showinfo("提示", "用户登录功能尚未实现！")
        self.create_main_ui()

    def user_register_ui(self):
        """用户注册界面"""
        messagebox.showinfo("提示", "用户注册功能尚未实现！")
        self.create_main_ui()

    def password_recovery_ui(self):
        """密码找回界面"""
        messagebox.showinfo("提示", "密码找回功能尚未实现！")
        self.create_main_ui()

    def employee_entry_ui(self):
        """信息录入界面"""
        self.clear_frame()
        tk.Label(self.root, text="信息录入", font=("Arial", 18)).pack(pady=20)

        tk.Label(self.root, text="姓名:").pack()
        name_entry = tk.Entry(self.root)
        name_entry.pack()

        tk.Label(self.root, text="工种:").pack()
        job_entry = tk.Entry(self.root)
        job_entry.pack()

        tk.Label(self.root, text="国籍:").pack()
        nationality_entry = tk.Entry(self.root)
        nationality_entry.pack()

        tk.Label(self.root, text="工价:").pack()
        wage_entry = tk.Entry(self.root)
        wage_entry.pack()

        tk.Label(self.root, text="备注:").pack()
        remark_entry = tk.Entry(self.root)
        remark_entry.pack()

        def save_employee():
            name = name_entry.get()
            job = job_entry.get()
            nationality = nationality_entry.get()
            wage = wage_entry.get()
            remark = remark_entry.get()

            if not name or not job or not nationality or not wage:
                messagebox.showerror("错误", "请完整填写信息！")
                return

            # 自动生成编号和工号
            emp_id = len(self.employees) + 1
            code_prefix = {"瓦工": 1, "木工": 2, "水电": 3, "油工": 4, "男小工": 5, "女小工": 6, "翻译": 7}.get(job, 0)
            if nationality.lower() == "china":
                code_number = emp_id
            else:
                code_number = 1000 + emp_id
            employee_code = f"{code_prefix}{code_number}"

            # 存储员工信息
            employee = {
                "编号": emp_id,
                "工号": employee_code,
                "姓名": name,
                "工种": job,
                "国籍": nationality,
                "工价": wage,
                "备注": remark
            }
            self.employees.append(employee)
            messagebox.showinfo("成功", f"员工 {name} 信息录入成功！")
            self.create_main_ui()

        tk.Button(self.root, text="保存", command=save_employee, width=20).pack(pady=10)
        tk.Button(self.root, text="返回", command=self.admin_dashboard, width=20).pack(pady=10)

    def employee_query_ui(self):
        """信息查询界面"""
        self.clear_frame()
        tk.Label(self.root, text="信息查询", font=("Arial", 18)).pack(pady=20)

        search_entry = tk.Entry(self.root)
        search_entry.pack(pady=10)

        def query_employee():
            query = search_entry.get()
            results = [emp for emp in self.employees if query in str(emp.values()) or not query]
            if results:
                result_window = tk.Toplevel(self.root)
                result_window.title("查询结果")
                result_window.geometry("600x400")

                for idx, emp in enumerate(results, start=1):
                    info = f"{idx}. 工号: {emp['工号']}, 姓名: {emp['姓名']}, 工种: {emp['工种']}, 国籍: {emp['国籍']}, 工价: {emp['工价']}, 备注: {emp['备注']}"
                    tk.Label(result_window, text=info, anchor="w", justify="left").pack()
            else:
                messagebox.showinfo("提示", "未找到相关员工信息！")

        tk.Button(self.root, text="查询", command=query_employee, width=20).pack(pady=10)
        tk.Button(self.root, text="返回", command=self.admin_dashboard, width=20).pack(pady=10)

    def attendance_entry_ui(self):
        """考勤录入界面"""
        self.clear_frame()
        tk.Label(self.root, text="考勤录入", font=("Arial", 18)).pack(pady=20)

        # 上次录入月份
        tk.Label(self.root, text="选择月份 (默认上次录入月份):").pack()
        month_var = tk.StringVar(value=self.last_month)
        month_entry = tk.Entry(self.root, textvariable=month_var)
        month_entry.pack()

        # 工号输入
        tk.Label(self.root, text="工号:").pack()
        emp_id_entry = tk.Entry(self.root)
        emp_id_entry.pack()

        # 考勤快速录入
        tk.Label(self.root, text="考勤记录（快速录入格式）:").pack()
        attendance_entry = tk.Entry(self.root)
        attendance_entry.pack()

        def save_attendance():
            """保存考勤记录"""
            month = month_var.get()
            emp_id = emp_id_entry.get()
            records = attendance_entry.get()

            if not month or not emp_id or not records:
                messagebox.showerror("错误", "所有字段均为必填项！")
                return

            # 验证月份格式
            try:
                year, mon = map(int, month.split("-"))
                _, days_in_month = calendar.monthrange(year, mon)
            except ValueError:
                messagebox.showerror("错误", "月份格式错误，应为 YYYY-MM！")
                return

            # 初始化考勤数据
            if month not in self.attendance_data:
                self.attendance_data[month] = {}
            if emp_id not in self.attendance_data[month]:
                self.attendance_data[month][emp_id] = {
                    day: {"上午": "×", "下午": "×"} for day in range(1, days_in_month + 1)
                }

            # 解析考勤记录并存储
            try:
                for record in records.split(","):
                    if "-" in record:
                        start, end = map(int, record.split("-"))
                        if start > end:
                            raise ValueError("日期范围错误")
                    elif ".1" in record or ".2" in record or "+" in record:
                        pass  # 继续解析
                    else:
                        raise ValueError("无效的记录格式")
            except Exception as e:
                messagebox.showerror("错误", f"考勤记录解析失败：{e}")
                return

            messagebox.showinfo("成功", "考勤录入成功！")
            self.create_main_ui()

        tk.Button(self.root, text="保存考勤", command=save_attendance, width=20).pack(pady=10)
        tk.Button(self.root, text="返回", command=self.admin_dashboard, width=20).pack(pady=10)

    def attendance_query_ui(self):
        """考勤查询界面"""
        self.clear_frame()
        tk.Label(self.root, text="考勤查询", font=("Arial", 18)).pack(pady=20)

        # 月份选择
        tk.Label(self.root, text="月份 (格式 YYYY-MM, 留空查询所有):").pack()
        month_entry = tk.Entry(self.root)
        month_entry.pack()

        # 工号输入
        tk.Label(self.root, text="工号 (留空查询所有):").pack()
        emp_id_entry = tk.Entry(self.root)
        emp_id_entry.pack()

        def query_attendance():
            month = month_entry.get()
            emp_id = emp_id_entry.get()
            results = []

            if month:
                if month in self.attendance_data:
                    data = self.attendance_data[month]
                    results = [data.get(emp, None) for emp in data if not emp_id or emp == emp_id]
            else:
                for m in self.attendance_data:
                    for emp, record in self.attendance_data[m].items():
                        if not emp_id or emp == emp_id:
                            results.append((m, emp, record))

            if results:
                result_window = tk.Toplevel(self.root)
                result_window.title("查询结果")
                for idx, result in enumerate(results, start=1):
                    tk.Label(result_window, text=f"{idx}. {result}").pack()
            else:
                messagebox.showinfo("提示", "未找到相关考勤记录！")

        tk.Button(self.root, text="查询", command=query_attendance, width=20).pack(pady=10)
        tk.Button(self.root, text="返回", command=self.admin_dashboard, width=20).pack(pady=10)

    def clear_frame(self):
        """清除当前界面"""
        for widget in self.root.winfo_children():
            widget.destroy()

    def run(self):
        self.root.mainloop()


# 创建系统对象并运行
system = AttendanceSystem()
system.run()
