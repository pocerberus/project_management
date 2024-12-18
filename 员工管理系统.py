# 员工管理系统
import mysql
import pillow

# 业务需求

# 1、账号登录
#   管理员账号、客户账号、游客账号
#   账号功能要有注册、密码找回、查找、删除等功能

# 2、人工信息
#   人工信息表（必须含有工号、姓名、工种、工价、籍别、备注）
#   人工信息表要有增删改查功能，信息要包含照片  
#   工种包括瓦工、木工、水电、油工、男小工、女小工、翻译
#   人工工号暂定各工种前1000编号为中工使用，后1001-9999为其他人工使用，工种编号1-7，工号为工种编号+人工编号，即5位数编码。设置超出编号范围后显示错误，需重新设计工号。

# 3、每天进度信息
#   以各工种工序容易区分工作命名单项项目名称，每天记录该项工作进度及人工情况
#   单项项目施工计量以房间为单位，房间房号、楼层层号、项目名称、施工日期为联合主键

# 4、人工考勤统计表
#   每日登记每人出勤

# 5、输出成果
#   人工个人当前工资查询
#   每月人工工资汇总
#   项目情况汇报
#       当前总进度及人工总成本
#       当前单项工种进度及人工总成本
#       当前单项项目进度及人工总成本


# 员工管理系统

# 创建数据库

CREATE
DATABASE
company;

USE
company;

CREATE
TABLE
employees(
    id
INT
AUTO_INCREMENT
PRIMARY
KEY,
name
VARCHAR(100),
position
VARCHAR(100),
salary
DECIMAL(10, 2),
photo
BLOB
);



# 安装库
pip
install
mysql - connector
pillow

# 数据库连接操作

import mysql.connector
from mysql.connector import Error
from PIL import Image
import io


# 连接到MySQL数据库
def create_connection():
    try:
        connection = mysql.connector.connect(
            host='localhost',
            database='company',
            user='root',  # 根据你的数据库用户名调整
            password='password'  # 根据你的数据库密码调整
        )
        if connection.is_connected():
            print("成功连接到数据库")
        return connection
    except Error as e:
        print(f"数据库连接失败: {e}")
        return None


# 插入员工信息
def insert_employee(name, position, salary, photo_path):
    try:
        conn = create_connection()
        cursor = conn.cursor()

        with open(photo_path, 'rb') as file:
            photo = file.read()

        query = """INSERT INTO employees (name, position, salary, photo) 
                   VALUES (%s, %s, %s, %s)"""
        cursor.execute(query, (name, position, salary, photo))
        conn.commit()
        print("员工信息插入成功")
    except Error as e:
        print(f"插入失败: {e}")
    finally:
        if conn.is_connected():
            cursor.close()
            conn.close()


# 获取所有员工信息
def get_all_employees():
    try:
        conn = create_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM employees")
        employees = cursor.fetchall()
        return employees
    except Error as e:
        print(f"查询失败: {e}")
        return []
    finally:
        if conn.is_connected():
            cursor.close()
            conn.close()


# 创建图形界面

import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter import ttk
from PIL import ImageTk, Image


# 选择文件对话框，选择照片
def upload_photo():
    filename = filedialog.askopenfilename(filetypes=[("Image files", "*.jpg;*.jpeg;*.png")])
    if filename:
        photo_path.set(filename)
        img = Image.open(filename)
        img.thumbnail((100, 100))
        img = ImageTk.PhotoImage(img)
        photo_label.config(image=img)
        photo_label.image = img


# 添加员工
def add_employee():
    name = name_entry.get()
    position = position_entry.get()
    salary = salary_entry.get()
    photo = photo_path.get()

    if not name or not position or not salary or not photo:
        messagebox.showwarning("输入错误", "所有字段必须填写")
        return

    insert_employee(name, position, float(salary), photo)
    messagebox.showinfo("成功", "员工信息已添加")
    clear_form()
    update_employee_list()


# 清空输入框
def clear_form():
    name_entry.delete(0, tk.END)
    position_entry.delete(0, tk.END)
    salary_entry.delete(0, tk.END)
    photo_path.set("")
    photo_label.config(image='')


# 更新员工列表
def update_employee_list():
    for row in tree.get_children():
        tree.delete(row)

    employees = get_all_employees()
    for emp in employees:
        tree.insert('', 'end', values=(emp[0], emp[1], emp[2], emp[3]))


# 创建主界面
root = tk.Tk()
root.title("员工管理系统")

# 添加员工界面
frame = tk.Frame(root)
frame.pack(pady=10)

tk.Label(frame, text="姓名").grid(row=0, column=0)
tk.Label(frame, text="职位").grid(row=1, column=0)
tk.Label(frame, text="薪资").grid(row=2, column=0)
tk.Label(frame, text="照片").grid(row=3, column=0)

name_entry = tk.Entry(frame)
name_entry.grid(row=0, column=1)

position_entry = tk.Entry(frame)
position_entry.grid(row=1, column=1)

salary_entry = tk.Entry(frame)
salary_entry.grid(row=2, column=1)

photo_path = tk.StringVar()

tk.Button(frame, text="上传照片", command=upload_photo).grid(row=3, column=1)
photo_label = tk.Label(frame)
photo_label.grid(row=3, column=2)

tk.Button(frame, text="添加员工", command=add_employee).grid(row=4, column=1, pady=10)

# 员工列表界面
tree_frame = tk.Frame(root)
tree_frame.pack(pady=10)

columns = ('ID', '姓名', '职位', '薪资')
tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
tree.heading('ID', text='ID')
tree.heading('姓名', text='姓名')
tree.heading('职位', text='职位')
tree.heading('薪资', text='薪资')

tree.pack()

# 更新员工列表
update_employee_list()

root.mainloop()
