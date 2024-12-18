import os
import mysql.connector
import re
import configparser


# 读取配置文件的函数
def read_config(file_path):
    config = configparser.ConfigParser()
    with open(file_path, 'r', encoding='utf-8') as f:
        config.read_file(f)
    return config


# 创建数据库连接的函数
def create_db_connection(config):
    try:
        db = mysql.connector.connect(
            host=config['Database']['Host'],
            user=config['Database']['User'],
            password=config['Database']['Password'],
            database=config['Database']['Database']
        )
        return db
    except mysql.connector.Error as err:
        print(f"Error: {err}")
        exit(1)  # 如果数据库连接失败，则退出程序


# 提取ID的统一函数
def extract_id_from_filename(filename):
    # 从文件名中提取最后两位数字部分作为ID。
    # 例如：'worker_1234.jpg' -> 34

    matches = re.findall(r'\d+', filename)
    if matches:
        last_number = matches[-1]  # 获取最后一个数字串
        return int(last_number[-2:])  # 取最后两位数字并返回
    return None  # 如果没有找到数字，返回 None


# 更新数据库中图片数据的函数
def update_image_in_db(cursor, table_name, image_data, image_id):
    try:
        cursor.execute(
            f"UPDATE {table_name} SET image = %s WHERE number = %s",  # 假设所有表都用 id 作为主键
            (image_data, image_id)
        )
        print(f"Updated image for {table_name} with ID {image_id}")
    except mysql.connector.Error as err:
        print(f"Error executing query for ID {image_id}: {err}")
        return False  # 如果出错，返回 False
    return True  # 成功执行 SQL 时返回 True


# 处理每个文件夹中的图片文件
def process_images_in_folder(folder_path, table_name, cursor):
    for image_name in os.listdir(folder_path):
        if image_name.endswith(('jpg', 'jpeg', 'png', 'gif', 'bmp')):  # 仅处理图片文件
            image_path = os.path.join(folder_path, image_name)
            try:
                with open(image_path, 'rb') as image_file:
                    image_data = image_file.read()

                # 使用统一规则提取ID
                image_id = extract_id_from_filename(image_name)
                if image_id is None:
                    print(f"Skipping file {image_name} due to invalid filename format.")
                    continue  # 跳过格式不正确的文件

                # 执行数据库更新操作
                if not update_image_in_db(cursor, table_name, image_data, image_id):
                    cursor.connection.rollback()  # 如果更新失败，回滚事务
                    continue  # 跳过当前图片，继续处理其他图片

            except Exception as e:
                print(f"Error processing file {image_name}: {e}")
                continue  # 跳过当前图片，继续处理其他图片


# 主要执行流程
def main():
    # 读取配置文件
    config = read_config('config.ini')

    # 创建数据库连接
    db = create_db_connection(config)
    cursor = db.cursor()

    # 获取配置项
    base_dir = config['Paths']['BaseDir']

    # 事先定义一个合法的表名列表
    valid_tables = ['carpenter', 'electrician', 'labourer_man', 'labourer_woman', 'painter', 'plasterer', 'translator']

    # 数据库处理逻辑
    try:
        # 遍历上一级目录下的子文件夹（假设每个子文件夹名对应一个表名）
        for folder_name in os.listdir(base_dir):
            folder_path = os.path.join(base_dir, folder_name)

            if os.path.isdir(folder_path):
                print(f"Processing folder: {folder_name}")

                # 验证文件夹名是否是合法的表名
                table_name = folder_name.lower()  # 转为小写作为表名
                if table_name not in valid_tables:
                    print(f"Skipping invalid folder/table name: {table_name}")
                    continue  # 如果文件夹名不是合法的表名，跳过

                # 处理当前文件夹中的所有图片文件
                process_images_in_folder(folder_path, table_name, cursor)

                # 提交事务（处理完每个文件夹后提交）
                db.commit()

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # 关闭数据库连接
        cursor.close()
        db.close()


# 执行主函数
if __name__ == "__main__":
    main()

# '''
# 代码解析

# import os
# import mysql.connector
# import re
# import configparser
# 导入必要的模块：

# os：用于处理操作系统相关的文件和目录操作。
# mysql.connector：用于连接和操作 MySQL 数据库。
# re：用于正则表达式匹配，特别是用来从文件名中提取数字。
# configparser：用于读取和解析配置文件（.ini 文件）。

# 配置读取部分
# def read_config(file_path):
#    config = configparser.ConfigParser()
#    with open(file_path, 'r', encoding='utf-8') as f:
#        config.read_file(f)
#    return config
# read_config(file_path):
# 功能：该函数用于读取指定路径的配置文件并返回一个配置对象。
# config = configparser.ConfigParser()：创建一个 ConfigParser 实例，用于读取 .ini 格式的配置文件。
# with open(file_path, 'r', encoding='utf-8') as f:：使用 with 语句以只读模式打开指定的配置文件，确保在文件操作完成后自动关闭文件。
# config.read_file(f)：将文件内容读取到 config 对象中，config 对象现在包含了 .ini 文件的内容。
# return config：返回 config 对象，以便在程序的其他地方使用配置内容。

# 数据库连接部分
# def create_db_connection(config):
#    try:
#        db = mysql.connector.connect(
#            host=config['Database']['Host'],
#            user=config['Database']['User'],
#            password=config['Database']['Password'],
#            database=config['Database']['Database']
#        )
#        return db
#    except mysql.connector.Error as err:
#        print(f"Error: {err}")
#        exit(1)  # 如果数据库连接失败，则退出程序
# create_db_connection(config)：
# 功能：该函数使用从配置文件中读取的数据库连接参数来连接 MySQL 数据库，并返回数据库连接对象。
# db = mysql.connector.connect(...)：通过 mysql.connector.connect 创建一个连接对象。连接参数（如 host, user, password, database）来自 config 参数传入的配置文件中。
# except mysql.connector.Error as err:：如果连接过程中出现错误，会捕获到 mysql.connector.Error 异常。
# print(f"Error: {err}")：打印出错误信息。
# exit(1)：如果数据库连接失败，程序会退出并返回状态码 1（表示错误）。

# 提取ID部分
# def extract_id_from_filename(filename):

#    从文件名中提取最后两位数字部分作为ID。
#   例如：'worker_1234.jpg' -> 34
#
#    matches = re.findall(r'\d+', filename)
#    if matches:
#        last_number = matches[-1]  # 获取最后一个数字串
#        return int(last_number[-2:])  # 取最后两位数字并返回
#    return None  # 如果没有找到数字，返回 None
# extract_id_from_filename(filename)：
# 功能：从文件名中提取最后两位数字，作为 ID 返回。
# matches = re.findall(r'\d+', filename)：使用正则表达式 \d+ 查找文件名中的所有数字，返回一个数字串列表。
# 例如，'worker_1234.jpg' 会返回 ['1234']。
# if matches:：如果找到了数字（即 matches 不为空）。
# last_number = matches[-1]：取 matches 列表中的最后一个数字串。
# return int(last_number[-2:])：从这个数字串中取最后两位（last_number[-2:]），然后将其转为整数并返回。
# 例如，如果 last_number 是 '1234'，则返回 34。
# return None：如果文件名中没有数字，返回 None。

# 更新数据库部分
# def update_image_in_db(cursor, table_name, image_data, image_id):
#    try:
#        cursor.execute(
#            f"UPDATE {table_name} SET image = %s WHERE number = %s",  # 假设所有表都用 id 作为主键
#            (image_data, image_id)
#        )
#        print(f"Updated image for {table_name} with ID {image_id}")
#    except mysql.connector.Error as err:
#        print(f"Error executing query for ID {image_id}: {err}")
#        return False  # 如果出错，返回 False
#    return True  # 成功执行 SQL 时返回 True
# update_image_in_db(cursor, table_name, image_data, image_id)：
# 功能：将图片数据更新到数据库中。
# cursor.execute(...)：执行 SQL 语句，更新指定表（table_name）中指定 id 的记录的图片字段（image）。
# 使用 %s 占位符来传递参数，防止 SQL 注入。
# image_data 是图片的二进制数据，image_id 是从文件名中提取的 ID。
# except mysql.connector.Error as err:：如果在执行 SQL 语句时遇到 MySQL 错误，则捕获并打印错误信息。
# return False：如果执行失败，返回 False。
# return True：如果执行成功，返回 True。

# 处理每个文件夹中的图片文件
# def process_images_in_folder(folder_path, table_name, cursor):
#    for image_name in os.listdir(folder_path):
#        if image_name.endswith(('jpg', 'jpeg', 'png', 'gif', 'bmp')):  # 仅处理图片文件
#            image_path = os.path.join(folder_path, image_name)
#            try:
#                with open(image_path, 'rb') as image_file:
#                    image_data = image_file.read()
#
#                # 使用统一规则提取ID
#                image_id = extract_id_from_filename(image_name)
#                if image_id is None:
#                    print(f"Skipping file {image_name} due to invalid filename format.")
#                    continue  # 跳过格式不正确的文件
#
#                # 执行数据库更新操作
#                if not update_image_in_db(cursor, table_name, image_data, image_id):
#                    cursor.connection.rollback()  # 如果更新失败，回滚事务
#                    continue  # 跳过当前图片，继续处理其他图片
#
#            except Exception as e:
#                print(f"Error processing file {image_name}: {e}")
#               continue  # 跳过当前图片，继续处理其他图片
# process_images_in_folder(folder_path, table_name, cursor)：
# 功能：遍历文件夹中的所有图片文件，并将每个图片文件的数据更新到数据库。
# for image_name in os.listdir(folder_path):：遍历文件夹中的所有文件。
# if image_name.endswith(('jpg', 'jpeg', 'png', 'gif', 'bmp')):：检查文件是否为图片文件（根据扩展名）。
# image_path = os.path.join(folder_path, image_name)：获取图片的完整路径。
# with open(image_path, 'rb') as image_file:：以二进制模式打开图片文件。
# image_data = image_file.read()：读取文件的二进制数据。
# image_id = extract_id_from_filename(image_name)：调用 extract_id_from_filename 函数提取图片的 ID。
# if image_id is None:：如果 ID 提取失败（例如文件名不符合格式），跳过该文件。
# if not update_image_in_db(...)：调用 update_image_in_db 函数将图片数据更新到数据库。
# 如果更新失败，回滚数据库事务，继续处理下一个图片。
# except Exception as e:：捕获其他异常（如文件读取错误），并跳过当前文件。

# 主执行流程
# def main():
#    # 读取配置文件
#    config = read_config('config.ini')

#   # 创建数据库连接
#    db = create_db_connection(config)
#    cursor = db.cursor()

#    # 获取配置项
#    base_dir = config['Paths']['BaseDir']

#    # 事先定义一个合法的表名列表
#    valid_tables = ['carpenter', 'electrician', 'labourer_man', 'labourer_woman', 'painter', 'plasterer', 'translator']

#    # 数据库处理逻辑
#    try:
#        # 遍历上一级目录下的子文件夹（假设每个子文件夹名对应一个表名）
#        for folder_name in os.listdir(base_dir):
#            folder_path = os.path.join(base_dir, folder_name)

#            if os.path.isdir(folder_path):
#                print(f"Processing folder: {folder_name}")
#
#                # 验证文件夹名是否是合法的表名
#               table_name = folder_name.lower()  # 转为小写作为表名
#                if table_name not in valid_tables:
#                    print(f"Skipping invalid folder/table name: {table_name}")
#                    continue  # 如果文件夹名不是合法的表名，跳过
#
#                # 处理当前文件夹中的所有图片文件
#                process_images_in_folder(folder_path, table_name, cursor)
#
#                # 提交事务（处理完每个文件夹后提交）
#                db.commit()
#
#    except Exception as e:
#        print(f"An error occurred: {e}")
#    finally:
#        # 关闭数据库连接
#        cursor.close()
#        db.close()
# main()：
# 功能：程序的主执行逻辑，管理配置文件读取、数据库连接、文件夹遍历和图片处理等工作。
# config = read_config('config.ini')：读取配置文件。
# db = create_db_connection(config)：创建数据库连接。
# base_dir = config['Paths']['BaseDir']：从配置文件中获取存放图片的根目录。
# valid_tables = [...]：预定义合法的表名列表。
# for folder_name in os.listdir(base_dir):：遍历根目录下的每个文件夹（每个文件夹假设对应一个数据库表）。
# if os.path.isdir(folder_path):：如果是文件夹，则进行处理。
# table_name = folder_name.lower()：将文件夹名转为小写作为表名。
# if table_name not in valid_tables:：检查文件夹名是否是合法的表名，若不是，跳过。
# process_images_in_folder(folder_path, table_name, cursor)：调用 process_images_in_folder 处理文件夹中的图片。
# db.commit()：提交事务。
# except Exception as e:：捕获异常并打印错误信息。
# finally:：确保在结束时关闭数据库连接。

# 执行主函数
# if __name__ == "__main__":
#    main()
# if __name__ == "__main__":：确保程序仅在作为主程序执行时调用 main() 函数，而在作为模块导入时不执行 main()。

# 总结
# 模块化：通过将配置读取、数据库连接、图片处理等功能封装成函数，使得代码结构清晰且便于维护和扩展。
# 错误处理：在关键操作中使用了 try-except 语句，确保在发生错误时能够处理并继续执行。
# 灵活性：通过配置文件和函数化设计，使得该程序具备较高的灵活性，便于以后修改、扩展或调整。
