import datetime
import re
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart

import pandas as pd
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk

import xlwings

# import timedelta

# 创建窗口
window = tk.Tk()

# 设置窗口标题
window.title("Tracking List Tool")

# 创建包含两个文本框的框架
frame = tk.Frame(window)
frame.pack(padx=5, pady=5, side="top")


######################################################################
# 创建获取Excel  待开发获取第二个工作表


def get_excel():
    global sheet_names, workbook
    # 获取当前打开的Excel文件的工作簿对象
    workbook = xlwings.books.active
    # 获取工作簿中的所有工作表名称
    sheet_names = [sheet.name for sheet in workbook.sheets]
    filename = workbook.name
    if filename != "EE_Task_Tracking_List.xlsx":
        messagebox.showerror(parent=window, title="文件识别错误",
                             message="文件识别错误，请选择正确文件后，再次点击此按钮")
    else:
        messagebox.showinfo(parent=window, title="文件识别成功",
                            message="文件识别成功，请选择Sheet")
        # 获取第一个工作表
        combo_box3["values"] = sheet_names
        combo_box3.current(0)


# 获取Excel按钮事件
button = tk.Button(frame, text="获取Excel文件",
                   font=("Arial", 10), command=get_excel)
button.pack(padx=5, pady=5, side="left")
######################################################################


######################################################################
# 创建下拉列表
combo_box3 = ttk.Combobox(frame, state="readonly")
# combo_box3.current(0)  # 设置默认选项为第一个工作表
combo_box3.pack(padx=5, pady=5, side="left")

# 创建获取下拉列表选项的按钮


def get_selected_sheet():
    global worksheet
    selected_sheet = combo_box3.get()
    selected_index = sheet_names.index(selected_sheet)
    worksheet = workbook.sheets[selected_index]
    # root = tk.Tk()
    # root.withdraw()
    messagebox.showinfo(parent=window, title="",
                        message=f"你选择了工作表：{selected_sheet}")
    # print(f"你选择了工作表：{selected_sheet}")


button1 = tk.Button(frame, text="获取选择的工作表",
                    font=("Arial", 10), command=get_selected_sheet)
button1.pack(padx=5, pady=5, side="left")
######################################################################


######################################################################
# 创建负责人下拉列表
combo_box = ttk.Combobox(frame,  values=[
                         'Kitty', 'Burce', 'Isaac', 'Liu Jie', 'Mark', 'Carrey', 'Ligen', 'Siyuan'], state="readonly")
combo_box.current(5)
combo_box.pack(side="left")

# 查询按钮事件


def get_principal():
    # 获取负责人下拉列表的文本
    global selected_principal_index
    selected_principal_index = combo_box.current()
    combo_box1.current(selected_principal_index)
    selected = combo_box.get()
    if selected:
        try:
            filter_data(selected)
        except Exception as e:
            messagebox.showerror(
                parent=window, title="出错了！", message="出错了！请联系管理员！")
    else:
        messagebox.showerror(parent=window, title="负责人空白", message="请输入负责人名称")


button1 = tk.Button(frame, text="查询",
                    font=("Arial", 10), command=get_principal)
button1.pack(padx=5, pady=5, side="left")
######################################################################

######################################################################
# 创建邮件地址下拉列表
combo_box1 = ttk.Combobox(frame, values=[
    'kitty.lee@grammer.com', 'bruce.yang@grammer.com',
    'isaac.xie@grammer.com', 'jie.liu@grammer.com',
    'mark.yu@grammer.com', 'carrey.xu@grammer.com',
    'ligen.zhao@grammer.com', 'siyuan.gong@grammer.com'], state="readonly")
combo_box1.current(5)
combo_box1.pack(side="left")

# 邮件发送按钮事件


def email_sent_button_event():
    send_email(combo_box1.get())


button2 = tk.Button(frame, text="发送邮件",
                    font=("Arial", 10), command=email_sent_button_event)
button2.pack(padx=5, pady=5, side="left")
######################################################################


######################################################################
# 甘特图 待开发


# def gantt_window():
#     start = gantt[6][0]
#     end = gantt[6][1]
#     per = gantt[6][2]
#     draw_gantt_chart(start, end, per)


# # 甘特图按钮
# button3 = tk.Button(frame, text="甘特图",
#                     font=("Arial", 10), command=gantt_window)
# button3.pack(padx=15, pady=5, side="right")
######################################################################


# 设置窗口大小
window.geometry("1300x650")

# 创建表格，用于显示筛选结果
result_table = ttk.Treeview(window)

# 添加表格列
result_table["columns"] = ("项目序号", "任务名称", "责任人",
                           "开始时间", "完成时间", "计划天数", "实际完成情况")

# 设置表格列属性
result_table.column("#0", width=0, stretch=tk.NO)
result_table.column("项目序号", anchor=tk.CENTER, width=30)
result_table.column("任务名称", anchor=tk.CENTER, width=300)
result_table.column("责任人", anchor=tk.CENTER, width=66)
result_table.column("开始时间", anchor=tk.CENTER, width=80)
result_table.column("完成时间", anchor=tk.CENTER, width=80)
result_table.column("计划天数", anchor=tk.CENTER, width=30)
result_table.column("实际完成情况", anchor=tk.CENTER, width=50)


# 添加表头
result_table.heading("#0", text="", anchor=tk.W)
result_table.heading("项目序号", text="项目序号", anchor=tk.CENTER)
result_table.heading("任务名称", text="任务名称", anchor=tk.CENTER)
result_table.heading("责任人", text="责任人", anchor=tk.CENTER)
result_table.heading("开始时间", text="开始时间", anchor=tk.CENTER)
result_table.heading("完成时间", text="完成时间", anchor=tk.CENTER)
result_table.heading("计划天数", text="计划天数", anchor=tk.CENTER)
result_table.heading("实际完成情况", text="实际完成情况", anchor=tk.CENTER)


# def on_cell_edit(event):
#     """
#     处理表格单元格编辑事件的函数
#     """
#     # 获取编辑后的值
#     new_value = event.widget.get_children()[3].item(event.widget.focus())['values'][event.column]
#     print("New value:", new_value)

# # 绑定单元格编辑事件
# result_table.bind("<Double-1>", lambda event: result_table.item(event.widget.focus())['text'])

# 创建滚动条
table_scroll = tk.Scrollbar(
    window, orient="vertical", command=result_table.yview)
table_scroll.pack(side="right", fill="y")

# 将滚动条绑定到表格
result_table.configure(yscrollcommand=table_scroll.set)

#gantt = []


def filter_data(search_text):
    # 删除之前的表格数据
    result_table.delete(*result_table.get_children())
    pattern = re.compile(search_text, re.IGNORECASE)
    # 取出数据范围
    data_range = worksheet.range(
        (8, 1), (worksheet.used_range.last_cell.row, 12))
    # 读取数据到数组
    data = data_range.value

    # 迭代筛选需要的行并添加到表格中
    for row in data:
        if pattern.search(str(row[3])) or row[3] is None:

            if isinstance(row[7], datetime.datetime):
                start_time = row[7].strftime('%Y-%m-%d')
            else:
                start_time = ''

            if isinstance(row[8], datetime.datetime):
                end_time = row[8].date()
                end_time_str = end_time.strftime('%Y-%m-%d')
            else:
                end_time_str = ''

            today = datetime.date.today()
            if row[10] == None or row[10] == ' - ':
                percentage = ''
            else:
                percentage_value = float(row[10])
                percentage = f"{float(row[10])*100:.2f}%"

            #gantt.append((row[7], row[8], percentage))

            if row[0] and '.' not in row[0]:
                result_table.insert("", tk.END, text="", values=(
                    row[0], row[1], row[2], start_time, end_time_str, '', percentage))
            elif row[1] and start_time != '':
                delta = (end_time - today).days
                # if percentage_value == 1.0:
                #     fg = "black"
                # else:
                #     if delta < 0:
                #         fg = "red"
                #     else:
                #         fg = "black"
                # result_table.insert('', tk.END, text="", values=(
                #     row[0], row[1], row[3], start_time, end_time_str, row[9], percentage),
                #     tags=(fg,))
                # result_table.tag_configure("red", foreground="red")
                if percentage_value != 1.0 and delta < 0:
                    result_table.insert('', tk.END, text="", values=(
                        row[0], row[1], row[3], start_time, end_time_str, row[9], percentage),
                        tags=("red",))
                    result_table.tag_configure("red", foreground="red")
            else:
                result_table.insert('', tk.END, text="", values=(
                    row[0], row[1], row[3], start_time, end_time_str, row[9], percentage))

    # 获取所有行
    rows = result_table.get_children()
    result_table.delete(rows[-1])
    prev_item_id = None
    for item_id in result_table.get_children():
        # item_value_deletenone = result_table.item(item_id)["values"][2]
        # if str(item_value_deletenone) == 'None':
        #     result_table.delete(item_id)
        try:
            if prev_item_id is not None:
                prev_item_value = result_table.item(prev_item_id)["values"][0]
                item_value = result_table.item(item_id)["values"][0]
                if '.' not in str(prev_item_value):
                    if prev_item_value is not int:
                        prev_item_value = int(
                            str(prev_item_value).split('.')[0])
                    if item_value is not int:
                        item_value = int(str(item_value).split('.')[0])
                    if prev_item_value != item_value:
                        result_table.delete(prev_item_id)
                    # do something with prev_item_value and item_value
            prev_item_id = item_id
        except IndexError:
            # 处理访问超过长度的情况，比如跳过该循环
            pass

    red_rows = []
    a = 0
    for item_id1 in result_table.get_children():
        item_value1 = result_table.item(item_id1)["values"][0]
        if '.' not in str(item_value1):
            a += 1
            if a == 1:
                red_rows.append(result_table.item(item_id1)["values"])
            else:
                red_rows.pop()
                red_rows.append(result_table.item(item_id1)["values"])
        tags = result_table.item(item_id1, 'tags')
        if 'red' in tags:
            # do something
            a = 0
            red_rows.append(result_table.item(item_id1)["values"])

    df = pd.DataFrame(red_rows, columns=[
                      'ID', 'Project Name', 'Owner', 'Start Time', 'End Time', 'Progress', 'Completion Rate'])
    df.to_excel('List of unfinished projects.xlsx', index=False)

    # red_rows_str = "\n".join([str(row) for row in red_rows])


# def draw_gantt_chart(start_date, end_date, percent):
#     # 根据起始日期和结束日期生成日期列表
#     date_range = pd.date_range(start=start_date, end=end_date, freq='D')

#     # 生成甘特图数据
#     gantt_data = {
#         'Task': ['Task'],
#         'Start': [datetime.datetime.strftime(start_date, '%Y-%m-%d')],
#         'Finish': [datetime.datetime.strftime(end_date, '%Y-%m-%d')],
#         'Percent Complete': [percent],
#     }

#     # 绘制甘特图
#     fig, ax = plt.subplots(figsize=(8, 4))
#     ax.set_title('Gantt Chart')
#     ax.grid(True)
#     ax.set_xlabel('Date')
#     ax.set_xlim([start_date - datetime.timedelta(days=1),
#                 end_date + datetime.timedelta(days=1)])
#     ax.set_ylim([0, 1])
#     ax.xaxis_date()
#     ax.yaxis.set_visible(False)
#     ax.broken_barh([(gantt_data['Start'][0], (end_date - start_date).days)],
#                    (0.2, 0.6), facecolors='#1f77b4')
#     ax.text(end_date + datetime.timedelta(days=1), 0.5,
#             percent + 'Complete', va='center')
#     plt.show()


def send_email(receiver_email):
    # SMTP服务器的主机名和端口号
    smtp_host = 'smtp.163.com'
    smtp_port = 25

    msg = MIMEMultipart()
    msg['From'] = 'tracking_list@163.com'
    msg['To'] = receiver_email
    msg['Subject'] = 'List of unfinished projects'
    msg['Cc'] = 'Bruce.Yang@grammer.com'

    with open('List of unfinished projects.xlsx', 'rb') as f:
        attach = MIMEApplication(f.read(), _subtype='xlsx')
        attach.add_header('Content-Disposition',
                          'attachment', filename='List of unfinished projects.xlsx')
        msg.attach(attach)
    try:
        # 发送邮件
        smtp = smtplib.SMTP(smtp_host, smtp_port)
        smtp.starttls()
        smtp.login('tracking_list@163.com', 'ONTQTITQVFPGBMJV')
        smtp.sendmail('tracking_list@163.com', receiver_email, msg.as_string())
        messagebox.showinfo(parent=window, title="邮件发送成功",
                            message="邮件发送成功，请点击确认")
        smtp.quit()
    except Exception as e:
        messagebox.showerror(parent=window, title="邮件发送失败",
                             message="邮件发送失败，请检查")


# 将表格添加到窗口中
result_table.pack(pady=10, fill="both", expand=True)

# 进入主事件循环
window.mainloop()
