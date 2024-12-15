import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import csv
import pandas as pd

FILE_NAME = "employee_data.csv"

def save_to_csv(data):
    try:
        with open(FILE_NAME, mode="a", newline="", encoding="utf-8") as file:
            writer = csv.writer(file)
            writer.writerow(data)
    except Exception as e:
        messagebox.showerror("Error", f"Không thể lưu dữ liệu: {e}")


def show_birthdays_today():
    try:
        today = datetime.today().strftime("%d/%m/%Y")
        with open(FILE_NAME, mode="r", encoding="utf-8") as file:
            reader = csv.reader(file)
            next(reader)
            birthdays = [row for row in reader if row[4] == today]
        if birthdays:
            messagebox.showinfo("Sinh nhật hôm nay", "\n".join([f"{row[1]} ({row[0]})" for row in birthdays]))
        else:
            messagebox.showinfo("Sinh nhật hôm nay", "Không có nhân viên nào sinh nhật hôm nay.")
    except FileNotFoundError:
        messagebox.showwarning("Warning", "Chưa có dữ liệu nhân viên.")


def export_to_excel():
    try:
        data = pd.read_csv(FILE_NAME)
        data["Ngày sinh"] = pd.to_datetime(data["Ngày sinh"], format="%d/%m/%Y")
        data = data.sort_values(by="Ngày sinh", ascending=False)
        output_file = "employee_data.xlsx"
        data.to_excel(output_file, index=False)
        messagebox.showinfo("Success", f"Dữ liệu đã được xuất ra file {output_file}.")
    except Exception as e:
        messagebox.showerror("Error", f"Không thể xuất dữ liệu: {e}")


def add_employee():
    emp_id = entry_id.get()
    name = entry_name.get()
    department = combobox_department.get()
    position = combobox_position.get()
    dob = entry_dob.get()
    gender = gender_var.get()
    id_card = entry_id_card.get()
    issue_place = entry_issue_place.get()

    if not (emp_id and name and department and dob and gender):
        messagebox.showwarning("Warning", "Vui lòng nhập đầy đủ các trường bắt buộc.")
        return

    try:
        datetime.strptime(dob, "%d/%m/%Y")
    except ValueError:
        messagebox.showerror("Error", "Ngày sinh không đúng định dạng DD/MM/YYYY.")
        return

    save_to_csv([emp_id, name, department, position, dob, gender, id_card, issue_place])
    messagebox.showinfo("Success", "Thêm nhân viên thành công!")
    clear_form()


def clear_form():
    entry_id.delete(0, tk.END)
    entry_name.delete(0, tk.END)
    combobox_department.set("")
    combobox_position.set("")
    entry_dob.delete(0, tk.END)
    gender_var.set("Nam")
    entry_id_card.delete(0, tk.END)
    entry_issue_place.delete(0, tk.END)


root = tk.Tk()
root.title("Quản lý nhân viên")

frame_form = ttk.LabelFrame(root, text="Thông tin nhân viên")
frame_form.grid(column=0, row=0, padx=10, pady=10, sticky="NSEW")

ttk.Label(frame_form, text="Mã*:").grid(column=0, row=0, padx=5, pady=5, sticky="E")
entry_id = ttk.Entry(frame_form)
entry_id.grid(column=1, row=0, padx=5, pady=5, sticky="W")

ttk.Label(frame_form, text="Tên*:").grid(column=0, row=1, padx=5, pady=5, sticky="E")
entry_name = ttk.Entry(frame_form)
entry_name.grid(column=1, row=1, padx=5, pady=5, sticky="W")

ttk.Label(frame_form, text="Đơn vị*:").grid(column=0, row=2, padx=5, pady=5, sticky="E")
combobox_department = ttk.Combobox(frame_form, values=["Phân xưởng que hàn"])
combobox_department.grid(column=1, row=2, padx=5, pady=5, sticky="W")

ttk.Label(frame_form, text="Chức danh:").grid(column=0, row=3, padx=5, pady=5, sticky="E")
combobox_position = ttk.Combobox(frame_form, values=["Nhân viên"])
combobox_position.grid(column=1, row=3, padx=5, pady=5, sticky="W")

ttk.Label(frame_form, text="Ngày sinh (DD/MM/YYYY):").grid(column=0, row=4, padx=5, pady=5, sticky="E")
entry_dob = ttk.Entry(frame_form)
entry_dob.grid(column=1, row=4, padx=5, pady=5, sticky="W")

ttk.Label(frame_form, text="Giới tính:").grid(column=0, row=5, padx=5, pady=5, sticky="E")
gender_var = tk.StringVar(value="Nam")
ttk.Radiobutton(frame_form, text="Nam", variable=gender_var, value="Nam").grid(column=1, row=5, padx=5, pady=5,
                                                                               sticky="W")
ttk.Radiobutton(frame_form, text="Nữ", variable=gender_var, value="Nữ").grid(column=2, row=5, padx=5, pady=5,
                                                                             sticky="W")

ttk.Label(frame_form, text="Số CMND:").grid(column=0, row=6, padx=5, pady=5, sticky="E")
entry_id_card = ttk.Entry(frame_form)
entry_id_card.grid(column=1, row=6, padx=5, pady=5, sticky="W")

ttk.Label(frame_form, text="Nơi cấp:").grid(column=0, row=7, padx=5, pady=5, sticky="E")
entry_issue_place = ttk.Entry(frame_form)
entry_issue_place.grid(column=1, row=7, padx=5, pady=5, sticky="W")

frame_buttons = ttk.Frame(root)
frame_buttons.grid(column=0, row=1, padx=10, pady=10, sticky="EW")

btn_add = ttk.Button(frame_buttons, text="Thêm", command=add_employee)
btn_add.grid(column=0, row=0, padx=5, pady=5)

btn_birthday = ttk.Button(frame_buttons, text="Sinh nhật hôm nay", command=show_birthdays_today)
btn_birthday.grid(column=1, row=0, padx=5, pady=5)

btn_export = ttk.Button(frame_buttons, text="Xuất danh sách", command=export_to_excel)
btn_export.grid(column=2, row=0, padx=5, pady=5)

root.mainloop()