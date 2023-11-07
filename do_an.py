import tkinter as tk
from tkinter import ttk
import openpyxl

def load_data():
    path = "C:\Do_an_cuoi_khoa\people.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active

    list_values = list(sheet.values)
    print(list_values)
    for col_name in list_values[0]:
        treeview.heading(col_name, text=col_name)

    for value_tuple in list_values[1:]:
        treeview.insert('', tk.END, values=value_tuple)

def insert_row():
    ma_sv = masv_entry.get()
    name = name_entry.get()
    year = int(year_spinbox.get())
    gender = gender_combobox.get()
    status = "Tốt nghiệp" if a.get() else "Chưa tốt nghiệp"

    print(ma_sv, name, year, gender, status)

    # Thêm dữ liệu vào Excel sheet
    path = "C:\Do_an_cuoi_khoa\people.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    row_values = [ma_sv, name, year, gender, status]
    sheet.append(row_values)
    workbook.save(path)

    # Thêm dữ liệu vào treeview
    treeview.insert('', tk.END, values=row_values)
    
    masv_entry.delete(0, "end")
    masv_entry.insert(0, "Mã sinh viên")
    name_entry.delete(0, "end")
    name_entry.insert(0, "Họ và tên")
    year_spinbox.delete(0, "end")
    year_spinbox.insert(0, "Năm sinh")
    gender_combobox.set(combo_list[0])
    checkbutton.state(["!selected"])

root = tk.Tk()
root.title("phần mềm quản li sinh viên")
root.geometry("700x400")

#Tạo màn hình đăng nhập
def man_hinh_dang_nhap():    
    login_ui = tk.Toplevel(root)        
    login_ui.geometry("300x300")    
    login_ui.title("LOGIN") 
    
    tk.Label(login_ui, text="ĐĂNG NHẬP TÀI KHOẢN", fg="red", font="Arial 12").place(x=70, y = 10) 
    tk.Label(login_ui, text="Username:").place(x=30, y=50)    
    tk.Entry(login_ui).place(x=100, y=50)    
    tk.Label(login_ui, text="Password:").place(x=30, y=80)    
    tk.Entry(login_ui, show= '*').place(x=100, y=80)  
    tk.Button(login_ui, text="Đăng nhập", command= login_ui.destroy).place(x=110, y=130)  
    
    login_ui.mainloop()

combo_list = ["Nam", "Nữ"]

frame = ttk.Frame(root)
frame.pack()

widgets_frame = ttk.LabelFrame(frame, text="Nhập dữ liệu")
widgets_frame.grid(row=0, column=0, padx=20, pady=10)

masv_entry = ttk.Entry(widgets_frame)
masv_entry.insert(0, "Mã sinh viên")
masv_entry.grid(row=0, column=0, padx=5, pady=(0, 5))

name_entry = ttk.Entry(widgets_frame)
name_entry.insert(0, "Họ và tên")
name_entry.bind("<FocusIn>", lambda e: name_entry.delete('0', 'end'))
name_entry.grid(row=1, column=0, padx=5, pady=(0, 5))
year_spinbox = ttk.Spinbox(widgets_frame, from_=1990, to=2005)
year_spinbox.insert(0, "Năm sinh")
year_spinbox.grid(row=2, column=0, padx=5, pady=5)

gender_combobox = ttk.Combobox(widgets_frame, values=combo_list)
gender_combobox.current(0)
gender_combobox.grid(row=3, column=0, padx=5, pady=5)

a = tk.BooleanVar()
checkbutton = ttk.Checkbutton(widgets_frame, text="Tốt nghiệp", variable= a)
checkbutton.grid(row=4, column=0, padx=5, pady=5)

button = ttk.Button(widgets_frame, text="Thêm", command=insert_row)
button.grid(row=5, column=0, padx=5, pady=5)

button = ttk.Button(root, text="Thoát", command=root.quit)
button.pack()

#Tạo menu Login
menu_bar= tk.Menu(root)
root.config(menu= menu_bar)
menu_log = tk.Menu(menu_bar,tearoff = False)
menu_log.add_command(label = "Đăng nhập", command= man_hinh_dang_nhap)
menu_bar.add_cascade(label = "Tài khoản", menu = menu_log)


treeFrame = ttk.Frame(frame)
treeFrame.grid(row=0, column=1, pady=10)
treeScroll = ttk.Scrollbar(treeFrame)
treeScroll.pack(side="right", fill="y")

cols = ("Mã sinh viên","Họ và tên", "Năm sinh", "Giới tính", "Tình trạng")
treeview = ttk.Treeview(treeFrame, show="headings",
                        yscrollcommand=treeScroll.set, columns=cols, height=13)
treeview.column("Mã sinh viên", width=100)
treeview.column("Họ và tên", width=100)
treeview.column("Năm sinh", width=70)
treeview.column("Giới tính", width=100)
treeview.column("Tình trạng", width=100)
treeview.pack()
treeScroll.config(command=treeview.yview)

#Đọc dữ liệu
load_data()

root.mainloop()