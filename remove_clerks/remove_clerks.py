import pandas as pd
from openpyxl import load_workbook
from tkinter import filedialog, Tk, Frame, Button, Scrollbar, Checkbutton, IntVar

def read_excel():
    file_path = filedialog.askopenfilename(title="选择Excel文件", filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        df = pd.read_excel(file_path, engine='openpyxl')
        if "店员" in df.columns:
            display_clerks(df, file_path)
        else:
            print("店员列未找到")
    else:
        print("未选择文件")

def display_clerks(df, file_path):
    clerks = df["店员"].unique()

    def on_select():
        selected_clerks = [clerk for idx, clerk in enumerate(clerks) if var_list[idx].get()]
        filtered_df = df[df["店员"].isin(selected_clerks)]
        output_path = file_path.replace('.xlsx', '销售数据（离职店员已被摘除）.xlsx')
        filtered_df.to_excel(output_path, index=False, engine='openpyxl')
        print("输出文件已保存:", output_path)

    for widget in frame.winfo_children():
        widget.destroy()

    var_list = [IntVar(value=1) for _ in clerks]
    for idx, clerk in enumerate(clerks):
        checkbutton = Checkbutton(frame, text=clerk, variable=var_list[idx])
        checkbutton.pack(anchor="w")

    select_btn.config(text="导出数据", command=on_select)

root = Tk()
root.title("选择店员")

frame = Frame(root)
frame.pack(pady=10, fill="both", expand=True)

scrollbar = Scrollbar(frame, orient="vertical")
scrollbar.pack(side="right", fill="y")

select_btn = Button(root, text="选择销售数据", command=lambda: read_excel())
select_btn.pack(pady=5)

root.mainloop()
