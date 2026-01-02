print("App started")
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
root = tk.Tk()
root.title("Excel Viewer")
root.geometry("900x500")
upload_btn = tk.Button(
    root,
    text = "Upload Excel File",
    command=lambda: None,
    font = ("Arial", 12),
    bg ="green",
    fg= "white"
)
upload_btn.pack(pady=10)
frame = tk.Frame(root)
frame.pack (fill = "both", expand = True)
tree = ttk.Treeview(frame)
tree.pack (side="left", fill="both", expand = True)
y_scroll = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
x_scroll = ttk.Scrollbar(root, orient="horizontal", command=tree.xview)
tree.configure(
    yscrollcommand=y_scroll.set,
    xscrollcommand=x_scroll.set
)
y_scroll.pack(side="right", fill="y")
x_scroll.pack(fill="x")
def load_excel():
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if file_path == "":
        return
    try:
        df=pd.read_excel(file_path)
        display_data(df)
    except Execution as e:
        messagebox.showerror ("Error", str(e))
def display_data(df):
    tree.delete(*tree.get_children())
    tree["columns"]=list(df.columns)
    tree["show"]="headings"
    for col in df.columns:
        tree.heading(col, text=col)
        tree.column(col, width=120)
    for _, row in df.iterrows():
        tree.insert("", "end", values=list(row))
upload_btn.config(command=load_excel)
root.mainloop()
