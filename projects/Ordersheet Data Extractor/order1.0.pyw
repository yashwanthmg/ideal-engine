import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from tkinter import messagebox
import pandas as pd


def select_file1():
    file_path1 = filedialog.askopenfilename()
    file_path1_var.set(file_path1)


def select_file2():
    file_path2 = filedialog.askopenfilename()
    file_path2_var.set(file_path2)


def extract_data():
    file1_path = file_path1_var.get()
    file2_path = file_path2_var.get()

    if not file1_path or not file2_path:
        messagebox.showerror("Error", "Please select both files.")
        return

    try:
        file1 = pd.read_excel(file1_path, engine='openpyxl')
        ids = file1.iloc[:, 0].tolist()

        file2 = pd.read_excel(file2_path, engine='openpyxl')
        file2_filtered = file2[file2['business_account_id'].isin(ids)][
            ['order_id', 'units', 'business_name', 'ops_without_tax', 'order_day']
        ]

        save_path = filedialog.asksaveasfilename(defaultextension='.csv')
        if save_path:
            if not save_path.endswith('.csv'):
                save_path += '.csv'
            file2_filtered.to_csv(save_path, index=False)
            messagebox.showinfo("Done", "Data extraction complete!")
            file_path1_var.set("")
            file_path2_var.set("")
    except Exception as e:
        messagebox.showerror("Error", str(e))


root = tk.Tk()
root.title('OS Data Extractor')

file_path1_var = tk.StringVar()
file_path2_var = tk.StringVar()

file1_label = tk.Label(root, text='Select Business ID File:')
file1_label.grid(row=0, column=0, padx=10, pady=10)
file1_entry = tk.Entry(root, textvariable=file_path1_var)
file1_entry.grid(row=0, column=1, padx=10, pady=10)
file1_button = tk.Button(root, text='Browse', command=select_file1)
file1_button.grid(row=0, column=2, padx=10, pady=10)

file2_label = tk.Label(root, text='Select Order Sheet File:')
file2_label.grid(row=1, column=0, padx=10, pady=10)
file2_entry = tk.Entry(root, textvariable=file_path2_var)
file2_entry.grid(row=1, column=1, padx=10, pady=10)
file2_button = tk.Button(root, text='Browse', command=select_file2)
file2_button.grid(row=1, column=2, padx=10, pady=10)

extract_button = tk.Button(root, text='Extract Data', command=extract_data)
extract_button.grid(row=2, column=1, padx=10, pady=10)

root.mainloop()
