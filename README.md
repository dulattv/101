
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

def read_dataframe(file_path, skip_header):
    read_opts = {"header": None}
    if skip_header:
        read_opts["skiprows"] = [0]

    if file_path.endswith(".csv"):
        return pd.read_csv(file_path, **read_opts)
    elif file_path.endswith(".xlsx"):
        return pd.read_excel(file_path, engine="openpyxl", **read_opts)
    else:
        raise Exception("Поддерживаются только .csv и .xlsx файлы")

def create_folders(file_path, base_dir, column_index, nested, skip_header, status_label):
    try:
        df = read_dataframe(file_path, skip_header)

        if df.shape[1] < 1:
            raise Exception("Файл не содержит данных")

        if column_index == -1:
            raise Exception("Колонка не выбрана")

        if column_index >= df.shape[1]:
            raise Exception(f"В файле всего {df.shape[1]} колонок, а выбрана Колонка {column_index}")

        created = []
        for folder_name in df.iloc[:, column_index]:
            name = str(folder_name).strip()
            if not name:
                continue
            folder_path = os.path.join(base_dir, *name.split("/")) if nested else os.path.join(base_dir, name)
            os.makedirs(folder_path, exist_ok=True)
            created.append(folder_path)

        status_label.config(text=f"✅ Создано: {len(created)} папок", fg="green")
    except Exception as e:
        status_label.config(text=f"❌ Ошибка: {e}", fg="red")

def update_preview(file_path, skip_header, column_box, preview_box):
    try:
        df = read_dataframe(file_path, skip_header)
        column_box["values"] = [f"Колонка {i}" for i in range(df.shape[1])]
        column_box.current(0)
        preview_box.delete(0, tk.END)
        preview_box.insert(tk.END, *df.head(10).astype(str).agg(" | ".join, axis=1).values)
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось прочитать файл: {e}")

def browse_file(entry, skip_header_var, column_box, preview_box):
    file_path = filedialog.askopenfilename(filetypes=[("CSV/Excel files", "*.csv *.xlsx")])
    if file_path:
        entry.delete(0, tk.END)
        entry.insert(0, file_path)
        update_preview(file_path, skip_header_var.get(), column_box, preview_box)

def browse_folder(entry):
    folder_path = filedialog.askdirectory()
    if folder_path:
        entry.delete(0, tk.END)
        entry.insert(0, folder_path)

def build_ui():
    root = tk.Tk()
    root.title("NameFolder")
    root.geometry("580x560")
    root.resizable(False, False)
    root.configure(bg="#f8f8f8")

    font_label = ("Segoe UI", 10)
    font_entry = ("Segoe UI", 10)
    font_button = ("Segoe UI", 10, "bold")

    def add_labeled_row(label_text, browse_command=None):
        frame = tk.Frame(root, bg="#f8f8f8")
        frame.pack(padx=20, pady=5, fill="x")
        label = tk.Label(frame, text=label_text, bg="#f8f8f8", font=font_label)
        label.pack(anchor="w")
        entry = tk.Entry(frame, font=font_entry)
        entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        if browse_command:
            button = tk.Button(frame, text="Обзор", font=font_button, command=lambda: browse_command(entry))
            button.pack(side="right")
        return entry, frame

    csv_entry, csv_frame = add_labeled_row("Файл (.csv или .xlsx):")
    folder_entry, _ = add_labeled_row("Папка назначения:", browse_folder)

    column_frame = tk.Frame(root, bg="#f8f8f8")
    column_frame.pack(padx=20, pady=5, fill="x")
    tk.Label(column_frame, text="Выбор колонки:", bg="#f8f8f8", font=font_label).pack(anchor="w")
    column_box = ttk.Combobox(column_frame, state="readonly")
    column_box.pack(fill="x")

    preview_frame = tk.Frame(root, bg="#f8f8f8")
    preview_frame.pack(padx=20, pady=10, fill="both", expand=False)
    tk.Label(preview_frame, text="Предпросмотр данных:", bg="#f8f8f8", font=font_label).pack(anchor="w")
    preview_box = tk.Listbox(preview_frame, height=6, font=("Courier", 9))
    preview_box.pack(fill="both", expand=True)

    nested_var = tk.BooleanVar()
    skip_header_var = tk.BooleanVar()

    def on_skip_header_toggle():
        file_path = csv_entry.get()
        if file_path:
            update_preview(file_path, skip_header_var.get(), column_box, preview_box)

    options_frame = tk.Frame(root, bg="#f8f8f8")
    options_frame.pack(padx=20, pady=5, fill="x")
    tk.Checkbutton(options_frame, text="Использовать вложенные папки (по / )", variable=nested_var,
                   bg="#f8f8f8", font=font_label).pack(anchor="w")
    tk.Checkbutton(options_frame, text="Пропустить первую строку", variable=skip_header_var,
                   bg="#f8f8f8", font=font_label, command=on_skip_header_toggle).pack(anchor="w")

    status_label = tk.Label(root, text="", bg="#f8f8f8", font=font_label)
    status_label.pack(pady=5)

    tk.Button(csv_frame, text="Обзор", font=font_button,
              command=lambda: browse_file(csv_entry, skip_header_var, column_box, preview_box)).pack(side="right")

    tk.Button(
        root, text="Создать папки", font=font_button, bg="#444", fg="white",
        height=2, width=20,
        command=lambda: create_folders(
            csv_entry.get(), folder_entry.get(), column_box.current(),
            nested_var.get(), skip_header_var.get(), status_label)
    ).pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    build_ui()
