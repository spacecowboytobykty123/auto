import tkinter as tk
import tkinter.ttk as ttk  # ✅ Add this at the top of your script
from tkinter import filedialog, messagebox, scrolledtext
from PIL import Image, ImageTk
import threading
import sys
import os

# ✅ Import the unified pipeline function
from main import process_combined_excel_pipeline


class RedirectText:
    def __init__(self, text_widget):
        self.output = text_widget

    def write(self, string):
        self.output.insert(tk.END, string)
        self.output.see(tk.END)

    def flush(self):
        pass


def process_file_in_thread(file_path, output_path, button):
    try:
        print("⏳ Обработка файла началась...\n")
        process_combined_excel_pipeline(file_path, output_path)
        messagebox.showinfo("Готово", f"Файл обработан успешно.\nРезультат: {output_path}")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Что-то пошло не так:\n{e}")
    finally:
        button.config(state=tk.NORMAL)


def select_file(button):
    file_path = filedialog.askopenfilename(
        title="Выберите Excel файл",
        filetypes=[("Excel Files", "*.xlsx")]
    )
    if file_path:
        confirm = messagebox.askyesno(
            "Подтверждение",
            f"Вы действительно хотите перезаписать файл?\n{file_path}"
        )
        if not confirm:
            return

        output_path = file_path  # Overwrite
        button.config(state=tk.DISABLED)

        thread = threading.Thread(
            target=process_file_in_thread,
            args=(file_path, output_path, button)
        )
        thread.start()

def main():
    root = tk.Tk()
    root.title("NITEC: HTML + DB Excel Анализ")
    root.geometry("750x750")
    root.configure(bg="#f0f2f5")

    # === Title ===
    label = tk.Label(
        root,
        text="NITEC: HTML + DB Excel Анализ",
        font=("Segoe UI", 18, "bold"),
        bg="#f0f2f5",
        fg="#222"
    )
    label.pack(pady=(10, 5))

    # === Create Notebook Tabs ===
    notebook = ttk.Notebook(root)
    notebook.pack(padx=10, pady=10, expand=True, fill="both")

    main_tab = tk.Frame(notebook, bg="#f0f2f5")
    db_tab = tk.Frame(notebook, bg="#ffffff")

    notebook.add(main_tab, text="📁 Главная")
    notebook.add(db_tab, text="⚙️ Настройки базы данных")

    # === DB Settings Tab ===
    db_entries = {}
    fields = {
        "host": "Хост",
        "port": "Порт",
        "dbname": "Имя базы",
        "user": "Пользователь",
        "password": "Пароль"
    }

    center_frame = tk.Frame(db_tab, bg="#ffffff")
    center_frame.pack(anchor="center", pady=20)

    form_frame = tk.Frame(center_frame, bg="#ffffff", bd=1, relief="solid")
    form_frame.pack(padx=10, pady=10)

    title_label = tk.Label(
        form_frame,
        text="Настройки подключения к базе данных",
        font=("Segoe UI", 12, "bold"),
        bg="#ffffff"
    )
    title_label.pack(pady=(10, 15))

    for key, label_text in fields.items():
        row = tk.Frame(form_frame, bg="#ffffff")
        row.pack(fill="x", padx=20, pady=5)
        tk.Label(row, text=label_text + ":", width=15, anchor="w", font=("Segoe UI", 10), bg="#ffffff").pack(
            side="left")
        entry = tk.Entry(row, show="*" if key == "password" else None, font=("Segoe UI", 10), width=30)
        entry.pack(side="left")
        db_entries[key] = entry

    # Default values
    db_entries["host"].insert(0, "192.168.175.27")
    db_entries["port"].insert(0, "5432")
    db_entries["dbname"].insert(0, "egov")
    db_entries["user"].insert(0, "alisher_ibrayev")
    db_entries["password"].insert(0, "ASTkazkorp2010!@#")

    # === File Select and Processing ===
    def select_file_with_db(button):
        file_path = filedialog.askopenfilename(
            title="Выберите Excel файл",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if file_path:
            confirm = messagebox.askyesno(
                "Подтверждение",
                f"Вы действительно хотите перезаписать файл?\n{file_path}"
            )
            if not confirm:
                return

            output_path = file_path
            button.config(state=tk.DISABLED)

            db_config = {k: db_entries[k].get() for k in db_entries}
            db_config["port"] = int(db_config["port"])

            thread = threading.Thread(
                target=process_file_in_thread,
                args=(file_path, output_path, button, db_config)
            )
            thread.start()

    def process_file_in_thread(file_path, output_path, button, db_config):
        try:
            print("⏳ Обработка файла началась...\n")
            process_combined_excel_pipeline(file_path, output_path, db_config)
            messagebox.showinfo("Готово", f"Файл обработан успешно.\nРезультат: {output_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Что-то пошло не так:\n{e}")
        finally:
            button.config(state=tk.NORMAL)

    # === Main Tab Content ===
    process_button = tk.Button(
        main_tab,
        text="📂 Выбрать Excel файл",
        font=("Segoe UI", 12, "bold"),
        bg="#4CAF50",
        fg="white",
        activebackground="#45a049",
        padx=20,
        pady=10,
        relief="raised",
        bd=2,
        command=lambda: select_file_with_db(process_button)
    )
    process_button.pack(pady=20)

    log_box = scrolledtext.ScrolledText(
        main_tab,
        height=22,
        width=90,
        font=("Courier New", 10),
        bg="#ffffff"
    )
    log_box.pack(pady=10, padx=20, fill="both", expand=True)

    sys.stdout = sys.stderr = RedirectText(log_box)

    root.mainloop()



if __name__ == "__main__":
    main()
