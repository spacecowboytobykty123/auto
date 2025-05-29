import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from PIL import Image, ImageTk
import threading
import sys
import os
from excel_processor_dynamic import process_excel_with_dynamic_fetch

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
        process_excel_with_dynamic_fetch(file_path, output_path)
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
        # Ask user for confirmation before overwriting
        confirm = messagebox.askyesno(
            "Подтверждение",
            f"Вы действительно хотите перезаписать файл?\n{file_path}"
        )
        if not confirm:
            return

        output_path = file_path  # Overwrite the same file
        button.config(state=tk.DISABLED)

        thread = threading.Thread(target=process_file_in_thread, args=(file_path, output_path, button))
        thread.start()
def main():
    root = tk.Tk()
    root.title("Excel Обработчик")
    root.geometry("640x500")
    root.configure(bg="#f4f4f4")

    # ======== Logo ========
    try:
        logo_img = Image.open("nitec.png")  # Use your company logo file here
        logo_img = logo_img.resize((200, 120), Image.Resampling.BILINEAR)
        logo = ImageTk.PhotoImage(logo_img)
        logo_label = tk.Label(root, image=logo, bg="#f4f4f4")
        logo_label.image = logo
        logo_label.pack(pady=10)
    except Exception as e:
        print("Логотип не найден или поврежден:", e)

    # ======== Header ========
    label = tk.Label(root, text="Nitec Monitoring", font=("Arial", 14, "bold"), bg="#f4f4f4")
    label.pack(pady=10)

    # ======== Button ========
    process_button = tk.Button(root, text="📂 Выбрать Excel файл", font=("Arial", 12), bg="#4CAF50", fg="white", padx=10, pady=5)
    process_button.config(command=lambda: select_file(process_button))
    process_button.pack(pady=5)

    # ======== Log Output ========
    log_box = scrolledtext.ScrolledText(root, height=15, width=80, font=("Courier New", 10))
    log_box.pack(pady=10, padx=10)

    # ======== Redirect Output ========
    sys.stdout = sys.stderr = RedirectText(log_box)

    root.mainloop()

if __name__ == "__main__":
    main()
