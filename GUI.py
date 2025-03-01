import tkinter as tk
from tkinter import messagebox
import subprocess

def run_bdu_parser():
    try:
        subprocess.run(["python", "BDU_parser.py"], check=True)
        messagebox.showinfo("Успех", "BDU_parser завершён успешно!")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Ошибка", f"Ошибка при выполнении BDU_parser: {e}")

def run_capec_parser():
    try:
        subprocess.run(["python", "CAPEC_parser.py"], check=True)
        messagebox.showinfo("Успех", "CAPEC_parser завершён успешно!")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Ошибка", f"Ошибка при выполнении CAPEC_parser: {e}")

def run_cwe_parser():
    try:
        subprocess.run(["python", "CWE_parser.py"], check=True)
        messagebox.showinfo("Успех", "CWE_parser завершён успешно!")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Ошибка", f"Ошибка при выполнении CWE_parser: {e}")

def run_capec_translation():
    try:
        subprocess.run(["python", "CAPEC_translation.py"], check=True)
        messagebox.showinfo("Успех", "CAPEC_translation завершён успешно!")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Ошибка", f"Ошибка при выполнении CAPEC_translation: {e}")

# Создание основного окна
root = tk.Tk()
root.title("Управление парсерами")
root.geometry("300x200")

# Создание кнопок
btn_bdu_parser = tk.Button(root, text="1. Запустить BDU_parser", command=run_bdu_parser)
btn_capec_parser = tk.Button(root, text="2. Запустить CAPEC_parser", command=run_capec_parser)
btn_cwe_parser = tk.Button(root, text="3. Запустить CWE_parser", command=run_cwe_parser)
btn_capec_translation = tk.Button(root, text="Запустить CAPEC_translation", command=run_capec_translation)

# Размещение кнопок в окне
btn_bdu_parser.pack(pady=10)
btn_capec_parser.pack(pady=10)
btn_cwe_parser.pack(pady=10)
btn_capec_translation.pack(pady=10)

# Запуск основного цикла обработки событий
root.mainloop()