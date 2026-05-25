import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import xlwings as xw
import os

file1 = None
file2 = None
template_file = "template_12k_20_45k.xlsx"
output_file = "result.xlsx"


def choose_file1():
    global file1
    file1 = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if file1:
        btn1.config(text="1-й файл выбран")
        lbl_file1.config(text=os.path.basename(file1))


def choose_file2():
    global file2
    file2 = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if file2:
        btn2.config(text="2-й файл выбран")
        lbl_file2.config(text=os.path.basename(file2))


def choose_template():
    global template_file
    template_file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if template_file:
        lbl_template.config(text=os.path.basename(template_file))


def choose_output_file():
    global output_file
    output_file = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        initialfile="result.xlsx"
    )
    if output_file:
        lbl_output.config(text=os.path.basename(output_file))

def generate_excel():
    if not file1 or not file2:
        messagebox.showerror("Ошибка", "Выберите оба CSV файла")
        return

    # читаем CSV
    # df1 = pd.read_csv(file1, header=None, delimiter=";")
    # df2 = pd.read_csv(file2, header=None, delimiter=";")
    #
    # x = df1.iloc[:, 0]
    # y1 = df1.iloc[:, 1]
    # y2 = df2.iloc[:, 1]
    # diff = y1 - y2

    df1 = pd.read_csv(file1, header=None, encoding="utf-8-sig", delimiter=";")
    df2 = pd.read_csv(file2, header=None, encoding="utf-8-sig", delimiter=";")

    # Переименуем столбцы для удобства
    df1.columns = ["x", "y1"]
    df2.columns = ["x", "y2"]

    # Делаем полное объединение по X
    df = pd.merge(df1, df2, on="x", how="outer")

    # Сортируем по X (обязательно!)
    df = df.sort_values("x")

    # Считаем разницу (если нет пары — будет NaN)
    df["diff"] = df["y1"] - df["y2"]

    # открываем шаблон через xlwings (НЕ ломает элементы управления)
    app = xw.App(visible=False)
    wb = app.books.open(template_file)
    ws = wb.sheets[0]

    # записываем имена файлов
    ws["B1"].value = file1
    ws["B2"].value = file2

    # формируем DataFrame для вставки
    # df = pd.DataFrame({
    #     "Freq": x,
    #     "Data 1": y1,
    #     "Data 2": y2,
    #     "diff": diff
    # })

    # очищаем старые данные (A7:D30000)
    ws.range("A7:D30000").clear_contents()

    # вставляем новые данные
    #ws["A6"].value = df

    ws.range("A7:D30000").value = None
    ws["A7"].options(index=False, header=False).value = df

    # сохраняем результат
    # wb.save("result.xlsx")
    wb.save(output_file)
    wb.close()
    app.quit()

    messagebox.showinfo("Готово", f"Файл сохранен как:\n{output_file}")
    if output_file:
        lbl_btn3.config(text=os.path.basename(output_file))

# -----------------------------
# GUI
# -----------------------------
root = tk.Tk()
root.title("Генератор Excel по шаблону (xlwings)")
root.geometry("420x320")

btn1 = tk.Button(root, text="Выбрать 1-й CSV", command=choose_file1)
btn1.pack(pady=5)
lbl_file1 = tk.Label(root, text="(не выбран)")
lbl_file1.pack()

btn2 = tk.Button(root, text="Выбрать 2-й CSV", command=choose_file2)
btn2.pack(pady=5)
lbl_file2 = tk.Label(root, text="(не выбран)")
lbl_file2.pack()

btn_template = tk.Button(root, text="Выбрать шаблон Excel", command=choose_template)
btn_template.pack(pady=10)
lbl_template = tk.Label(root, text="template_12k_20_45k.xlsx")
lbl_template.pack()

btn_output = tk.Button(root, text="Выбрать выходной файл", command=choose_output_file)
btn_output.pack(pady=10)
lbl_output = tk.Label(root, text="result.xlsx")
lbl_output.pack()

btn3 = tk.Button(root, text="Сгенерировать", command=generate_excel)
btn3.pack(pady=10)
lbl_btn3 = tk.Label(root, text="(не сгенерирован)")
lbl_btn3.pack()

root.mainloop()
