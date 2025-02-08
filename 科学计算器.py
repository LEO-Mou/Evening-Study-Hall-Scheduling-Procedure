import tkinter as tk
import math

root = tk.Tk()
root.title("科学计算器")

entry_font = ('等线', 14)
button_font = ('等线', 12)

root.configure(bg='#f0f0f0')

entry = tk.Entry(root, width=30, font=entry_font)
entry.grid(row=0, column=0, columnspan=6, padx=10, pady=10)

def button_click(number):
    current = entry.get()
    entry.delete(0, tk.END)
    entry.insert(0, str(current) + str(number))

def clear():
    entry.delete(0, tk.END)

def clear_all():
    clear()

def calculate():
    try:
        result = eval(entry.get())
        entry.delete(0, tk.END)
        entry.insert(0, result)
    except:
        entry.delete(0, tk.END)
        entry.insert(0, "错误")

def scientific_function(func):
    try:
        number = float(entry.get())
        if func == "sin":
            result = math.sin(math.radians(number))
        elif func == "cos":
            result = math.cos(math.radians(number))
        elif func == "tan":
            result = math.tan(math.radians(number))
        elif func == "sqrt":
            if number < 0:
                raise ValueError("不能对负数开平方根")
            result = math.sqrt(number)
        elif func == "x²":
            result = number**2
        elif func == "√":
            if number < 0:
                raise ValueError("不能对负数开平方根")
            result = math.sqrt(number)
        elif func == "x³":
            result = number**3
        entry.delete(0, tk.END)
        entry.insert(0, result)
    except ValueError as e:
        entry.delete(0, tk.END)
        entry.insert(0, str(e))
    except:
        entry.delete(0, tk.END)
        entry.insert(0, "错误")

def backspace():
    current = entry.get()
    entry.delete(0, tk.END)
    entry.insert(0, current[:-1])

def add_minus_sign():
    current = entry.get()
    if current and (current[0] not in ('-', '+')):
        entry.insert(0, '-')

buttons = [
    '7', '8', '9', '/',
    '4', '5', '6', '*',
    '1', '2', '3', '-',
    '0', '.', '=', '+',
]
row_val = 1
col_val = 0
for button in buttons:
    btn = tk.Button(root, text=button, font=button_font, padx=15, pady=15, command=lambda b=button: button_click(b))
    btn.grid(row=row_val, column=col_val, padx=5, pady=5)
    col_val += 1
    if col_val > 3:
        col_val = 0
        row_val += 1

clear_btn = tk.Button(root, text='C', font=button_font, padx=15, pady=15, command=clear)
clear_btn.grid(row=row_val, column=col_val, padx=5, pady=5)

backspace_btn = tk.Button(root, text='←', font=button_font, padx=15, pady=15, command=backspace)
backspace_btn.grid(row=row_val, column=col_val + 1, padx=5, pady=5)

minus_sign_btn = tk.Button(root, text='±', font=button_font, padx=15, pady=15, command=add_minus_sign)
minus_sign_btn.grid(row=row_val, column=col_val + 2, padx=5, pady=5)

scientific_buttons = [
    ('sin', 5, 0),
    ('cos', 5, 1),
    ('tan', 5, 2),
    ('sqrt', 5, 3),
    ('x²', 5, 4),
    ('√', 5, 5),
    ('x³', 5, 6),
]
for func, row, col in scientific_buttons:
    btn = tk.Button(root, text=func, font=button_font, padx=15, pady=15, command=lambda f=func: scientific_function(f))
    btn.grid(row=row, column=col, padx=5, pady=5)

clear_all_btn = tk.Button(root, text='清除所有', font=button_font, padx=15, pady=15, command=clear_all)
clear_all_btn.grid(row=row + 1, column=0, columnspan=3, padx=5, pady=5)

root.bind('<Return>', lambda event: calculate())

description_label = tk.Label(root, text="若按键失效，请使用键盘操作；当进行函数计算时，请先按下数字，再按下功能性按键", font=('等线', 10))
description_label.grid(row=7, column=0, columnspan=6, padx=10, pady=5)

entry.focus_set()

root.mainloop()