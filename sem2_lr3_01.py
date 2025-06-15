import tkinter as tk
from tkinter import ttk
from tkinter import scrolledtext
import numpy as np
import openpyxl

# --- МАТРИЦЫ ИЗ ЗАДАНИЯ (system.png) ---
# Исходная матрица коэффициентов A
A_matrix = np.array([
    [1.42, 0.32, -0.42, 0.85],
    [0.63, -0.43, 1.27, -0.58],
    [0.84, -2.23, -0.52, 0.47],
    [0.27, 1.37, 0.64, -1.27]
], dtype=float)

# Вектор свободных членов B
B_vector = np.array([1.32, -0.44, 0.64, 0.85], dtype=float)

# --- ПРОЦЕДУРЫ РЕШЕНИЯ СИСТЕМЫ ---

def gauss_method(A, b):
    """
    Решает систему линейных уравнений A*x = b методом Гаусса.
    Возвращает вектор решения x.
    """
    n = len(b)
    # Создаем копии, чтобы не изменять исходные матрицы
    a = np.copy(A)
    b = np.copy(b)

    # Прямой ход (приведение к верхнетреугольному виду)
    for i in range(n):
        # Находим главный элемент в текущем столбце для повышения точности
        pivot_row = i
        for j in range(i + 1, n):
            if abs(a[j, i]) > abs(a[pivot_row, i]):
                pivot_row = j
        a[[i, pivot_row]] = a[[pivot_row, i]]
        b[[i, pivot_row]] = b[[pivot_row, i]]

        # Делим текущую строку на ведущий элемент
        pivot = a[i, i]
        a[i] = a[i] / pivot
        b[i] = b[i] / pivot

        # Вычитаем текущую строку из всех последующих
        for j in range(i + 1, n):
            factor = a[j, i]
            a[j] = a[j] - factor * a[i]
            b[j] = b[j] - factor * b[i]

    # Обратный ход (нахождение решения)
    x = np.zeros(n)
    for i in range(n - 1, -1, -1):
        x[i] = b[i] - np.dot(a[i, i+1:], x[i+1:])
        
    return x

def gauss_jordan_method(A, b):
    """
    Решает систему линейных уравнений A*x = b методом Гаусса-Жордана.
    Возвращает вектор решения x.
    """
    n = len(b)
    # Создаем расширенную матрицу [A|b]
    augmented_matrix = np.hstack([A, b.reshape(-1, 1)])

    # Прямой и обратный ход (приведение к единичной матрице)
    for i in range(n):
        # Находим главный элемент в текущем столбце
        pivot_row = i
        for j in range(i + 1, n):
            if abs(augmented_matrix[j, i]) > abs(augmented_matrix[pivot_row, i]):
                pivot_row = j
        augmented_matrix[[i, pivot_row]] = augmented_matrix[[pivot_row, i]]

        # Нормализуем строку, чтобы на диагонали была 1
        pivot = augmented_matrix[i, i]
        augmented_matrix[i] = augmented_matrix[i] / pivot

        # Обнуляем элементы в текущем столбце (кроме диагонального)
        for j in range(n):
            if i != j:
                factor = augmented_matrix[j, i]
                augmented_matrix[j] = augmented_matrix[j] - factor * augmented_matrix[i]

    # Решение находится в последнем столбце расширенной матрицы
    x = augmented_matrix[:, n]
    return x

# --- ФУНКЦИИ ИНТЕРФЕЙСА И ВЫВОДА ---

def solve_system():
    """
    Выполняет решение системы выбранным методом и выводит результат.
    """
    method = method_selector.get()
    
    if method == "Метод Гаусса":
        solution = gauss_method(A_matrix, B_vector)
    elif method == "Метод Гаусса-Жордана":
        solution = gauss_jordan_method(A_matrix, B_vector)
    else:
        solution = None

    # Очищаем текстовое поле перед новым выводом
    output_text.delete('1.0', tk.END)

    # Формируем строку с результатами
    result_string = f"--- РЕЗУЛЬТАТЫ РАСЧЕТА ---\n\n"
    result_string += f"Выбранный метод: {method}\n\n"
    result_string += "Исходная матрица A:\n"
    result_string += str(A_matrix) + "\n\n"
    result_string += "Исходный вектор B:\n"
    result_string += str(B_vector) + "\n\n"
    
    if solution is not None:
        result_string += "Найденное решение X:\n"
        # Форматируем вывод решения для наглядности
        for i, val in enumerate(solution):
            result_string += f"  x{i+1} = {val:.2f}\n"
        
        # Вывод в консоль
        print(result_string)
        
        # Вывод в текстовое поле в окне
        output_text.insert(tk.END, result_string)
        
        # Сохранение в файлы
        save_to_txt(result_string)
        save_to_excel(method, A_matrix, B_vector, solution)
        output_text.insert(tk.END, "\n\nРезультаты сохранены в файлы 'results.txt' и 'results.xlsx'")
    else:
        output_text.insert(tk.END, "Ошибка: метод не выбран или не найден.")


def save_to_txt(content):
    """Сохраняет текстовое содержимое в файл results.txt."""
    with open("results.txt", "w", encoding="utf-8") as f:
        f.write(content)

def save_to_excel(method, A, B, X):
    """Сохраняет исходные данные и результат в Excel-файл."""
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Результаты решения СЛУ"

    sheet.append(["Метод решения", method])
    sheet.append([]) # Пустая строка для отступа

    sheet.append(["Матрица A"])
    for row in A.tolist():
        sheet.append(row)
    sheet.append([])

    sheet.append(["Вектор B"])
    for val in B.tolist():
        sheet.append([val])
    sheet.append([])

    sheet.append(["Вектор решения X"])
    for i, val in enumerate(X.tolist()):
        sheet.append([f"x{i+1}", val])

    # Автоподбор ширины колонок для красоты
    for col in sheet.columns:
        max_length = 0
        column = col[0].column_letter # Получаем букву колонки
        for cell in col:
            if len(str(cell.value)) > max_length:
                max_length = len(str(cell.value))
        adjusted_width = (max_length + 2)
        sheet.column_dimensions[column].width = adjusted_width

    workbook.save("results.xlsx")


# --- СОЗДАНИЕ ГРАФИЧЕСКОГО ИНТЕРФЕЙСА ---

# Главное окно
root = tk.Tk()
root.title("Решение СЛУ")
root.geometry("650x700")

# Заголовок
title_label = tk.Label(root, text="Решение системы линейных уравнений A*X = B", font=("Arial", 16, "bold"))
title_label.pack(pady=10)

# Фрейм для отображения постановки задачи
task_frame = tk.LabelFrame(root, text="Постановка задачи", padx=10, pady=10, font=("Arial", 12))
task_frame.pack(pady=10, padx=10, fill="x")

# Формируем красивое отображение системы
system_str = "Матрица A:\n" + np.array2string(A_matrix, precision=2, separator=', ')
system_str += "\n\nВектор B:\n" + np.array2string(B_vector, precision=2, separator=', ')
task_label = tk.Label(task_frame, text=system_str, font=("Courier", 11), justify=tk.LEFT)
task_label.pack()

# Фрейм для управления
control_frame = tk.Frame(root)
control_frame.pack(pady=10)

# Выбор метода
method_label = tk.Label(control_frame, text="Выберите метод:", font=("Arial", 12))
method_label.pack(side=tk.LEFT, padx=5)

method_selector = ttk.Combobox(control_frame, values=["Метод Гаусса", "Метод Гаусса-Жордана"], font=("Arial", 12))
method_selector.current(0) # Установить значение по умолчанию
method_selector.pack(side=tk.LEFT, padx=5)

# Кнопка для решения
solve_button = tk.Button(control_frame, text="Решить", font=("Arial", 12, "bold"), command=solve_system)
solve_button.pack(side=tk.LEFT, padx=10)

# Окно для вывода результатов
output_frame = tk.LabelFrame(root, text="Результаты", padx=10, pady=10, font=("Arial", 12))
output_frame.pack(pady=10, padx=10, fill="both", expand=True)

output_text = scrolledtext.ScrolledText(output_frame, width=70, height=20, font=("Courier", 10))
output_text.pack(fill="both", expand=True)

# Запуск главного цикла приложения
root.mainloop()
