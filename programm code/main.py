import os
import pandas as pd
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
def анализировать_файл(input_file, output_file, min_repeats, selected_statuses, group_column,
                        column_text, column_object):
    df = pd.read_excel(input_file)
    if not selected_statuses:
        raise ValueError("Не выбран ни один статус")
    df = df[df["Статус замечания"].isin(selected_statuses)].copy().reset_index(drop=True)
    if group_column:
        group_values = df[group_column].dropna().unique()
    else:
        group_values = [None]
    #model = SentenceTransformer("paraphrase-MiniLM-L6-v2")
    script_dir = os.path.dirname(os.path.abspath(__file__))
    model_path = os.path.join(script_dir, "local_model")
    model = SentenceTransformer(model_path)
    df[column_text] = df[column_text].fillna("").astype(str)
    df[column_object] = df[column_object].fillna("").astype(str)
    embeddings = model.encode(df[column_text].tolist(), show_progress_bar=True)
    similarity_matrix = cosine_similarity(embeddings)
    группа_по_индексу = {}
    group_id_counter = 1
    global_visited = set()
    def найти_группы(df_subset, indices):
        nonlocal group_id_counter
        visited = set()
        for i in range(len(indices)):
            idx_i = indices[i]
            if idx_i in visited:
                continue
            temp_group = []
            obj_i = df_subset.loc[idx_i, column_object].strip().lower()
            for j in range(i + 1, len(indices)):
                idx_j = indices[j]
                if idx_j in visited:
                    continue
                obj_j = df_subset.loc[idx_j, column_object].strip().lower()
                sim = similarity_matrix[idx_i][idx_j]
                if sim >= 0.99 and obj_i != obj_j:
                    temp_group.append(idx_j)
            if temp_group:
                temp_group.append(idx_i)
                main_text = df_subset.loc[idx_i, column_text]
                full_group = [k for k in df_subset.index if df_subset.loc[k, column_text] == main_text]
                group = list(set(temp_group + full_group))
                if len(group) >= min_repeats:
                    visited.update(group)
                    global_visited.update(group)
                    for idx in group:
                        группа_по_индексу[idx] = f"Группа {group_id_counter}"
                    group_id_counter += 1
    for value in group_values:
        if value:
            df_group = df[df[group_column] == value]
        else:
            df_group = df
        indices = df_group.index.tolist()
        найти_группы(df_group, indices)
    лист = []
    for value in group_values:
        if value:
            df_group = df[(df[group_column] == value) & (df.index.isin(global_visited))]
            if df_group.empty:
                continue
            лист.append([value, "", ""])
        else:
            df_group = df[df.index.isin(global_visited)]
        grouped = df_group.groupby(группа_по_индексу.get)
        for group_name, group_df in grouped:
            if not group_name:
                continue
            example_text = group_df.iloc[0][column_text]
            заголовок = ["", group_name, example_text]
            while len(заголовок) < len(df.columns) + 2:
                заголовок.append("")
            лист.append(заголовок)
            for _, row in group_df.iterrows():
                первое_значение = str(int(row[0])) if isinstance(row[0], float) and row[0].is_integer() else str(row[0])
                лист.append(["", f"{group_name}.{первое_значение}", *row.tolist()])
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df_sheet = pd.DataFrame(лист)
        заголовки = df.columns.tolist()
        пустая_строка = ["", ""] + заголовки
        df_sheet = pd.concat([pd.DataFrame([пустая_строка]), df_sheet], ignore_index=True)
        df_sheet.to_excel(writer, sheet_name="Типовые", index=False, header=False)
def выбрать_файл():
    путь = filedialog.askopenfilename(filetypes=[("Excel файлы", "*.xlsx")])
    if путь:
        entry_input.delete(0, tk.END)
        entry_input.insert(0, путь)
        обновить_список_столбцов()
def выбрать_сохранение():
    путь = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel файлы", "*.xlsx")])
    if путь:
        entry_output.delete(0, tk.END)
        entry_output.insert(0, путь)
def обновить_список_столбцов():
    путь = entry_input.get()
    if not путь:
        return
    try:
        df = pd.read_excel(путь, nrows=1)
        cols = list(df.columns)
        combo_text['values'] = cols
        combo_object['values'] = cols
        combo_group['values'] = [""] + cols
        if "Содержание замечания" in cols:
            combo_text.set("Содержание замечания")
        if "Наименование объекта" in cols:
            combo_object.set("Наименование объекта")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось прочитать файл:\n{e}")
def запуск():
    try:
        input_path = entry_input.get()
        output_path = entry_output.get()
        min_repeats = int(entry_min.get())
        selected_statuses = []
        if var_p.get():
            selected_statuses.append("П")
        if var_r.get():
            selected_statuses.append("Р")
        group_column = combo_group.get().strip()
        if group_column == "":
            group_column = None
        column_text = combo_text.get().strip()
        column_object = combo_object.get().strip()
        if not column_text or not column_object:
            messagebox.showerror("Ошибка", "Выберите обязательные столбцы!")
            return
        анализировать_файл(
            input_path, output_path, min_repeats, selected_statuses,
            group_column, column_text, column_object
        )
        messagebox.showinfo("Успех", "Готово!")
    except Exception as e:
        messagebox.showerror("Ошибка", str(e))
root = tk.Tk()
root.title("Поиск типовых замечаний")
root.geometry("700x400")
tk.Label(root, text="Выбор исходного файла:").pack()
entry_input = tk.Entry(root, width=80)
entry_input.pack()
tk.Button(root, text="Открыть", command=выбрать_файл).pack(pady=3)
tk.Label(root, text="Сохранить файл как:").pack()
entry_output = tk.Entry(root, width=80)
entry_output.pack()
tk.Button(root, text="Сохранить как", command=выбрать_сохранение).pack(pady=3)
tk.Label(root, text="Минимум повторений:").pack()
entry_min = tk.Entry(root, width=10)
entry_min.insert(0, "3")
entry_min.pack()
frame_status = tk.Frame(root)
tk.Label(frame_status, text="Статус анализируемого замечания:").pack(side=tk.LEFT)
var_p = tk.BooleanVar(value=True)
var_r = tk.BooleanVar(value=True)
tk.Checkbutton(frame_status, text="Принятые (П)", variable=var_p).pack(side=tk.LEFT)
tk.Checkbutton(frame_status, text="Рассмотренные (Р)", variable=var_r).pack(side=tk.LEFT)
frame_status.pack(pady=5)
tk.Label(root, text="Столбец с текстом:").pack()
combo_text = ttk.Combobox(root, width=70)
combo_text.pack()
tk.Label(root, text="Столбец с наименованием:").pack()
combo_object = ttk.Combobox(root, width=70)
combo_object.pack()
tk.Label(root, text="Группировка по столбцу:").pack()
combo_group = ttk.Combobox(root, width=70)
combo_group.pack()
tk.Button(root, text="Запустить", command=запуск, bg="green", fg="white").pack(pady=15)
root.mainloop()
