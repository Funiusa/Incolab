import tkinter as tk
from tkinter.filedialog import askopenfilename, asksaveasfilename
from tkinter import messagebox as mb
from parsXLSX import povogonka


class Window(tk.Tk):
    def __init__(self):
        super().__init__()
        self.main_path = None
        self.new_path = None
        self.tuple_wagons = None

        self.title("Wagons")
        self.maxsize(370, 610)
        self.minsize(370, 610)
        self.txt_edit = tk.Text(self, bd=4, width=20, height=20, relief=tk.SUNKEN)
        frm_buttons = tk.Frame(self, relief=tk.FLAT, bd=4, height=10)
        main_file_path = tk.Button(text="Путь к исходному файлу", command=self.open_file)
        new_file_path = tk.Button(text="Путь к новому файлу", command=self.save_file)
        btn_start = tk.Button(text="Старт", command=self.start)
        btn_close = tk.Button(text="Выход", command=self.close_window)

        self.txt_edit.pack(side=tk.TOP, fill=tk.X)
        frm_buttons.pack(side=tk.BOTTOM)
        main_file_path.pack(anchor="n", padx=4, pady=1, fill=tk.X)
        new_file_path.pack(anchor="n", padx=4, pady=1, fill=tk.X)
        btn_start.pack(anchor="n", padx=4, pady=1, fill=tk.X)
        btn_close.pack(anchor="n", padx=4, pady=1, fill=tk.X)

    def open_file(self):
        path_to_file = askopenfilename(filetypes=[("Text Files", "*.xlsx")])
        self.main_path = path_to_file

    def save_file(self):
        path_to_file = asksaveasfilename(defaultextension=".xlsx", filetypes=[("Text Files", "xlsx")])
        if not path_to_file:
            return "./Wagons.xlsx"
        self.new_path = path_to_file

    def start(self):

        try:

            text = self.txt_edit.get("1.0", tk.END)
            self.tuple_wagons = tuple(int(elem) for elem in text.splitlines())
            povogonka(self.main_path, self.new_path, self.tuple_wagons)
            self.txt_edit.delete("1.0", tk.END)
            self.txt_edit.insert(tk.END, "Success!")
        except TypeError:
            mb.showerror("Ошибка", "Не указан путь к исходному файлу")
        except ValueError:
            mb.showerror("Ошибка", "Номера вагонов отсутствуют или указаны с ошибкой")
        except Exception as e:
            mb.showerror("Ошибка", f"{e}")

    def close_window(self):
        self.destroy()
