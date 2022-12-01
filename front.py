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

        self.title("Incolab - Wagons")
        self.geometry("400x300")
        self.txt_edit = tk.Text(self, bd=4, width=20, height=30)
        frm_buttons = tk.Frame(self, relief=tk.FLAT, bd=4)
        main_file_path = tk.Button(frm_buttons, text="Путь к основному файлу", command=self.open_file)
        new_file_path = tk.Button(frm_buttons, text="Путь к новому файлу", command=self.save_file)
        btn_start = tk.Button(frm_buttons, text="Старт", command=self.start)
        btn_close = tk.Button(frm_buttons, text="Выход", command=self.close_window)

        main_file_path.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        new_file_path.grid(row=1, column=0, sticky="ew", padx=5, pady=5)
        btn_start.grid(row=2, column=0, sticky="ew", padx=5, pady=5)
        btn_close.grid(row=3, column=0, sticky="ew", padx=5, pady=5)
        frm_buttons.grid(row=0, column=0, sticky="ns")
        self.txt_edit.grid(row=0, column=1, sticky="nsew")

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
        except ValueError:
            mb.showerror("Ошибка", "Номером вагона может быть только числом")
        except Exception as e:
            mb.showerror("Ошибка", f"{e}")

    def close_window(self):
        self.destroy()


if __name__ == "__main__":
    root = Window()
    root.mainloop()



