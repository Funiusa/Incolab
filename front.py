import tkinter as tk
from tkinter.filedialog import askopenfilename, asksaveasfilename
from parsXLSX import povogonka


def get_main_file():
    """Open a file for editing."""
    filepath = askopenfilename(filetypes=[("Text Files", "*.xlsx")])
    if not filepath:
        return "./Wagons.xlsx"
    return filepath


def save_results():
    """Save the current file as a new file."""
    filepath = asksaveasfilename(defaultextension=".xlsx", filetypes=[("Text Files", "xlsx")])
    if not filepath:
        return
    text = txt_edit.get("1.0", tk.END)
    try:

    #povogonka(tuple(int(elem) for elem in text.splitlines()), )
        print(tuple(int(elem) for elem in text.splitlines()), )
    except ValueError:
        print(f"Номером вагона может быть только число")
    except Exception as e:
        print(e)
    txt_edit.delete("1.0", tk.END)
    window.title(f"Simple Text Editor - {filepath}")


def close_window():
    window.destroy()




if __name__ == "__main__":
    window = tk.Tk()
    window.title("Incolab - Wagons")

    window.rowconfigure(0, minsize=80, weight=1)
    window.columnconfigure(1, minsize=80, weight=1)

    txt_edit = tk.Text(window, bd=4, width=20, height=30)
    frm_buttons = tk.Frame(window, relief=tk.FLAT, bd=4)
    mainfile_path = tk.Button(frm_buttons, text="Путь к основному файлу", command=get_main_file)
    newfile_path = tk.Button(frm_buttons, text="Путь к новому файлу", command=save_results)
    btn_start = tk.Button(frm_buttons, text="Старт", command=save_results)
    btn_close = tk.Button(frm_buttons, text="Выход", command=close_window)

    mainfile_path.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
    newfile_path.grid(row=1, column=0, sticky="ew", padx=5, pady=5)
    btn_start.grid(row=2, column=0, sticky="ew", padx=5, pady=5)
    btn_close.grid(row=3, column=0, sticky="ew", padx=5, pady=5)
    frm_buttons.grid(row=0, column=0, sticky="ns")
    txt_edit.grid(row=0, column=1, sticky="nsew")
    window.mainloop()

