import tkinter as tk
from tkinter import ttk
from tkinter.messagebox import showinfo
from tkinter import Frame
from tkinter import messagebox as mb
from tkinter.filedialog import askopenfilename, asksaveasfilename

from parsXLSX import povogonka
from sentEmail import mail_sandler


class Window(tk.Tk):
    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)
        # Adding a title to the window
        self.wm_title("Incolab - Wagons")
        self.minsize(370, 790)
        self.maxsize(370, 790)
        # creating a frame and assigning it to container
        container = tk.Frame(self)
        # specifying the region where the frame is packed in root
        container.pack(side="left", anchor='n', expand=True)

        # configuring the location of the container using grid
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        # We will now create a dictionary of frames
        self.frames = {}
        # we'll create the frames themselves later but let's add the components to the dictionary.
        for F in (MainPage, EmailPage):
            frame = F(container, self)
            # the windows class acts as the root window for the frames.
            self.frames[F] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        # Using a method to switch frames
        self.show_frame(MainPage)

    def show_frame(self, cont):
        frame = self.frames[cont]
        # raises the current frame to the top
        frame.tkraise()

# TODO reformat the classes


class MainPage(tk.Frame):
    def get_main_file_path(self):
        self.main_file_path = askopenfilename(filetypes=[("Text Files", "*.xlsx")])

    def save_new_file(self):
        self.new_file_path = asksaveasfilename(defaultextension=".xlsx", filetypes=[("Text Files", "xlsx")])

    def start(self):
        try:
            text = self.txt.get("1.0", tk.END)
            self.tuple_wagons = tuple(int(elem) for elem in text.splitlines())
            povogonka(self.main_file_path, self.new_file_path, self.tuple_wagons)
            self.txt.delete("1.0", tk.END)  # Maby it would destroy
            self.txt.insert(tk.END, "Success!")
            if self.email.get():
                self.controller.show_frame(EmailPage)
        except TypeError:
            mb.showerror("Ошибка", "Не указан путь к исходному файлу")
        except ValueError:
            mb.showerror("Ошибка", "Номера вагонов отсутствуют или указаны с ошибкой")
        except Exception as e:
            mb.showerror("Ошибка", f"{e}")

    def __init__(self, parent, controller):
        self.tuple_wagons = None
        self.main_file_path = None
        self.new_file_path = None
        self.txt = None
        self.email = tk.BooleanVar()
        self.controller = controller

        tk.Frame.__init__(self, parent)
        self.pack(anchor="n", side=tk.LEFT)
        buttons_dict = {"Путь к исходному файлу": self.get_main_file_path,
                        "Путь к новому файлу": self.save_new_file,
                        "Старт": lambda: self.start(),
                        "Выход": lambda: self.quit(),
                        }

        def texts(root):
            row = tk.Frame(root)
            row.pack(side="top", padx=5, pady=5)
            lbl = tk.Label(row, text="Insert list of wagons")
            ent = tk.Text(row, width=30)
            lbl.pack()
            ent.pack(side="left")
            return ent

        def checkboxes(root):
            row = tk.Frame(root)
            checkbox = tk.Checkbutton(
                row,
                text="Put if you want to send a file",
                variable=root.email,
                onvalue=True,
                offvalue=False,
                width=30,
            )
            checkbox.pack(side="top", expand=True)
            row.pack(side="top", fill=tk.X, padx=5, pady=5)

        def buttons(root, bns):
            row = tk.Frame(root)
            for name, cmd in bns.items():
                b = tk.Button(row, text=name, command=cmd, width=30,)
                b.pack(side="top", expand=True, padx=2, pady=2)
            row.pack(side="top", padx=5, pady=5)

        """ Options elements """
        self.txt = texts(self)
        checkboxes(self)
        buttons(self, buttons_dict)


class EmailPage(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        fields = 'Адрес получателя: ', 'Адрес отправителя: ', 'Пароль: ', 'Тема письма: '

        def send(entries):
            mail_sandler()

        def makeform(root, fields):
            entries = []
            for field in fields:
                row = tk.Frame(root)
                lab = tk.Label(row, width=15, text=field, anchor='w')
                ent = tk.Entry(row)
                row.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
                lab.pack(side=tk.LEFT)
                ent.pack(side=tk.RIGHT, expand=tk.YES, fill=tk.X)
                entries.append((field, ent))
            return entries

        def message_body(root):
            row = tk.Frame(root)
            ent = tk.Text(row)
            row.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
            ent.pack(side=tk.RIGHT, expand=tk.YES, fill=tk.X)

        evnts = makeform(self, fields)
        message_body(self)
        self.bind('<Return>', (lambda event, e=evnts: send(e)))
        b1 = tk.Button(self, text='Send', command=(lambda e=evnts: send(e)))
        b1.pack(side=tk.LEFT, padx=5, pady=5)
        b2 = tk.Button(self, text='Quit', command=self.quit)
        b2.pack(side=tk.LEFT, padx=5, pady=5)

# class CompletionScreen(tk.Frame):
#     def __init__(self, parent, controller):
#         tk.Frame.__init__(self, parent)
#         label = tk.Label(self, text="Completion Screen, we did it!")
#         label.pack(padx=10, pady=10)
#         switch_window_button = ttk.Button(
#             self, text="Return to menu", command=lambda: controller.show_frame(MainPage)
#         )
#         switch_window_button.pack(side="bottom", fill=tk.X)

# class Window(tk.Tk):
#     def __init__(self):
#         super().__init__()
#
#         # configure the root window
#         self.title('My Awesome App')
#         self.geometry("300x300")
#         self.main_frame = tk.Frame(self, relief=tk.FLAT, border=4)
#         self.main_frame.pack(anchor="n", side=tk.LEFT)
#
#         self.text_edit(self.main_frame)
#         self.path_main_button(self.main_frame)
#         # # label
#         # self.label = ttk.Label(self, text='Hello, Tkinter!')
#         # self.label.pack()
#
#
#         # # button
#         # self.button = ttk.Button(self, text='Click Me')
#         # self.button['command'] = self.button_clicked
#         # self.button.pack()
#
#     def text_edit(self, master):
#         text = tk.Text(master, bd=4, relief=tk.SUNKEN)
#         text.pack(side=tk.TOP, fill=tk.X)
#
#     def open_file(self):
#         return askopenfilename(filetypes=[("Text Files", "*.xlsx")])
#
#     def path_main_button(self, master):
#         button_main = tk.Button(master, text="Путь к исходному файлу", command=self.open_file)
#         button_main.pack(pady=2, fill=tk.X)


# class Window(tk.Tk):
#     def __init__(self):
#         super().__init__()
#         self.main_path = None
#         self.new_path = None
#         self.tuple_wagons = None
#         """ Main window """
#         self.title("Wagons")
#         self.maxsize(370, 720)
#         self.minsize(370, 720)
#         """ Text form """
#         self.window = self
#         """ Button form """
#         # frm = tk.Frame(self.window, relief=tk.FLAT, bd=4, padx=5, pady=5)
#         # frm.pack(anchor="n", side=tk.LEFT)
#         # self.txt_edit = tk.Text(frm, bd=4, relief=tk.SUNKEN)
#         # main_file_path = tk.Button(frm, text="Путь к исходному файлу", command=self.open_file)
#         # new_file_path = tk.Button(frm, text="Путь к новому файлу", command=self.save_file)
#         # btn_start = tk.Button(frm, text="Старт", command=self.start)
#         # btn_close = tk.Button(frm, text="Выход", command=self.close_window)
#         #
#         # """ Options elements """
#         # self.txt_edit.pack(side=tk.TOP, fill=tk.X)
#         # main_file_path.pack(pady=2, fill=tk.X)
#         # new_file_path.pack(pady=2, fill=tk.X)
#         # btn_start.pack(pady=2, fill=tk.X)
#         # btn_close.pack(pady=2, fill=tk.X)
#
#     def base(self, window):
#         frm = tk.Frame(window, relief=tk.FLAT, bd=4, padx=5, pady=5)
#         frm.pack(anchor="n", side=tk.LEFT)
#         txt_edit = tk.Text(frm, bd=4, relief=tk.SUNKEN)
#         main_file_path = tk.Button(frm, text="Путь к исходному файлу", command=self.open_file)
#         new_file_path = tk.Button(frm, text="Путь к новому файлу", command=self.save_file)
#         btn_start = tk.Button(frm, text="Старт", command=self.start)
#         btn_close = tk.Button(frm, text="Выход", command=self.close_window)
#
#         """ Options elements """
#         txt_edit.pack(side=tk.TOP, fill=tk.X)
#         main_file_path.pack(pady=2, fill=tk.X)
#         new_file_path.pack(pady=2, fill=tk.X)
#         btn_start.pack(pady=2, fill=tk.X)
#         btn_close.pack(pady=2, fill=tk.X)
#
#     def open_file(self):
#         path_to_file = askopenfilename(filetypes=[("Text Files", "*.xlsx")])
#         self.main_path = path_to_file
#
#     def save_file(self):
#         path_to_file = asksaveasfilename(defaultextension=".xlsx", filetypes=[("Text Files", "xlsx")])
#         if not path_to_file:
#             return "./Wagons.xlsx"
#         self.new_path = path_to_file
#
#     def start(self):
#
#         try:
#             self.base()
#             text = self.txt_edit.get("1.0", tk.END)
#             self.tuple_wagons = tuple(int(elem) for elem in text.splitlines())
#             povogonka(self.main_path, self.new_path, self.tuple_wagons)
#             self.txt_edit.delete("1.0", tk.END)
#             self.txt_edit.insert(tk.END, "Success!")
#         except TypeError:
#             mb.showerror("Ошибка", "Не указан путь к исходному файлу")
#         except ValueError:
#             mb.showerror("Ошибка", "Номера вагонов отсутствуют или указаны с ошибкой")
#         except Exception as e:
#             mb.showerror("Ошибка", f"{e}")
#
#
#     def close_window(self):
#         self.destroy()
