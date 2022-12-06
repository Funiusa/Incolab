import tkinter as tk
from tkinter import ttk
from tkinter.messagebox import showinfo
from tkinter import Frame
from tkinter import messagebox as mb
from tkinter.filedialog import askopenfilename, asksaveasfilename

from parsXLSX import povogonka


class Window(tk.Tk):
    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)
        # Adding a title to the window
        self.wm_title("Incolab - Wagons")
        self.minsize(370, 790)
        # creating a frame and assigning it to container
        container = tk.Frame(self, height=370, width=720)
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


class MainPage(tk.Frame):
    def get_main_file_path(self):
        self.main_file_path = askopenfilename(filetypes=[("Text Files", "*.xlsx")])

    def save_new_file(self):
        self.new_file_path = asksaveasfilename(defaultextension=".xlsx", filetypes=[("Text Files", "xlsx")])

    def start(self):
        if self.email.get():
            self.controller.show_frame(EmailPage)
        # try:
        #     text = self.txt_edit.get("1.0", tk.END)
        #     self.tuple_wagons = tuple(int(elem) for elem in text.splitlines())
        #     povogonka(self.main_file_path, self.new_file_path, self.tuple_wagons)
        #     self.txt_edit.delete("1.0", tk.END)
        #     self.txt_edit.insert(tk.END, "Success!")
        #     if self.email.get():
        #         self.controller.show_frame(EmailPage)
        # except TypeError:
        #     mb.showerror("Ошибка", "Не указан путь к исходному файлу")
        # except ValueError:
        #     mb.showerror("Ошибка", "Номера вагонов отсутствуют или указаны с ошибкой")
        # except Exception as e:
        #     mb.showerror("Ошибка", f"{e}")


    def __init__(self, parent, controller):

        self.tuple_wagons = None
        self.main_file_path = None
        self.new_file_path = None
        self.email = tk.BooleanVar()
        self.controller = controller

        tk.Frame.__init__(self, parent)
        label = tk.Label(self, text="Addition wagons list")
        label.pack(padx=10, pady=10)
        self.txt_edit = tk.Text(self, bd=4, relief=tk.SUNKEN)
        self.txt_edit.pack(side=tk.TOP, fill=tk.X)

        """ Options elements """
        # We use the switch_window_button in order to call the show_frame() method as a lambda function
        checkbox = tk.Checkbutton(
            self,
            text="Send file on the email",
            variable=self.email,
            onvalue=True,
            offvalue=False,
        )
        checkbox.pack(pady=2, fill=tk.X)
        # Get path for main file
        main_file_path_button = tk.Button(
            self,
            text="Путь к исходному файлу",
            command=self.get_main_file_path,
        )
        main_file_path_button.pack(pady=2, fill=tk.X)
        # Save new file
        new_file_path_button = tk.Button(
            self,
            text="Путь к новому файлу",
            command=self.save_new_file,
        )
        new_file_path_button.pack(pady=2, fill=tk.X)
        # Start button
        exit_button = tk.Button(
            self,
            text="Старт",
            command=lambda: self.start(),
        )
        exit_button.pack(pady=2, fill=tk.X)
        # Exit button
        exit_button = tk.Button(
            self,
            text="Exit",
            command=lambda: controller.destroy(),
        )
        exit_button.pack(pady=2, fill=tk.X)

        self.pack(anchor="n", side=tk.LEFT)


class EmailPage(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        label = tk.Label(self, text="This is page for sending email")
        label.pack(padx=10, pady=10)

        self.txt_edit = tk.Text(self, width=5, height=1, bd=4, relief=tk.SUNKEN)
        self.txt_edit.pack(side=tk.TOP, fill=tk.X)

        self.txt_edit = tk.Text(self, width=10, height=20, bd=4, relief=tk.SUNKEN)
        self.txt_edit.pack(side=tk.TOP, fill=tk.X)
        switch_window_button = tk.Button(
            self,
            text="Exit",
            command=controller.destroy,
        )
        switch_window_button.pack(side="bottom", fill=tk.X)

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
