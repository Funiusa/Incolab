import errno
import os.path
import tkinter as tk
import traceback
from tkinter import messagebox as mb
from tkinter.filedialog import askopenfilename, asksaveasfilename

from sendEmail import mail_sendler


class Window(tk.Tk):

    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)
        # Adding a title to the window
        self.wm_title("Incolab - Wagons")
        self.minsize(370, 780)
        self.maxsize(370, 780)
        # creating a frame and assigning it to container
        container = tk.Frame(self)
        # specifying the region where the frame is packed in root
        container.pack(side="left", anchor='n')
        # configuring the location of the container using grid
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        # We will now create a dictionary of frames
        self.frames = {}
        # we'll create the frames themselves later but let's add the components to the dictionary.
        for F in (MainPage, EmailPage):
            frame = F(container, self)  # the windows class acts as the root window for the frames.
            self.frames[F] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        # Using a method to switch frames
        self.show_frame(MainPage)
        # Variable for using in ather frames
        self.new_file_path = None

    def show_frame(self, cont):
        frame = self.frames[cont]
        # raises the current frame to the top
        frame.tkraise()


# TODO reformat the classes


class MainPage(tk.Frame):
    def get_main_file_path(self):
        self.main_file_path = askopenfilename(filetypes=[("Text Files", "*.xlsx")])

    def save_new_file(self):
        self.controller.new_file_path = asksaveasfilename(defaultextension=".xlsx", filetypes=[("Text Files", "xlsx")])

    def start(self):
        try:
            # text = self.txt_form.get("1.0", tk.END)
            # self.tuple_wagons = tuple(int(elem) for elem in text.splitlines())
            # povogonka(self.main_file_path, self.new_file_path, self.tuple_wagons)
            # self.txt_form.delete("1.0", tk.END)  # Maby it would destroy
            # self.txt_form.insert(tk.END, "Success!")
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
        # self.new_file_path = None
        self.txt_form = None
        self.email = tk.BooleanVar()
        self.controller = controller

        tk.Frame.__init__(self, parent)
        self.pack(anchor="n", side="left")
        buttons_dict = {"Путь к исходному файлу": self.get_main_file_path,
                        "Путь к новому файлу": self.save_new_file,
                        "Старт": lambda: self.start(),
                        "Выход": lambda: self.quit(),
                        }

        def texts(root):
            row = tk.Frame(root)
            row.pack(side="top", padx=5, pady=10)
            txt = tk.Text(row, width=30)
            txt.insert('1.0', "Insert list of wagons here")
            txt.pack(side="left")
            return txt

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
                b = tk.Button(row, text=name, command=cmd, width=30, )
                b.pack(side="top", expand=True, padx=2, pady=2)
            row.pack(side="top", padx=5, pady=5)

        """ Options elements """
        texts(self)
        self.new_file = None
        # self.txt_form.bind("<Button-1>", lambda event: self.txt_form.delete('1.0', tk.END))
        checkboxes(self)
        buttons(self, buttons_dict)


class EmailPage(tk.Frame):
    def send_email(self, entrs, body):
        values = [entr.get() for entr in entrs]
        values.append(body.get("1.0", tk.END))
        values.append(self.controller.new_file_path)  # Append file path
        try:
            #               To, From who,  Password, Theme, Body message, file_path
            mail_sendler(values[0], values[1], values[2], values[3], values[4], values[5])
            mb.showinfo("", "Сообщение успешно отправлено")
        except TypeError:
            mb.showerror("Error", "Не указан файл для отправки")
        except Exception as e:
            mb.showerror("ERROR", f"{e}")
            # if errno.EIO == 5:
            #     mb.showerror("Error",
            #                  f"Ошибка: файл с именем {os.path.basename(self.controller.new_file_path)} не существует")
            # self.controller.show_frame(MainPage)
            self.controller.new_file_path = askopenfilename(filetypes=[("Text Files", "*.xlsx")])
        except:
            mb.showerror("ERROR",
                         f"""Логин или пароль указаны не верно. Или не настроена почта для отправки сообщений.""")

    def __init__(self, parent, controller):
        self.controller = controller
        tk.Frame.__init__(self, parent)
        fields = 'Кому', 'От кого', 'Пароль', 'Тема письма'

        def makeform(root, flds):
            entrys = []
            for field in flds:
                row = tk.Frame(root)
                row.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
                if field == 'Тема письма':
                    ent = tk.Entry(row)
                    ent.insert('0', field)
                    ent.bind('<Button-1>', lambda event: ent.delete('0', tk.END))
                    ent.pack(fill=tk.X, padx=4)
                else:
                    lab = tk.Label(row, padx=4, text=field, anchor='w')
                    ent = tk.Entry(row, width=30)
                    lab.pack(side=tk.LEFT)
                    ent.pack(side=tk.RIGHT, fill=tk.X, padx=4)
                entrys.append(ent)
            return entrys

        def message_body(root):
            row = tk.Frame(root)
            scroll_bar = tk.Scrollbar(row)
            scroll_bar.pack(side=tk.RIGHT, fill=tk.Y)
            text = tk.Text(row, yscrollcommand=scroll_bar.set)
            text.insert('1.0', "Insert your message here")
            row.pack(side=tk.TOP, fill=tk.X, padx=8, pady=4)
            text.pack(side=tk.RIGHT, expand=tk.YES, fill=tk.X)
            return text

        entries = makeform(self, fields)
        txt = message_body(self)
        txt.bind('<Button-1>', lambda event: txt.delete("1.0", tk.END))
        send_btn = tk.Button(self, text='Send', command=lambda: self.send_email(entries, txt))
        send_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=5)
        quit_bnt = tk.Button(self, text='Quit', command=self.quit)
        quit_bnt.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=5, pady=5)
