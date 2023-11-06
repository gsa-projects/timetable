import tkinter as tk
from tkinter import ttk, filedialog
from preset import *

class App(ttk.Frame):
    def __init__(self, parent):
        super().__init__()

        # set size of window
        self.parent = parent
        self.parent.title("Timetable")

        # set window col=7, row=10
        for i in range(7):
            self.parent.columnconfigure(i, weight=1)
        for i in range(11):
            self.parent.rowconfigure(i, weight=1)

        self.students = None
        self.student = None

        self.load_screen()

    def load_screen(self):
        def load_students():
            f = filedialog.askopenfilename(filetypes=[("timetable xlsx file", '*.xlsx')])
            if f:
                self.students = StudentList.from_xlsx(f)
                is_loaded.config(text='loaded')

        button = ttk.Button(self, text='load timetable', command=load_students)
        button.pack()

        is_loaded = ttk.Label(self, text='not loaded')
        is_loaded.pack()

        name_var = tk.StringVar()
        name_entry = ttk.Entry(self, textvariable=name_var)
        name_entry.pack()

        def go_to_timetable():
            name = name_var.get().strip()

            if self.students is None:
                print('denied')
                return

            student = self.students[name]
            if self.students is not None and student is not None:
                self.student = student
                self.timetable_screen(self.student)
            else:
                print('denied')

        go_to_button = ttk.Button(self, text='go to', command=go_to_timetable)
        go_to_button.pack()

    def timetable_screen(self, student: Student):
        # clear all widgets
        for widget in self.winfo_children():
            widget.destroy()

        days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
        for i, day in enumerate(days):
            label = ttk.Button(self, text=day, width=10, style="Accent.TButton")
            label.grid(row=0, column=i)

        for (week, period), subject in student.timetable.items():
            # Button width 20, 너무 길면 아래로 늘어나기
            # label = ttk.Button(self, text=subject, width=10)
            # label.grid(row=week.value, column=period)
            label = ttk.Button(self, text=subject, width=10)
            label.grid(row=week.value+1, column=period)


if __name__ == "__main__":
    root = tk.Tk()
    root.title("")

    # Simply set the theme
    root.tk.call("source", "../azure.tcl")
    # root.tk.call("set_theme", "dark")
    root.tk.call("set_theme", "dark")

    app = App(root)
    app.pack(fill="both", expand=True)

    # Set a minsize for the window, and place it in the middle
    root.update()
    root.minsize(root.winfo_width(), root.winfo_height())
    x_coordinate = int((root.winfo_screenwidth() / 2) - (root.winfo_width() / 2))
    y_coordinate = int((root.winfo_screenheight() / 2) - (root.winfo_height() / 2))
    root.geometry("+{}+{}".format(x_coordinate, y_coordinate - 20))

    root.mainloop()
