import re
import sv_ttk
import tkinter as tk
from tkinter import ttk, filedialog
from preset import *

class MyButton(ttk.Frame):
    def __init__(self, parent, height=None, width=None, text="", command=None, style='TButton'):
        ttk.Frame.__init__(self, parent, height=height, width=width, style=style)

        self.pack_propagate(False)
        self._btn = ttk.Button(self, text=text, command=command, style=style)
        self._btn.pack(fill=tk.BOTH, expand=1)

class App(ttk.Frame):
    def __init__(self, parent):
        super().__init__()

        # set size of window
        self.parent = parent
        self.parent.title("Timetable")

        self.period_width = 35
        self.default_width = 100
        self.week_height = 40
        self.default_height = 60

        self.parent.geometry("200x150")

        self.students = None
        self.student = None

        self.tabs = []

        self.load_screen()

    def load_screen(self):
        def load_students():
            f = filedialog.askopenfilename(filetypes=[("timetable xlsx file", '*.xlsx')])
            if f:
                self.students = StudentList.from_xlsx(f)
                # button.configure(style='Accent.TButton', state='disabled')
                button.destroy()

                global name_var, name_entry
                name_var = tk.StringVar()
                name_entry = ttk.Entry(self, textvariable=name_var)
                name_entry.bind('<Return>', lambda _: go_to_timetable())
                name_entry.place(x=100, y=100, anchor='center')


        logo = ttk.Label(self, text='Timetable', font=("Archivo Bold", 20))
        logo.place(x=100, y=50, anchor='center')

        button = ttk.Button(self, text='Load', command=load_students)
        button.place(x=100, y=100, anchor='center')

        def go_to_timetable():
            name = name_var.get().strip()

            if self.students is None:
                print('denied')
                return

            student = self.students[name]
            if self.students is not None and student is not None:
                self.student = student

                new = tk.Toplevel(self.parent)
                new.title(f"Timetable - {self.student}")
                new.geometry(f"{self.period_width + len(Week)*self.default_width}x{self.week_height + 8*self.default_height}")
                new.resizable(False, False)

                self.timetable_screen(new, self.student)
                self.tabs.append(new)
            else:
                print('denied')

    def timetable_screen(self, new: tk.Toplevel, student: Student):
        for i in range(1, 9):
            period_block = MyButton(new, text=f'{i}', width=self.period_width, height=self.default_height)
            period_block.place(x=0, y=self.week_height + (i-1)*self.default_height)

        for week in Week:
            week_block = MyButton(new, text=week.name, width=self.default_width, height=self.week_height, style='Accent.TButton')
            week_block.place(x=self.period_width + week.value*self.default_width, y=0)

            E = []
            for _, period in student.timetable[week]:
                subject = student.timetable[week][period].value()

                subject_name = subject.name.replace(' ', '')
                label_txt = '\n'.join(subject_name[i:i + 5].strip() for i in range(0, len(subject_name), 5))

                E.append((period, label_txt))

            ret = []
            b4 = E[0]
            cnt = 0
            for i, (period, subject) in enumerate(E):
                if subject == b4[1]:
                    cnt += 1
                else:
                    ret.append((b4[0], cnt, b4[1]))
                    b4 = E[i]
                    cnt = 1

                if i == len(E) - 1:
                    ret.append((b4[0], cnt, b4[1]))

            for start, cnt, label_txt in ret:
                subject_block = MyButton(new, text=label_txt, width=self.default_width, height=self.default_height*cnt)
                subject_block.place(x=self.period_width + week.value*self.default_width, y=self.week_height + (start-1)*self.default_height)

if __name__ == "__main__":
    root = tk.Tk()
    root.title("")

    # Simply set the theme
    sv_ttk.set_theme("light")

    app = App(root)
    app.pack(fill="both", expand=True)

    style = ttk.Style()
    style.configure('TButton', font=("Pretendard Variable", 12))

    # Set a minsize for the window, and place it in the middle
    root.update()
    root.minsize(root.winfo_width(), root.winfo_height())
    x_coordinate = int((root.winfo_screenwidth() / 2) - (root.winfo_width() / 2))
    y_coordinate = int((root.winfo_screenheight() / 2) - (root.winfo_height() / 2))
    root.geometry("+{}+{}".format(x_coordinate, y_coordinate - 20))

    root.mainloop()
