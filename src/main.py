import tkinter as tk
from tkinter import ttk, filedialog

import sv_ttk
from timetable import *

class Button(ttk.Frame):
    def __init__(self, parent, height=None, width=None, text="", command=None, style='TButton', bg=['#eeeeee', '#cccccc'], fg='#000000'):
        ttk.Frame.__init__(self, parent, height=height, width=width, style=style)

        self.pack_propagate(False)
        # self._btn = ttk.Button(self, text=text, command=command, style=style)
        self._btn = tk.Button(
            self,
            bg=bg[0],
            fg=fg,
            relief='flat',
            text=text,
            command=command,
            font=("Pretendard Variable", 11),
            borderwidth=0,
            activebackground=bg[1],
            activeforeground=fg
        )
        self._btn.pack(fill=tk.BOTH, expand=1)

class App(ttk.Frame):
    def __init__(self, parent):
        super().__init__()

        # set size of window
        self.parent = parent
        self.parent.title("시간표")

        self.period_width = 35
        self.default_width = 100
        self.week_height = 40
        self.default_height = 60

        self.parent.geometry("250x150")

        self.students = None
        self.student = None

        self.tabs = []

        self.load_screen()

    def new_tab(self, entry: tk.StringVar, warn: ttk.Label):
        name = entry.get().strip()

        if self.students is None:
            warn.configure(text='학생 목록을 불러와주세요')
            return

        student = self.students[name]
        if self.students is not None and student is not None:
            warn.configure(text='')

            self.student = student

            new = tk.Toplevel(self.parent)
            new.title(f"{self.student}의 시간표")
            new.geometry(
                f"{self.period_width + len(Week) * self.default_width}x{self.week_height + 8 * self.default_height}")
            new.resizable(False, False)

            self.timetable_screen(new, self.student)
            self.tabs.append(new)
        else:
            warn.configure(text='학생을 찾을 수 없습니다')

    def load_students(self, button):
        f = filedialog.askopenfilename(filetypes=[("timetable xlsx file", '*.xlsx')])
        if not f:
            return

        self.students = StudentList.from_xlsx(f)
        # button.configure(style='Accent.TButton', state='disabled')
        button.destroy()

        warn = ttk.Label(self, font=("Pretendard Variable", 9))
        warn.place(x=125, y=125, anchor='center')

        name_var = tk.StringVar(value='학번이나 이름을 입력하세요')
        name_entry = ttk.Entry(self, textvariable=name_var)
        name_entry.bind('<Return>', lambda _: self.new_tab(name_var, warn))
        name_entry.bind('<Button-1>', lambda _: name_var.set(''))
        name_entry.place(x=125, y=100, anchor='center')

    def load_screen(self):
        logo = ttk.Label(self, text='Timetable', font=("Product Sans Bold", 25))
        logo.place(x=125, y=50, anchor='center')

        button = ttk.Button(self, text='Load File', command=lambda: self.load_students(button), style='Accent.TButton')
        button.place(x=125, y=100, anchor='center')

    def timetable_screen(self, new: tk.Toplevel, student: Student):
        block = Button(new, width=self.period_width, height=self.week_height)
        block.place(x=0, y=0)

        for i in range(1, 9):
            period_block = Button(new, text=f'{i}', width=self.period_width, height=self.default_height)
            period_block.place(x=0, y=self.week_height + (i-1)*self.default_height)

        for week in Week:
            week_block = Button(new, text=week.name, width=self.default_width, height=self.week_height)
            week_block.place(x=self.period_width + week.value*self.default_width, y=0)

            E = []
            for _, period in student.timetable[week]:
                subject = student.timetable[week][period].value()
                E.append((period, subject))

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

            for start, cnt, subject in ret:
                subject_name = subject.name.replace(' ', '')
                label_txt = '\n'.join(subject_name[i:i + 5].strip() for i in range(0, len(subject_name), 5))

                subject_block = Button(new, text=label_txt, width=self.default_width, height=self.default_height * cnt, bg=subject.type.color)
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
