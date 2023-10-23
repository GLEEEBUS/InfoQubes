import pandas as pd
import tkinter as tk
from tkinter import BOTH, LEFT, RIGHT, VERTICAL, Y, ttk, messagebox
import tkinter.filedialog as fd
from tkinter.font import BOLD, Font

import warnings
warnings.filterwarnings("ignore")


class GuiProcess:
    def __init__(self, processes: list):
        self.root = tk.Tk()

        self.processes = processes
        self.vars = []
        self.chosen_processes = []

        self.root.title("Еженедельный отчет по NPS")
        self.font = Font(self.root, size=10, weight=BOLD)

        lbl1 = tk.Label(self.root, text="Выберите процессы", font=self.font)
        lbl1.pack(pady=5)

        btn1 = tk.Button(self.root, text="Выбрать", command=self.pick)
        btn1.pack(pady=1)

        self.update1 = tk.Label(self.root)
        self.update1.pack(pady=5)

        btn2 = tk.Button(self.root, text="Подтвердить", command=self.close)
        btn2.pack(pady=3)

        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=BOTH, expand=1)

        # Create A Canvas
        my_canvas = tk.Canvas(main_frame)
        my_canvas.pack(side=LEFT, fill=BOTH, expand=1)

        # Add A Scrollbar To The Canvas
        my_scrollbar = ttk.Scrollbar(main_frame, orient=VERTICAL, command=my_canvas.yview)
        my_scrollbar.pack(side=RIGHT, fill=Y)

        # Configure The Canvas
        my_canvas.configure(yscrollcommand=my_scrollbar.set)
        my_canvas.bind('<Configure>', lambda e: my_canvas.configure(scrollregion=my_canvas.bbox("all")))

        # Create ANOTHER Frame INSIDE the Canvas
        second_frame = tk.Frame(my_canvas)

        # Add that New frame To a Window In The Canvas
        my_canvas.create_window((0, 0), window=second_frame, anchor="nw")

        for i, p in enumerate(self.processes):
            var = tk.IntVar()
            c = tk.Checkbutton(second_frame, text=p,
                               variable=var,
                               onvalue=1, offvalue=0)
            c.grid(row=i, column=0, pady=1, padx=0)
            self.vars.append(var)

        self.root.mainloop()

    def state(self):
        return [p for p, v in zip(self.processes, self.vars) if v.get() == 1]

    def pick(self):
        self.chosen_processes = self.state()
        self.update1['text'] = ", ".join(self.chosen_processes)

    @staticmethod
    def show_warning():
        text = "Выбранно 0 процессов. Проверьте, что вы отметили галочками нужные, и нажмите кнопку `Выбрать`." + \
               "\n\nТолько после этого нажмите `Подтвердить`"
        messagebox.showinfo(title="Warning", message=text)

    def close(self):
        if len(self.chosen_processes) == 0:
            self.show_warning()
        else:
            self.root.destroy()


class GuiReport:
    def __init__(self):
        self.root = tk.Tk()

        self.file_now = None
        self.file_prev = None
        self.processes = None

        self.root.title("Еженедельный отчет по NPS")
        self.root.resizable(False, False)
        self.font = Font(self.root, size=10, weight=BOLD)

        lbl1 = tk.Label(self.root, text="Файл за текущую неделю", font=self.font)
        lbl1.grid(column=0, row=4)

        btn1 = tk.Button(self.root, text="Выбрать", command=self.clicked_now)
        btn1.grid(column=1, row=4)

        self.update1 = tk.Label(self.root)
        self.update1.grid(column=0, row=7)

        lbl2 = tk.Label(self.root,
                        text="Файл за предыдущую неделю",
                        font=self.font)
        lbl2.grid(column=0, row=8)

        btn2 = tk.Button(self.root, text="Выбрать", command=self.clicked_prev)
        btn2.grid(column=1, row=8)

        self.update2 = tk.Label(self.root)
        self.update2.grid(column=0, row=9)

        btn = tk.Button(self.root, text="Определить категории", command=self.close)
        btn.grid(column=0, row=12)

        self.root.mainloop()

    def clicked_now(self):
        file = fd.askopenfilename(parent=self.root,
                                  title='Выберите файл за текущий период')

        self.file_now = file
        sp = file.split("/")
        self.update1['text'] = sp[-1]

    def clicked_prev(self):
        file = fd.askopenfilename(parent=self.root,
                                  title='Выберите файл за текущий период')

        self.file_prev = file
        self.update2['text'] = file.split("/")[-1]

    @staticmethod
    def get_proc_from_file(file):
        df = pd.read_excel(file)
        return list(df['Process'].unique())

    def get_process(self):
        ls1 = self.get_proc_from_file(self.file_now)
        ls2 = self.get_proc_from_file(self.file_prev)

        return sorted(list(set(ls1 + ls2)))

    def close(self):
        self.processes = self.get_process()
        self.root.destroy()


if __name__ == '__main__':
    g = GuiReport()
    gp = GuiProcess(g.processes)
    print(g.file_now)
    print(g.file_prev)
    print(gp.chosen_processes)
