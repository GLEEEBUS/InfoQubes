from gui import GuiProcess, GuiReport
from negative_change import Report
import time


def main():
    g = GuiReport()
    print("текущий файл:", g.file_now)
    print('предыдущий файл:', g.file_prev)

    gp = GuiProcess(g.processes)
    print('процессы:', ", ".join(gp.chosen_processes))

    Report(now_file=g.file_now, prev_file=g.file_prev, processes=gp.chosen_processes)

    time.sleep(10)


if __name__ == '__main__':
    main()
