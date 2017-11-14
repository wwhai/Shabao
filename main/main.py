# coding=utf-8
from form.form import MainWindow, lis

if __name__ == "__main__":
    window = MainWindow()
    window.write_excel(lis)
    window.show()
