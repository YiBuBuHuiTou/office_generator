# -*- coding:UTF-8 -*-
import openpyxl
import os
import shutil


def copy_excel(src, dist):
    files = os.listdir(src)
    for file in files:
        if os.path.isdir(file):
            copy_excel(file,dist)
        else:
            shutil.c

    return


if __name__ == '__name__':
    copy_excel()




