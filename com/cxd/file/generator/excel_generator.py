# -*- coding:UTF-8 -*-
import openpyxl
import sys
import os
import shutil

message = None

def filesBackUp(src):
    outputdir = os.path.join(src, "output")
    message =  "判断目标地址是否已存在:"
    if os.path.exists(outputdir):
        message += "目标地址已存在\n进行删除。。。\n"
        shutil.rmtree(outputdir)
        message += "删除成功"
    print(message)

    message  = "开始将源文件或目录拷贝至目标文件夹。。。\n"
    try:
        shutil.copytree(src, outputdir)
        message += "拷贝成功"
    except Exception as err:
        message  += repr(err)
    finally:
        print(message)

    return outputdir


if __name__ == '__main__':
    files = r'E:\python\office_generator\test'
    print("读取输入： " + files)
    inputdir = None
    outputdir = None
    message = "判断输入内容： "
    try:
        if os.path.isfile(files):
            message += "输入类型为文件类型"
            inputdir = os.path.dirname(files)

        elif os.path.isdir(files):
            message += "输入类型为文件夹"

            inputdir = files
        else:
            message += "判断失败，请检查输入是否为文件或文件夹"
            raise Exception("Error: 判断失败，请检查输入是否为文件或文件夹")
    except Exception as err:
        sys.exit(1)
    finally:
        print(message)

        
    outputdir = filesBackUp(inputdir)
    print(outputdir)