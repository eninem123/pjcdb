"""
可以将当前目录的ppt文件全部解密
"""

import win32com.client.dynamic
import os
import time
# 思路：用能解密的电脑打开ppt，然后另存为不能加密的文件类型比如xx.temp  xx.ini，然后拿自己电脑改后缀
g = os.walk(os.getcwd())
if not os.path.isdir(os.getcwd()+'\\test'):
    os.mkdir(os.getcwd()+'\\test')  # 检查有没这个文件夹,没有就创建新文件夹
newpath = os.getcwd()+'\\test\\' # 获取新文件夹路径
App = win32com.client.Dispatch("PowerPoint.Application")  # 创建打开PPT的对象
for path,dir_list,file_list in g:
    for file_name in file_list:
        # 遍历当前目录的文件名有哪些
        check_file = newpath + file_name
        #print(check_file)
        if not (os.path.exists(check_file)or os.path.exists(check_file+ ".ini")):  # 检查程序所在目录是否有已经解密过的，有解密过的不重复解密
            allfile = os.path.join(path, file_name)  # 所有文件路径
            if allfile.endswith('.pptx')or allfile.endswith('.ppt') :
                pptfile = allfile  # ppt的路径
                Newfile = os.getcwd()+'\\test\\' + file_name + ".ini"  # 新文件的路径
                try:
                    Presentation = App.Presentations.Open(pptfile, WithWindow=0)
                except Exception as e:
                    print(e)
                finally:
                    Presentation.SaveAs(Newfile)  # 解密，另存为ini文件
                    Presentation.close()
App.Quit()
gg = os.walk(r".\\test")
# 将ini重新转化为pptx
for new_path,new_dir_list,new_file_list in gg:
    for new_file_name in new_file_list:
        # print(new_file_name)
        if new_file_name.endswith('.ini'):
            os.rename(".\\test\\"+new_file_name, ".\\test\\"+new_file_name[:-4])
