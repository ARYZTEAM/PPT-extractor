'''
本程序用于生成脚本文件用于提取PPT
工作原理是首先需要运行该目录中的鼠标.exe后手动依次按顺序双击打开PPT中嵌入的PPT
然后脚本会根据按键生成一个脚本模拟点击加上手动保存。
'''
## ================================================
##              分析mouse.txt
## ================================================



content = ""
def openmousetxt():
    try:
        with open("./mouse.txt","r") as obj:
            global content
            content = obj.readlines()
    except:
        print("文件打开失败")


openmousetxt()
i = 0
while i < len(content):
    print(f"i is {i} time: {content[i]} aix:{content[i+2].strip()}")
    i += 4