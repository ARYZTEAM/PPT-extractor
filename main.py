import os
import ppt
from treelib import Node,Tree
import zipfile

ppttree = Tree()


#输入zip文件的完全路径，和指定解压的文件夹,运行目录的子文件下
def extract_zip(zip_filename,extract_path):
    if(os.path.exists(zip_filename) is False):
        print("解压的ZIP文件不存在")
        exit(1)

    zip_f = zipfile.ZipFile(zip_filename)
    list_zip_f = zip_f.namelist()  # zip文件中的文件列表名
    for zip_fn in list_zip_f:
        zip_f.extract(zip_fn, extract_path)  # 第二个参数指定输出目录，此处保存在当前目录
    zip_f.close()


##待解决节点添加问题
def rename_all_files(directorys,parentnode):
    to_be_renamed = os.listdir(directorys)
    for i in to_be_renamed:
        if os.path.getsize(directorys + "\\" + i) != 0:
            if(i[i.rfind("."):] == ".bin"):
                os.rename(directorys + "\\"  + i,directorys + "\\"  + i[:i.rfind(".")] + ".ppt")
                ppttree.create_node(tag=i[:i.rfind(".")] + ".ppt",identifier=directorys + "\\"  + i[:i.rfind(".")] + ".ppt",parent=parentnode)

            elif(i[i.rfind("."):] == ".ppt"):
                ppttree.create_node(tag=i, identifier=directorys + "\\" + i, parent=parentnode)

            elif(i[i.rfind("."):] == ".pptx"):
                ppttree.create_node(tag=i, identifier=directorys + "\\" + i, parent=parentnode)
            else:
                if(os.path.isdir(directorys + "\\" + i)):
                    ppttree.create_node(tag=i, identifier=directorys + "\\" + i,parent=parentnode)
                else:
                    print(directorys + "\\" + i)
                    print("异常：发现特殊类型的文件在embeddings中，退出处理")
                    exit(1)
        else:
            try:
                os.system("del" + " " + directorys + "\\" + i)
            except:
                pass
            #os.system("del" + " " + i)


#输入一个PPTX文件查看是否解压，output是解压目录
def weather_extract(fullfilename):
    if os.path.isdir(fullfilename[:fullfilename.rfind("\\")]):
        os.chdir(fullfilename[:fullfilename.rfind("\\")])
    else:
        print("weather_extract不是文件夹")
        exit(1)
    output = fullfilename
    output = output.removesuffix(".pptx")
    output = output[output.rfind("\\")+1:]
    if len(output) == 0:
        print("文件输出路径错误")
        exit(1)

    des_filename = fullfilename.removesuffix(".pptx") + ".zip"
    os.system("copy" + " " + fullfilename + " " + des_filename)
    zip_f = zipfile.ZipFile(des_filename)
    list_zip_f = zip_f.namelist()  # zip文件中的文件列表名
    to_be_extracted = []
    for i in list_zip_f:
        if "ppt/embeddings" in i:
            to_be_extracted.append(i)
    if len(to_be_extracted) != 0 :
        for j in to_be_extracted:
            zip_f.extract(j,output)  # 第二个参数指定输出目录，此处保存在当前目录
        zip_f.close()
        os.chdir(os.getcwd() + "\\" + output)
        os.system("move" + " " +os.getcwd() + "\\ppt\\embeddings\\*" + " " + os.getcwd())
        os.system("RMDIR /S/Q" + " " + os.getcwd() + "\\ppt")
        ##这里必须大写的Q
        rename_all_files(os.getcwd(),fullfilename)
        return True
    else:
        zip_f.close()
        return False






def ispptorpptx(fullfilename):

    if(fullfilename.endswith('.ppt')):
        ppt.ppttopptx(fullfilename)
        node = ppttree.get_node(fullfilename)
        ppttree.update_node(node.identifier,identifier= fullfilename + "x")
        # output = node.identifier
        # output = output.removesuffix(".pptx")
        # output = output[output.rfind("\\")+1:]
        # if len(output) == 0:
        #     print("文件输出路径错误")
        #     exit(1)
        if weather_extract(node.identifier):
            ppttree.update_node(node.identifier,data=1)
            print("尝试更改"+node.identifier+"的数据")

        else:
            ppttree.update_node(node.identifier,data=0)

    elif(fullfilename.endswith('.pptx')):
        node = ppttree.get_node(fullfilename)
        # output = node.identifier
        # output = output.removesuffix(".pptx")
        # output = output[output.rfind("\\")+1:]
        # if output is None:
        #     print("文件输出路径错误")
        #     exit(1)
        if weather_extract(node.identifier):
            ppttree.update_node(node.identifier,data=1)
            print("尝试更改"+node.identifier+"的数据为1")
        else:
            ppttree.update_node(node.identifier,data=0)
    else:
        print()


def data_tree():
    ppttree = Tree()
    ppttree.create_node("harrt","whatisthis")
    ppttree.show()


def start_extract(filepath):
    if(filepath.endswith('.pptx')):
        ppttree.create_node(tag=filepath,identifier=filepath)
#        ispptorpptx(filepath)
        if (weather_extract(filepath)):
            levels = 1
            if(ppttree.size(levels) > 0 ):
                ##待优化通过添加过滤条件加快查找速度
                for i in ppttree.expand_tree():
                    if(ppttree.level(i) == levels):
                        ispptorpptx(i)

            for i in ppttree.expand_tree(filter=lambda x:x.data ==0):
                print(i)
            print("上面是0的")
            for i in ppttree.expand_tree(filter=lambda x:x.data ==None):
                print(i)
            print("上面是None的")



    else:
        print("非pptx文件，先转化为pptx文件")


if __name__ == "__main__":
    print("hello world")
    start_extract("C:\\Users\\Administrator\\Documents\\ppttest\\OS1.pptx")
    ppttree.show()