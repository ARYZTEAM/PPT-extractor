import os
import ppt
from treelib import Tree
import zipfile

# 最终生成的文件树
# 其节点的identifier和tag都是文件的路径
ppttree = Tree()
dirindex = 0
fileindex = 0


# 待解决节点添加问题
def rename_all_files(directorys, parentnode):
    """
    该函数会重命名embeding文件夹中解压出来的文件，并更改文件的扩展名: bin->ppt，并将其加入树节点
    而对于ppt和pptx文件则直接加入树节点
    遇到不是ppt，pptx，bin文件则会直接报错退出
    :param directorys: 文件夹
    :param parentnode:父节点
    """
    to_be_renamed = os.listdir(directorys)
    global fileindex
    for i in to_be_renamed:
        # 判断文件的大小不为0
        if os.path.getsize(directorys + "\\" + i) != 0:
            os.rename(directorys + "\\" + i, directorys + "\\" + "file" + str(fileindex) + i[i.rfind("."):])
            i = "file" + str(fileindex) + i[i.rfind("."):]
            fileindex += 1
            if i[i.rfind("."):] == ".bin":
                os.rename(directorys + "\\" + i, directorys + "\\" + i[:i.rfind(".")] + ".ppt")
                ppttree.create_node(tag=i[:i.rfind(".")] + ".ppt",
                                    identifier=directorys + "\\" + i[:i.rfind(".")] + ".ppt", parent=parentnode)

            elif i[i.rfind("."):] == ".ppt":
                ppttree.create_node(tag=i, identifier=directorys + "\\" + i, parent=parentnode)

            elif i[i.rfind("."):] == ".pptx":
                ppttree.create_node(tag=i, identifier=directorys + "\\" + i, parent=parentnode)
            else:
                if os.path.isdir(directorys + "\\" + i):
                    ppttree.create_node(tag=i, identifier=directorys + "\\" + i, parent=parentnode)
                else:
                    print(directorys + "\\" + i)
                    print("异常：发现特殊类型的文件在embeddings中，退出处理")
                    exit(1)
        else:
            try:
                os.system("del" + " " + directorys + "\\" + i)
            except:
                pass
            # os.system("del" + " " + i)


# 输入一个PPTX文件查看是否解压，output是解压目录
def weather_extract(fullfilename):
    """
    输入一个PPTX文件，判断embedding中是否嵌入了文件
    如果能够解压，则解压到目标文件的同一文件夹下，文件夹名字为dir+dirindex
    解压后调用rename函数重命名解压后的文件
    :param fullfilename:
    :return: 真或者假
    """
    if os.path.isdir(fullfilename[:fullfilename.rfind("\\")]):
        os.chdir(fullfilename[:fullfilename.rfind("\\")])
    else:
        print("weather_extract：无法移动到待解压PPTX文件的目录下")
        exit(1)
    global dirindex
    output = "dir" + str(dirindex)
    des_filename = fullfilename[:fullfilename.rfind(".")] + ".zip"
    os.system("copy" + " " + fullfilename + " " + des_filename)
    zip_f = zipfile.ZipFile(des_filename)
    list_zip_f = zip_f.namelist()  # zip文件中的文件列表名
    to_be_extracted = []
    for i in list_zip_f:
        if "ppt/embeddings" in i:
            to_be_extracted.append(i)
    if len(to_be_extracted) != 0:
        for j in to_be_extracted:
            zip_f.extract(j, output)  # 第二个参数指定输出目录，此处保存在当前目录下的output文件夹中
        zip_f.close()
        dirindex += 1
        os.chdir(os.getcwd() + "\\" + output)
        os.system("move" + " " + os.getcwd() + "\\ppt\\embeddings\\*" + " " + os.getcwd())
        os.system("RMDIR /S/Q" + " " + os.getcwd() + "\\ppt")
        # 这里必须大写的Q
        rename_all_files(os.getcwd(), fullfilename)
        return True
    else:
        zip_f.close()
        return False


def ispptorpptx(fullfilename):
    """
    输入PPT或者PPTX文件，对于PPT文件将会执行PPT转换为PPTX后将其节点更新为PPTX后判断是否内部
    还有嵌入的PPT文件，如果有的话，更新节点的data为1，否则为0
    对于PPTX直接判断是否能够解压，如果可以的话更新data为1，否则为0
    :param fullfilename:
    """
    if fullfilename.endswith('.ppt'):
        ppt.ppttopptx(fullfilename)
        node = ppttree.get_node(fullfilename)
        temp = node.tag + "x"
        ppttree.update_node(node.identifier, identifier=fullfilename + "x")
        ppttree.update_node(node.identifier, tag = temp)
        # output = node.identifier
        # output = output.removesuffix(".pptx")
        # output = output[output.rfind("\\")+1:]
        # if len(output) == 0:
        #     print("文件输出路径错误")
        #     exit(1)
        if weather_extract(fullfilename + "x"):
            ppttree.update_node(fullfilename + "x", data=1)

        else:
            ppttree.update_node(fullfilename + "x", data=0)

    elif fullfilename.endswith('.pptx'):
        node = ppttree.get_node(fullfilename)
        # output = node.identifier
        # output = output.removesuffix(".pptx")
        # output = output[output.rfind("\\")+1:]
        # if output is None:
        #     print("文件输出路径错误")
        #     exit(1)
        if weather_extract(node.identifier):
            ppttree.update_node(node.identifier, data=1)
        else:
            ppttree.update_node(node.identifier, data=0)
    else:
        print()


def start_extract(filepath):
    global fileindex
    if filepath.endswith('.pptx'):
        os.rename(filepath, filepath[:filepath.rfind("\\")] + "\\" + "file" + str(fileindex) + ".pptx")
        filepath = filepath[:filepath.rfind("\\")] + "\\" + "file" + str(fileindex) + ".pptx"
        fileindex += 1
        ppttree.create_node(tag=filepath, identifier=filepath,data=1)
        #        ispptorpptx(filepath)
        if weather_extract(filepath):
            levels = 1
            while ppttree.size(levels) > 0:
                # 待优化通过添加过滤条件加快查找速度
                for i in ppttree.expand_tree():
                    if ppttree.level(i) == levels:
                        ispptorpptx(i)
                levels += 1

            # print("-"*50)
            #
            # for i in ppttree.expand_tree():
            #     print(f"i:{i} data:{ppttree.get_node(i).data}")
            # print("-"*50)

    else:
        print("非pptx文件，先转化为pptx文件")


if __name__ == "__main__":
    start_extract("C:\\Users\\Administrator\\Documents\\ppttest\\OS1.pptx")
    ppttree.show()
