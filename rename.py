import docx
import os
import shutil

ori_file_path = 'D:/PycharmProjects/docx_demo/B/'
tar_file_path = 'D:/PycharmProjects/docx_demo/B/'


def main():
    #  遍历源文件夹所有文件  ===========================================================================================
    all_file_list = os.listdir(ori_file_path)
    total_file_num = len(all_file_list)
    index = 0
    for ori_file in all_file_list:
        index += 1
        ori_file_name = ori_file_path + ori_file
        #  读取源文件内容  获得标题 和 操作内容  =======================================================================
        tar_file_name = tar_file_path + ori_file[2:]
        #  拷贝文件  ==========================================================================
        try:
            shutil.copyfile(ori_file_name, tar_file_name)
            print('模板文件拷贝到临时文件成功')
        except Exception as e_string:
            print('模板文件拷贝到临时文件失败！错误代码：', end="")
            print(e_string)
        #  删除临时文件  ===================================================================================================
        os.remove(ori_file_name)
        print('源文件清除成功！')
        string = '共%d文件，已完成第%d个文件' % (total_file_num, index)
        print(string)


if __name__ == "__main__":
    main()
