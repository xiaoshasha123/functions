"""
说明：
1.版本：python3.5.3
2.功能：类FilesHandle“读取txt、doc、docx文件的内容”，所需参数为“文件路径”，
    需要通过“路径”调用辅助工具antiword读取doc文件内容，修改常量ANTIWORD_PATH
3.使用：实例化类后，调用方法handle()，ep：a = FilesHandle(absolute_path) content = a.handle()
4.执行过程：初始化-》handle-》判断文件格式-》1.txt：判断编码格式-》读文件
                                            2.word：区分doc和docx-》读取内容
"""
# -*- coding: UTF-8 -*-
import chardet
import docx
import subprocess
import os


class FilesHandle(object):
    """
    txt、doc、docx文件内容读取类
    """

    # antiword.exe绝对路径
    ANTIWORD_PATH = r"./antiword/antiword"

    def __init__(self, path: str):
        """
        初始化
        :param path: 文件路径
        """
        self.path = path

    def handle(self)->str:
        # 判断文件类型
        res = self.file_format_validation()
        content = ""
        if not res["error"]:
            try:
                if res["message"] == 'txt':
                    # txt文件编码格式解析
                    code_format = self.detect_code()
                    # 读取txt文件
                    res = self.read_text(code_format)
                    content = res["message"]
                elif res["message"] == 'word':
                    # 读取word文件
                    res = self.read_word()
                    content = res["message"]
                return content
            except Exception as e:
                return str(e)

        else:
            content = res["message"]
            return content

    def file_format_validation(self)->dict:
        """
        判断文件格式
        :return: 文件格式字典
        """
        allow_suffix = ["txt", "doc", "docx"]
        file_suffix = str(self.path.split(".")[-1]).lower()
        message = ""
        if file_suffix in allow_suffix:
            if file_suffix == "txt":
                message = "txt"
            elif file_suffix == "doc" or file_suffix == "docx":
                message = "word"
            return {"error": 0, "message": message}
        else:
            message = "文件格式不正确"
            return {"error": 1, "message": message}

    def detect_code(self)->str:
        """
        解析txt编码格式
        :return: 编码格式
        """
        try:
            with open(self.path, 'rb') as file:
                data = file.read(200000)
                dicts = chardet.detect(data)
                # print('***', dicts)
            return dicts["encoding"]
        except Exception as e:
            raise e

    def read_text(self, code_format: str)->dict:
        """
        读取txt文件
        :param code_format: txt文件编码格式
        :return: 文本信息
        """
        # 判断txt文件大小，单位为Mb, 区分限制值为0.5Gb
        text_size = os.path.getsize(self.path) / float(1024 * 1024)
        limit_size = 512 * text_size

        if text_size <= limit_size:
            try:
                with open(self.path, "r", encoding=code_format) as f:
                    content = f.read()
                return {"error": 0, "message": content}
            except Exception as e:
                raise e
        else:
            content_list = []
            try:
                with open(self.path, "r", encoding=code_format) as f:
                    for line in f:
                        content_list.append(line)
                content = ''.join(content_list)
                return {"error": 0, "message": content}
            except Exception as e:
                raise e

    def read_word(self)->dict:
        """
        读取word内容
        :return: 内容字典
        """
        suffix = self.path.split(".")[-1]
        content_dict = {}
        content_list = []
        try:
            if suffix == "docx":
                file = docx.Document(self.path)
                # 输出每一段的内容
                for para in file.paragraphs:
                    content_list.append(para.text)
                content_dict["message"] = "".join(content_list)

            elif suffix == "doc":
                content = subprocess.check_output([FilesHandle.ANTIWORD_PATH, "-mUTF-8", self.path])
                content_dict["message"] = content.decode("utf-8")

            return content_dict

        except Exception as e:
            raise e


# 模块内自测
if __name__ == '__main__':
    a = FilesHandle(r"./测试文件/8031基层党支部基础工作标准模板.docx")
    print(a.handle())
