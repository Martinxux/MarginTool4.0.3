#!/usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 2023/5/10 20:20
# @Author  : Martin
# @FileName: compare_data.py
# @Software: PyCharm
import os
from preset.extract_data_pre import extract_data_pre


def compare_data_pre(files, file_path, log_func=None):
    data_dict = {}
    for file in files:
        temp_data = extract_data_pre(file)
        if temp_data:
            folder_path, file_name, data = temp_data
            folder_name = os.path.basename(folder_path)  # 获取路径中的最后一个文件夹名
            if folder_name not in data_dict:
                data_dict[folder_name] = {}
            data_dict[folder_name][file_name] = data

    # 将数据列表转换成字符串
    data_str_list = []
    for folder_name, file_dict in data_dict.items():
        data_str_list.append(f"\n{folder_name}, width, height, area, fom")
        for file_name, data in file_dict.items():
            data_str_list.append(f"{file_name}, {', '.join([str(item) for item in data])}")

    data_str = '\n'.join(data_str_list)

    # 将数据写入文本文件
    with open(file_path, 'a+') as f:
        f.write(data_str)
    # print(file_path)
    # 将记录写入日志文件
    if log_func:
        log_func('提取眼高眼宽：')
        log_func('\n'.join(
            [f"{os.path.basename(path)}/{file_name}, {', '.join([str(r) for r in data])}" for path, file_dict in
             data_dict.items() for file_name, data in file_dict.items()]
        )
         )
        log_func('[=============所有Preset已提取=============]')

    # return data_str_list
