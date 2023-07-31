#!/usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 2023/5/11 10:21
# @Author  : Martin
# @FileName: compare.py
# @Software: PyCharm
from pcie.extra_data import extract_data
import numpy as np


def compare_data(files, file_path, file_list, log_func=None):
    data_list = []
    result_list = []
    for file in files:
        data_list.append(extract_data(file))

    # 将列表分成5个子列表
    data_lists = np.array_split(data_list, 5)
    # 遍历每个位置并获取对应位置的元素
    for i in range(len(data_lists[0])):
        # 获取每个子列表在当前位置上的元素
        lane_result = [sublist[i] for sublist in data_lists]
        # 找到每个键的最小值
        min_values = {}
        for result in lane_result:
            for key, value in result.items():
                if key not in min_values or value < min_values[key]:
                    min_values[key] = value

        # 记录结果字典
        result_dict = {'Min Result': min_values}
        # 将结果打印为列表形式
        result_list.append(list(result_dict['Min Result'].values()))

    with open(file_path, 'a+') as file:
        for i, item in enumerate(result_list):
            # 将 file_list 和 result_list 中的对应项用逗号连接，并写入同一行
            file.write(f"{file_list[i]},{','.join([str(r) for r in item])}\n")

    result_str = '\n'.join([str(r) for r in result_list])
    file_str = '\n'.join([str(f) for f in file_list])
    paired_str = '\n'.join([f"{r}:{f}" for r, f in zip(file_str.split('\n'), result_str.split('\n'))])

    log_func('[==========最小值已提取==========]')
    if log_func:
        log_func(paired_str)
