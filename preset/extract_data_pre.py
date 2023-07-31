#!/usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 2023/5/10 20:19
# @Author  : Martin
# @FileName: extract_data.py
# @Software: PyCharm

import re
import os


# 提取txt中所需数据
def extract_data_pre(file):
    with open(file, 'r') as f:
        data = f.read()
        # 匹配目标数据并提取眼高眼宽眼域
        width_pattern = r'max width: (\d+) units, (\d+\.\d+) UI'
        height_pattern = r'max height: (\d+) units(?:, (\d+\.\d+) (?:units|mv))?'
        area_pattern = r'area: (\d+) units'
        fom_pattern = r'fom (\d+)'

        width_match = re.search(width_pattern, data)
        height_match = re.search(height_pattern, data)
        area_match = re.search(area_pattern, data)
        fom_match = re.search(fom_pattern, data)

        if height_match:
            if height_match.group(2):
                height = float(height_match.group(2))
            else:
                height = float(height_match.group(1))

        if width_match and area_match and fom_match:
            width = float(width_match.group(2))
            area = float(area_match.group(1))
            fom = float(fom_match.group(1))

            # 将结果写入列表并返回
            result = (os.path.dirname(file), os.path.splitext(os.path.basename(file))[0], [width, height, area, fom])
            # print(result)
            return result
        else:
            print('未找到匹配的数据。')
