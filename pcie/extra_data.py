#!/usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 2023/5/16 13:40
# @Author  : Martin
# @FileName: extra_data4.py
# @Software: PyCharm

import re


# 提取txt中所需数据并写入字典
def extract_data(file):
    with open(file, 'r') as f:
        data = f.read()
        # 匹配目标数据并提取眼高眼宽眼域
        width_pattern = r'max width: (\d+) units, (\d+\.\d+) UI'
        height_pattern = r'max height: (\d+) units(?:, (\d+\.\d+) (?:units|mv))?'
        area_pattern = r'area: (\d+) units'

        width_match = re.search(width_pattern, data)
        height_match = re.search(height_pattern, data)
        area_match = re.search(area_pattern, data)

        if height_match:
            if height_match.group(2):
                height = float(height_match.group(2))
            else:
                height = float(height_match.group(1))

        if width_match and area_match:
            width = float(width_match.group(2))
            area = float(area_match.group(1))

            # 将结果写入字典并返回
            result = {'width': width, 'height': height, 'area': area}
            return result
        else:
            print('未找到匹配的数据。')
