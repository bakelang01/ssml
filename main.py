# This Python file uses the following encoding: utf-8
# 作者：black_lang
# 创建时间：2023/9/5 20:40
# 文件名：main.py

from ssml import SSML
import config

if __name__ == '__main__':
    ss = SSML()
    # 加载参数
    ss.config(config.data)
    # 获取学校信息
    school_li=ss.get_school_li()
    print(f'学校数量数量：{len(school_li)}')
    print("学校列表信息获取成功！")
    # 获取学校专业信息
    ss.get_zhuanye()
    # 保存
    ss.save('data2.xlsx')

