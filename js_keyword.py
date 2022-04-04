import pandas as pd
import datetime
import os
import re
import tkinter.messagebox
from tkinter import filedialog


class Datas:
    def __init__(self, nav_path):
        # 获取文件所在目录位置
        self.folder_path = os.path.dirname(nav_path)
        # 获取文件夹下所有的文件
        self.file_names = os.listdir(self.folder_path)
        # 获取到导航表中的Asin-Rank的映射关系字典
        self.rank_asin = dict(pd.read_excel(nav_path).to_dict(orient='split')['data'])
        # Asin-Rank映射，key与value互换
        self.asin_rank = dict(zip(self.rank_asin.values(), self.rank_asin.keys()))
        # 创建一个导航Asin列表
        self.asin_list = list(self.rank_asin.values())
        # 创建一个空列表，用来接收所有可以取到数据的Asin
        self.asin_name_lists = list()
        # 创建一个空列表，用来存放文件中取出来的DateFrame
        self.datas = list()
        # 创建一个空值，用来接收拼接好的数据
        self.data = None
        self.save_path = self.folder_path + '/' + str(datetime.datetime.now().timestamp()) + '.xlsx'
        # 需要的字段
        self.need_field = ['关键词',
                           '精确匹配搜索量(过去30天)',
                           '广泛匹配搜索量(过去30天)',
                           '新品促销量参考',
                           '头条广告建议出价',
                           '精确PPC建议出价',
                           '广泛PPC建议出价',
                           '投放广告商品数量',
                           '对标ASIN自然排名']
        # 拼接匹配的字段
        self.match_field = ['关键词',
                            '精确匹配搜索量(过去30天)',
                            '广泛匹配搜索量(过去30天)',
                            '新品促销量参考',
                            '头条广告建议出价',
                            '精确PPC建议出价',
                            '广泛PPC建议出价',
                            '投放广告商品数量']

    def get_data(self):
        for file_name in self.file_names:
            if '.csv' in file_name:
                # 获取详细的数据
                detail_data = pd.read_csv(self.folder_path + '/' + file_name, skiprows=3)

                # 只获取想要的数据
                keyword_data = detail_data[self.need_field]
                # 获取描述字符串
                data_des = pd.read_csv(self.folder_path + '/' + file_name, nrows=2)
                # 提取出描述字符串中Asin的部分
                asin_names = re.findall(r"'(.*?)'", data_des.iloc[1, 0])[0]
                # 将Asin字符串分割，获得一个Asin列表
                asin_name_list = asin_names.split(',')

                for asin_sort in range(len(asin_name_list)):
                    # 去掉ASIN前后的空格
                    asin_name = asin_name_list[asin_sort].strip()
                    if asin_sort == 0:  # 说明这个Asin要特殊处理，因为它的自然排名字段名为：【对标ASIN自然排名】，而非Asin
                        keyword_data = keyword_data.copy()
                        keyword_data.rename(columns={'对标ASIN自然排名': asin_name}, inplace=True)
                    else:  # 要将其他Asin列的数量追加进keyword_data中
                        keyword_data = keyword_data.copy()
                        keyword_data.loc[:, asin_name] = detail_data.loc[:, asin_name]
                    # 将ASIN存放到列表中
                    self.asin_name_lists.append(asin_name)

                # 将整理好的数据加入到待合并的数据列表中
                self.datas.append(keyword_data)

    def merge_data(self):
        # 合并原始数据
        file_num = len(self.datas)
        if file_num != 0:
            self.data = self.datas[0]
            if file_num > 1:
                for i in range(1, file_num):
                    self.data = pd.merge(self.data, self.datas[i], how='outer', on=self.match_field)
        else:
            tkinter.messagebox.showinfo('提示', '没有找到相关数据')

    # 格式化数据
    def format_data(self):
        """
        将除关键词以外的列，全部数字化
        :return:
        """
        for col in range(1, len(self.data.columns)):
            self.data.iloc[:, col] = self.data.iloc[:, col].apply(self.del_sign)

    @staticmethod
    def del_sign(value):
        """
        用于将将 >、 <、 ---以及货币符号替换掉,并将数据全部转化为浮点数
        :param value:
        :return:
        """
        value = str(value)
        if '<' in value:
            # return float(value.replace('<', ''))
            return 0
        elif '>' in value:
            # return float(value.replace('>', ''))
            return None
        elif '$' in value:
            return float(value.replace('$', ''))
        elif '€' in value:
            return float(value.replace('€', ''))
        elif '£' in value:
            return float(value.replace('€', ''))
        elif '---' in value:
            return None
        else:
            return float(value)

    def search_num(self):
        self.data['广泛/精准搜索量'] = self.data['广泛匹配搜索量(过去30天)'] / self.data['精确匹配搜索量(过去30天)']

        self.data['30天广泛搜索量规范化'] = (self.data['广泛匹配搜索量(过去30天)'] - self.data['广泛匹配搜索量(过去30天)'].min()) / \
                                   (self.data['广泛匹配搜索量(过去30天)'].max() - self.data['广泛匹配搜索量(过去30天)'].min())

        self.data['30天精准搜索量规范化'] = (self.data['精确匹配搜索量(过去30天)'] - self.data['精确匹配搜索量(过去30天)'].min()) / \
                                   (self.data['精确匹配搜索量(过去30天)'].max() - self.data['精确匹配搜索量(过去30天)'].min())

        self.data['广泛搜索量/精准搜索量规范化[1,4]区间最优,无法容忍下限为0.8，无法容忍上限为5'] = self.data['广泛/精准搜索量'].\
            apply(self.search_scale_specification)
        self.data['搜索量指数'] = self.data['30天广泛搜索量规范化'] * self.data['30天精准搜索量规范化'] * \
                             self.data['广泛搜索量/精准搜索量规范化[1,4]区间最优,无法容忍下限为0.8，无法容忍上限为5']

    def asin_num(self):
        self.data['相关个数'] = self.data.iloc[:, len(self.match_field):len(self.match_field) + len(self.asin_name_lists)].count(axis=1)
        # 相关个数规范化
        self.data['相关个数规范化'] = (self.data['相关个数'] - self.data['相关个数'].min()) / \
                               (self.data['相关个数'].max() - self.data['相关个数'].min())

    # 偏离量相关计算
    def deviation(self):
        asins = list(self.data.iloc[:, len(self.match_field):len(self.match_field) + len(self.asin_name_lists)].columns)
        for asin in asins:
            # self.data.loc[:, asin + ':第' + str(self.asin_rank[asin]) + '名偏移量'] =
            self.data.loc[:, asin + '偏移量'] = \
                self.data.loc[:, asin].apply(self.deviation_value, args=(self.asin_rank[asin],))

        # 创建一个新列表，用来接收偏移量和值
        deviation_sum = list()
        for index in self.data.index:
            dev_uni = self.data.loc[index, asins[0] + '偏移量':asins[len(asins) - 1] + '偏移量'].sum()
            deviation_sum.append(dev_uni)
        self.data.loc[:, '偏移量之和'] = pd.Series(deviation_sum, index=self.data.index)
        self.data['偏离程度反向规范化'] = (self.data['偏移量之和'].max() - self.data['偏移量之和']) / \
                                 (self.data['偏移量之和'].max() - self.data['偏移量之和'].min())
        self.data['自然排名相关程度指数'] = self.data['偏离程度反向规范化'] * self.data['相关个数规范化']

    @staticmethod
    def deviation_value(value, series_value):
        if pd.isnull(value):
            square = (series_value - 151) ** 2
        else:
            square = (series_value - value) ** 2
        return square

    @staticmethod
    def search_scale_specification(value):
        if value > 5 or value < 0.8:
            return 0
        elif value >= 4:
            return 1 - (value - 4) / (5 - 4)
        elif value >= 1:
            return 1
        else:
            return 1 - (1 - value) / (1 - 0.8)

    def result(self):
        self.data['标杆指数'] = self.data['自然排名相关程度指数'] * self.data['搜索量指数']
        self.data = self.data.loc[:, ['关键词', '标杆指数', '搜索量指数', '自然排名相关程度指数', '精确匹配搜索量(过去30天)', '广泛匹配搜索量(过去30天)',
                                      '新品促销量参考', '头条广告建议出价', '精确PPC建议出价', '广泛PPC建议出价', '投放广告商品数量']]
        self.data.sort_values(by='标杆指数', ascending=False, inplace=True)

    def save_data(self):
        path = pd.ExcelWriter(self.save_path)
        self.data.to_excel(path, sheet_name='原始数据合并', index=None)
        tkinter.messagebox.showinfo('提示', '数据分析已完成，文件存放路径为：' + self.save_path)
        path.save()

    def compare_list(self):
        asin_name_lists_set = set(self.asin_name_lists)
        if len(self.asin_name_lists) == len(asin_name_lists_set):
            # 值没有重复，可以下一步
            asin_name_set = set(self.asin_list)
            if (asin_name_set | asin_name_lists_set) == asin_name_set:
                if len(asin_name_set) != 0:
                    tkinter.messagebox.showinfo('恭喜', '数据准确，正在分析，请等待')
                    return True
            else:
                tkinter.messagebox.showinfo('错误', '数据源与导航表不匹配.导航中的Asin应与CSV表中的Asin完全一致')
        else:
            tkinter.messagebox.showinfo('错误', '数据源与导航表不匹配.导航中的Asin应与CSV表中的Asin完全一致')
        return False


def js_keyword():
    # 获取文件路径
    ask_nav_path = filedialog.askopenfilename()
    # 创建对象
    data = Datas(ask_nav_path)
    # 获取数据
    data.get_data()
    if data.compare_list():
        # 拼接数据
        data.merge_data()
        # 格式化数据为数字
        data.format_data()
        # 获取搜索指数
        data.search_num()
        # 获取有排名的的Asin个数
        data.asin_num()
        # 计算排名偏离
        data.deviation()
        # 计算标杆值
        data.result()
        # 保存数据
        data.save_data()


# if __name__ == '__main__':
#     ask_nav_path = filedialog.askopenfilename()
#     data = Datas(ask_nav_path)
#     data.get_data()
#     data.merge_data()
#     data.format_data()
#     data.search_num()
#     data.asin_num()
#     data.deviation()
#     data.result()
#     data.save_data()


