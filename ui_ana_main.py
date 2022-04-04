import pandas as pd
import tkinter as tk
import tkinter.messagebox
from tkinter import filedialog
import datetime
from js_keyword import js_keyword as js_keyword_ana


class Ui:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title('亚马逊分析小工具')
        self.window.geometry('600x500+500+200')
        self.window.resizable(width=False, height=False)

    def lab(self, text, bg, bd):
        lab = tk.Label(self.window, text=text, bg=bg, bd=bd,
                       width=50, height=1, fg='white',
                       font=('MicroSoft YaHei', 18, 'bold'))
        lab.pack()

    def button(self, text, command):
        button = tk.Button(self.window, text=text, command=command,
                           width=80, padx=200, pady=2, relief=tk.RAISED, cursor='hand2',
                           font=('MicroSoft YaHei', 12, 'bold', 'underline'),
                           bg='#407434', fg='white', bd=5)
        button.pack()


class ButtonMethod:
    def __init__(self):
        self.path = filedialog.askopenfilename()
        if self.path is not None and '.' in self.path:
            self.path_list = self.path.split('.')
            path_list_len = len(self.path_list)
            if self.path_list[path_list_len - 1] == 'csv':
                self.data = pd.read_csv(self.path)
            elif self.path_list[path_list_len - 1] == 'xlsx':
                self.data = pd.read_excel(self.path)
            else:
                tk.messagebox.showinfo('警告', '数据格式不正确,！！！')
                self.data = None

    def top_ana(self):
        if self.data is not None:
            try:
                self.data = self.data.loc[:, ['#', '排名', '产品名称', '卖家类型', '品牌', '售价', '月销量',
                                              '上架时间', '评分数', '星级', '类目', 'ASIN', 'Link']]
            except KeyError:
                tk.messagebox.showinfo('错误', '数据字段不全！！导入数据失败')
            else:
                try:
                    # 将售价处理为数字
                    self.data['售价'] = self.data['售价'].str.replace('$', '', regex=True).astype('float')
                except Exception:
                    self.data['售价'] = self.data['售价'].str.replace('€', '', regex=True).astype('float')
                finally:
                    # 若该链接没有品牌名，则空值替换为：“无品牌”，无星级，则替换为0
                    self.data.fillna({'品牌': '无品牌', '星级': 0}, inplace=True)
                    # 处理过后的数据，有任意值为空，都删除对应行******数据清洗完成
                    self.data.dropna(axis='index', how='any', inplace=True)
                    # 获取售价区间,为了保证所有的值都能透视到，最小值要减0.01
                    self.range('售价', num=0.01)
                    # 获取月销量区间,为了保证所有的值都能透视到，最小值要减1
                    self.range('月销量', num=1)
                    # 获取时间区间
                    self.range('上架时间', [-1, 100, 365, 100000000], time=True)
                    # 获取评分区间
                    self.range('评分数', [-1, 100, 500, 2000, 10000, 100000000])
                    # 获取星级区间
                    self.range('星级', [-1, 3.7, 4.2, 4.8, 5])

                    # 生成数据透视表
                    data_price = self.working_pivot_tables('售价区间')
                    data_moth_sale = self.working_pivot_tables('月销量区间')
                    data_days = self.working_pivot_tables('上架时间区间')
                    data_review = self.working_pivot_tables('评分数区间')
                    data_star = self.working_pivot_tables('星级区间')
                    data_category = self.working_pivot_tables('类目')
                    data_seller_type = self.working_pivot_tables('卖家类型')
                    data_brand = self.working_pivot_tables('品牌', True)

                    # 按链接上架的月份统计链接数
                    link_star_month = pd.DataFrame()
                    link_star_month['上架月份'] = self.data['上架时间'].dt.month.value_counts().index
                    link_star_month['链接数'] = self.data['上架时间'].dt.month.value_counts().values

                    # 选出潜在竞争对手 (上架100天以内，或Review数小于200的链接)
                    data_competitor = self.data.loc[(self.data['评分数'] < 200) |
                                                    ((datetime.datetime.now() - self.data['上架时间']).dt.days < 100), :]

                    '''
                    将打好标的数据存放到Excel中
                    '''
                    self.save_data(self_data={'data': self.data, 'sheet_name': 'Top分析',
                                              'start_col': 0, 'start_row': 0, 'header': True, 'index': None},
                                   data_summary={'data': ButtonMethod.top_ana_sum(self.data), 'sheet_name': 'Top分析',
                                                 'start_col': 20, 'start_row': 0, 'header': None, 'index': True},
                                   link_star_month={'data': link_star_month, 'sheet_name': 'Top分析',
                                                    'start_col': 20, 'start_row': 10, 'header': True, 'index': None},
                                   data_price={'data': data_price, 'sheet_name': 'Top分析',
                                               'start_col': 25, 'start_row': 0, 'header': True, 'index': True},
                                   data_moth_sale={'data': data_moth_sale, 'sheet_name': 'Top分析',
                                                   'start_col': 25, 'start_row': 6, 'header': True, 'index': True},
                                   data_days={'data': data_days, 'sheet_name': 'Top分析',
                                              'start_col': 25, 'start_row': 12, 'header': True, 'index': True},
                                   data_review={'data': data_review, 'sheet_name': 'Top分析',
                                                'start_col': 25, 'start_row': 17, 'header': True, 'index': True},
                                   data_star={'data': data_star, 'sheet_name': 'Top分析',
                                              'start_col': 25, 'start_row': 24, 'header': True, 'index': True},
                                   data_category={'data': data_category, 'sheet_name': 'Top分析',
                                                  'start_col': 35, 'start_row': 0, 'header': True, 'index': True},
                                   data_seller_type={'data': data_seller_type, 'sheet_name': 'Top分析',
                                                     'start_col': 25, 'start_row': 30, 'header': True, 'index': True},
                                   data_brand={'data': data_brand, 'sheet_name': 'Top分析',
                                               'start_col': 25, 'start_row': 35, 'header': True, 'index': True},
                                   data_competitor={'data': data_competitor, 'sheet_name': '潜在竞争对手（Review<200或'
                                                                                           '上架天数小于100天）',
                                                    'start_col': 0, 'start_row': 0, 'header': True, 'index': None})

    def keywords_ana(self):
        if self.data is not None:
            try:
                self.data = self.data.loc[:, ['关键词', '月搜索量', '相关度', '月购买量', '购买率',
                                              '点击集中度', '商品数', '均价', '评分数', '评分值', 'PPC价格']]
            except KeyError:
                tk.messagebox.showinfo('错误', '数据字段不全！！导入数据失败')
            else:
                try:
                    # 将均价处理为数字
                    self.data['均价'] = self.data['均价'].str.replace('$', '', regex=True).astype('float')
                except Exception:
                    self.data['均价'] = self.data['均价'].str.replace('€', '', regex=True).astype('float')
                finally:
                    # 扔掉没有PPC价格值的数据
                    self.data.dropna(axis='index', how='any', inplace=True)
                    # 计算出广告成单价格
                    self.data['单个广告订单成单费用'] = round(self.data['PPC价格'] / self.data['购买率'], 2)
                    self.data['单个广告订单理想成单费用'] = round(self.data['单个广告订单成单费用'] * 0.7, 2)

                    # 获取月搜索量区间,为了保证所有的值都能透视到，最小值要减1
                    month_num = self.range('月搜索量', num=1)
                    # 获取购买率区间,为了保证所有的值都能透视到，最小值要减0.0001
                    percent = self.range('购买率', num=0.0001, number_of_digits=4)
                    # 获取单个广告理想成单区间,为了保证所有的值都能透视到，最小值要减0.01
                    ad = self.range('购买率', num=0.01)
                    # 重要关键词列表
                    data_important = self.data.loc[
                                     (self.data['月搜索量'] >= month_num[1]) &
                                     ((self.data['购买率'] >= percent[1]) |
                                      (self.data['单个广告订单理想成单费用'] <= ad[0])), :]

                    self.save_data(self_data={'data': self.data, 'sheet_name': '关键词源数据',
                                   'start_col': 0, 'start_row': 0, 'header': True, 'index': None},
                                   self_data_summary={'data': ButtonMethod.key_ana_sum(self.data), 'sheet_name': '关键词源数据',
                                                      'start_col': 18, 'start_row': 0, 'header': None, 'index': True},
                                   data_important={'data': data_important, 'sheet_name': '重点关键词数据（购买率超过75%，流量超过75%）',
                                   'start_col': 0, 'start_row': 0, 'header': True, 'index': None},
                                   importrant_data_summary={'data': ButtonMethod.key_ana_sum(data_important), 'sheet_name': '重点关键词数据（购买率超过75%，流量超过75%）',
                                                            'start_col': 18, 'start_row': 0, 'header': None, 'index': True}
                                   )

    @staticmethod
    def key_ana_sum(data):
        # 统计结果
        # 月销售量总计
        month_search_sum = data['月搜索量'].sum()
        # 月购买量总计
        month_buy_sum = data['月购买量'].sum()
        # 展示价均值
        show_price_mean = data['均价'].mean()
        # 展示价中位值
        show_price_median = data['均价'].median()
        # 成交价均值
        clinch_price_mean = (data['月购买量'] * data['均价']).sum() / data['月购买量'].sum()
        # 总销售额
        sales_sum = (data['月购买量'] * data['均价']).sum()
        # 单个广告订单理想成单费用均值
        ad_order_price_mean = (data['月购买量'] * data['单个广告订单理想成单费用']).sum() \
                              / data['月购买量'].sum()

        dec_dic = {
            '月搜索量总计': month_search_sum,
            '月购买量总计': month_buy_sum,
            '展示价均值': round(show_price_mean, 2),
            '展示价中位值': round(show_price_median, 2),
            '成交价均值': round(clinch_price_mean, 2),
            '总销售额': round(sales_sum, 2),
            '单个广告订单理想成单费用均值': round(ad_order_price_mean, 2),
            '广告订单:自然订单为1:1时,要求最单最低利润': round(ad_order_price_mean / 2, 2),
            '广告订单:自然订单为1:2时,要求最单最低利润': round(ad_order_price_mean / 3, 2)
        }
        return pd.Series(dec_dic)

    def range(self, field, *args, num=0.00, time=False, number_of_digits=2):
        if len(args) != 0:
            bins = list(*args)
        else:
            bins = list()
            bins.append(self.data[field].describe().loc['min', ] - num)
            bins.append(self.data[field].describe().loc['25%', ])
            bins.append(self.data[field].describe().loc['50%', ])
            bins.append(self.data[field].describe().loc['75%', ])
            bins.append(self.data[field].describe().loc['max', ] + num)
            if type(num) is int:
                bins = list(map(int, bins))
            else:
                bins = list(map(lambda x: round(x, number_of_digits), bins))

        if time:
            # 序列化时间
            self.data[field] = pd.to_datetime(self.data[field])
            # 获取当前时间
            current_time = datetime.datetime.now()
            self.data[field + '区间'] = pd.cut((current_time - self.data[field]).dt.days, bins)
        else:
            self.data[field + '区间'] = pd.cut(self.data[field], bins)
        return bins[1], bins[3]

    def working_pivot_tables(self, field, sort=False):
        """
        生成数据透视表
        :param field:
        :param sort:
        :return:
        """
        new_data = pd.pivot_table(self.data, index=field, values=['月销量'],
                                  aggfunc=['count', 'sum', 'mean', 'median', 'max', 'min'])

        # 修改透视表的列索引
        new_data.columns = ['链接数', '月销量', '平均月销量', '月销量中位值', '最大月销量', '最小月销量']
        if sort:
            new_data.sort_values(by=['月销量', '链接数'], ascending=False, inplace=True)
        new_data.loc[:, '链接数占比'] = (new_data['链接数'] / new_data['链接数'].sum()).\
            apply(lambda x: format(x, ".2%"))
        new_data.loc[:, '月销量占比'] = (new_data['月销量'] / new_data['月销量'].sum()).\
            apply(lambda x: format(x, ".2%"))

        # 修改索引名称
        new_data.index = new_data.index.astype(str).str.replace('(', field + '(', regex=False)
        new_data.index = new_data.index.str.replace('-1', '0', regex=False)
        new_data.index = new_data.index.str.replace('100000000', '+∞', regex=False)

        # 将销量相关数据规范为整数型
        new_data.fillna(0, inplace=True)
        new_data.iloc[:, :6] = new_data.iloc[:, :6].astype(int)

        # 返回处理好的透视表
        return new_data

    def save_data(self, **kwargs):
        # 获取当前时间获得时间戳，定义输出文件的文件名，保证文件名的唯一性
        file_name_suffix = datetime.datetime.now()
        file_name_suffix = str(file_name_suffix.year) + str(file_name_suffix.month) + str(file_name_suffix.day) + \
                           str(file_name_suffix.hour) + str(file_name_suffix.minute) + str(file_name_suffix.second)
        path = self.path_list[0] + '分析文档' + file_name_suffix + '.xlsx'
        # 将源数据写入excel文件
        write_file_name = pd.ExcelWriter(path)
        for kwa in kwargs:
            kwargs[kwa]['data'].to_excel(write_file_name, sheet_name=kwargs[kwa]['sheet_name'],
                                         startcol=kwargs[kwa]['start_col'], startrow=kwargs[kwa]['start_row'],
                                         header=kwargs[kwa]['header'], index=kwargs[kwa]['index'])
        write_file_name.save()
        tkinter.messagebox.showinfo('提示', '数据分析已完成，文件存放路径为：' + path)

    @staticmethod
    def top_ana_sum(data):
        # 计算概述统计值
        listing_num = data['Link'].count()
        moth_sale_total = data['月销量'].sum()
        moth_sale_mean = data['月销量'].mean()
        price_mean1 = (data['月销量'] * data['售价']).sum() / data['月销量'].sum()
        price_mean2 = data['售价'].mean()
        star_mean = data['星级'].loc[data['星级'] > 0].mean()
        review_numer = data['评分数'].sum()
        dec_dic = {'Listing数量': listing_num,
                   '总销量为': moth_sale_total,
                   '平均月销量': int(round(moth_sale_mean, 0)),
                   '平均成交价': round(price_mean1, 2),
                   '平均展示价': round(price_mean2, 2),
                   '平均星级': round(star_mean, 1),
                   '总评论数': review_numer}
        return pd.Series(dec_dic)


if __name__ == '__main__':
    surface = Ui()
    surface.lab('↓↓↓按需求点击下方按钮提交数据源↓↓↓', '#03230e', 60)
    surface.lab('', 'white', 5)
    surface.button('>>Top分析点这里<<', lambda: ButtonMethod().top_ana())
    surface.lab('', 'white', 1)
    surface.button('>>卖家精灵关键词广告分析点这里<<', lambda: ButtonMethod().keywords_ana())
    surface.lab('', 'white', 1)
    surface.button('>>JS关键词获取<<', js_keyword_ana)
    surface.window.mainloop()
