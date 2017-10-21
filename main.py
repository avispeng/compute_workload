# -*- coding: UTF-8 -*-
import pandas as pd
import numpy as np
import Tkinter as tk
import tkFileDialog as filedialog


FILE_PATH = ""

def get_month(date):
    #drop day and keep year and month
    return date[:7]


def festival(minute_con):
    # workload times 3 on festival
    if minute_con[1] != u'\u5e73\u65f6':
        minute_con[0] *= 3
    return minute_con[0]


def details(record):
    minute = record[8]
    type_of = record[9]
    # details about which class is this record in
    if type_of == u'\u5267\u96c6\u65f6\u95f4\u8f74\u5236\u4f5c':
        record[19] += minute
        record[-1] += (minute * 15)
    elif type_of == u'\u5267\u96c6\u65f6\u95f4\u8f74\u8c03\u6574':
        record[20] += minute
        record[-1] += (minute * 10)
    elif type_of == u'\u7535\u5f71\u65f6\u95f4\u8f74\u5236\u4f5c':
        record[21] += minute
        record[-1] += (minute * 15)
    elif type_of == u'\u7eaa\u5f55\u7247\u65f6\u95f4\u8f74\u5236\u4f5c':
        record[22] += minute
        record[-1] += (minute * 20)
    elif type_of == u'\u771f\u4eba\u79c0\u65f6\u95f4\u8f74\u5236\u4f5c':
        record[23] += minute
        record[-1] += (minute * 25)
    elif type_of == u'\u97e9\u5267\u65f6\u95f4\u8f74\u5236\u4f5c':
        record[24] += minute
        record[-1] += (minute * 20)
    elif type_of == u'\u97e9\u56fd\u7efc\u827a\u65f6\u95f4\u8f74\u5236\u4f5c':
        record[25] += minute
        record[-1] += (minute * 25)
    elif type_of == u'\u65e5\u5267\u65f6\u95f4\u8f74\u5236\u4f5c':
        record[26] += minute
        record[-1] += (minute * 20)
    else:
        record[27] += minute
        record[-1] += (minute * 5)
    return record[19:]


def open_file():
    name = filedialog.askopenfilename(filetypes=(("Office Excel 2007", "*.xls"),("Excel Workbook", "*.xlsx")),
                                      title="Choose an excel file")
    global FILE_PATH
    try:
        with open(name, 'r'):
            FILE_PATH = name
            label_text.set("Opened")
    except:
        label_text.set("You haven't opened any files")


def do_something():
    df = pd.read_excel(FILE_PATH, header=0, index_col=0)
    save_path = FILE_PATH.rsplit('.', 1)[0]

    # select 删除记录 and 修改已登记工作量
    # delete the record they point to
    what_to_delete1 = df[u'\u767b\u8bb0\u7c7b\u578b'] == u'\u5220\u9664\u8bb0\u5f55'
    what_to_delete2 = df[u'\u767b\u8bb0\u7c7b\u578b'] == u'\u4fee\u6539\u5df2\u767b\u8bb0\u5de5\u4f5c\u91cf'
    what_to_delete = df[what_to_delete1 | what_to_delete2]
    for index, row in what_to_delete.iterrows():
        if df.ix[row[u'\u8981\u4fee\u6539\u6216\u5220\u9664\u7684\u5e8f\u53f7'],u'\u5de5\u53f7'] == row[u'\u5de5\u53f7']:
            # check the record at 要修改或删除的序号
            # whether 工号 of current record fits the one to delete
            # meaning you can't delete other's record
            df.drop([row[u'\u8981\u4fee\u6539\u6216\u5220\u9664\u7684\u5e8f\u53f7']], inplace=True)

    # delete 删除记录 record themselves
    df = df[df[u'\u767b\u8bb0\u7c7b\u578b'] != u'\u5220\u9664\u8bb0\u5f55']
    # extract year and month only
    df.ix[:, 1] = df.ix[:, 1].apply(get_month, 1)

    # workload times 3 if on festival
    df.ix[:, 8] = df.ix[:, [8, 10]].apply(festival, 1)

    # workload in different types
    s_len = len(df.index)
    series_edit = pd.Series(np.zeros(s_len), index=df.index)
    series_adjust = pd.Series(np.zeros(s_len), index=df.index)
    movie = pd.Series(np.zeros(s_len), index=df.index)
    doc = pd.Series(np.zeros(s_len), index=df.index)
    show = pd.Series(np.zeros(s_len), index=df.index)
    korean_series = pd.Series(np.zeros(s_len), index=df.index)
    korean_show = pd.Series(np.zeros(s_len), index=df.index)
    japanese_series = pd.Series(np.zeros(s_len), index=df.index)
    other = pd.Series(np.zeros(s_len), index=df.index)
    total = pd.Series(np.zeros(s_len), index=df.index)
    df = df.assign(series_edit=series_edit).assign(series_adjust=series_adjust).assign(movie=movie).assign(doc=doc) \
        .assign(show=show).assign(korean_series=korean_series).assign(korean_show=korean_show) \
        .assign(japanese_series=japanese_series).assign(other=other).assign(total=total)

    df.ix[:, 19:] = df.ix[:, :].apply(details, 1)

    df.rename(columns={"series_edit": u'\u5267\u96c6\u65f6\u95f4\u8f74\u5236\u4f5c',
                       "series_adjust": u'\u5267\u96c6\u65f6\u95f4\u8f74\u8c03\u6574',
                       "movie": u'\u7535\u5f71\u65f6\u95f4\u8f74\u5236\u4f5c',
                       "doc": u'\u7eaa\u5f55\u7247\u65f6\u95f4\u8f74\u5236\u4f5c',
                       "show": u'\u771f\u4eba\u79c0\u65f6\u95f4\u8f74\u5236\u4f5c',
                       "korean_series": u'\u97e9\u5267\u65f6\u95f4\u8f74\u5236\u4f5c',
                       "korean_show": u'\u97e9\u56fd\u7efc\u827a\u65f6\u95f4\u8f74\u5236\u4f5c',
                       "japanese_series": u'\u65e5\u5267\u65f6\u95f4\u8f74\u5236\u4f5c',
                       "other": u'\u5176\u4ed6',
                       "total": u'\u5de5\u8d44\u5355'}, inplace=True)

    # group lines with same year_month, name, number, ID
    dfgroup = df.groupby([u'\u65e5\u671f', u'\u6635\u79f0', u'\u5de5\u53f7', u'\u8bba\u575b' + 'ID'])

    # compute sum of minutes in each group
    minute_sum = dfgroup[u'\u5206\u949f\u6570',
                         u'\u5267\u96c6\u65f6\u95f4\u8f74\u5236\u4f5c',
                         u'\u5267\u96c6\u65f6\u95f4\u8f74\u8c03\u6574',
                         u'\u7535\u5f71\u65f6\u95f4\u8f74\u5236\u4f5c',
                         u'\u7eaa\u5f55\u7247\u65f6\u95f4\u8f74\u5236\u4f5c',
                         u'\u771f\u4eba\u79c0\u65f6\u95f4\u8f74\u5236\u4f5c',
                         u'\u97e9\u5267\u65f6\u95f4\u8f74\u5236\u4f5c',
                         u'\u97e9\u56fd\u7efc\u827a\u65f6\u95f4\u8f74\u5236\u4f5c',
                         u'\u65e5\u5267\u65f6\u95f4\u8f74\u5236\u4f5c',
                         u'\u5176\u4ed6',
                         u'\u5de5\u8d44\u5355'].agg('sum').reset_index()

    # display by descending order of minutes
    minute_sum.sort_values(u'\u5206\u949f\u6570', ascending=False, axis=0, inplace=True)
    minute_sum.sort_values(u'\u65e5\u671f', ascending=True, axis=0, inplace=True, kind='mergesort')

    minute_sum.to_excel(save_path + '_output.xls', 'sheet1', header='minutes', index=False)
    label_text.set("Ouput successfully!")


if __name__=='__main__':
    window = tk.Tk()
    window.title('Compute workload')
    window.geometry('500x300')

    label_text = tk.StringVar()
    label = tk.Label(window,
                     textvariable=label_text,
                     font=('Times New Roman',12), width=30, height=2)
    label_text.set("Please upload an excel file.")
    label.pack() # fix its position


    button = tk.Button(window,
                       text='Browse',
                       width=20, height=2,
                       command=open_file) # the command to execute when click
    button.pack()

    button2 = tk.Button(window,
                        text='Process',
                        width=20, height=2,
                        command=do_something)
    button2.pack()


    window.mainloop()

