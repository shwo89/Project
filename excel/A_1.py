# -*- coding: utf-8 -*-
import xlwings as xw
import pandas as pd
import PySimpleGUI as sg
import math
pd.set_option('display.max_rows',100) #打印显示100行
pd.set_option('display.max_columns',10) #打印显示100列

#界面程序,输入损耗量
def InputWastage():
    cancel_btn = sg.Button('Cancel')
    layout = [[sg.Text("绝缘子和耐张线夹损耗"), sg.InputText(default_text='2',size=(5, 1),key='wastage1'),sg.Text("%")],
              [sg.Text("其他金具损耗"), sg.InputText(default_text='1',size=(5, 1),key='wastage2'), sg.Text("%")],
              [sg.Button('开始计算'), cancel_btn]]

    # Create the Window
    window = sg.Window('损耗量设置（%）', layout)

    # Create the event loop
    while True:
        event, values = window.read()
        if event in (None, 'Cancel'):
            # User closed the Window or hit the Cancel button
            break
        #将字典中的val改为float数据类型
        for key in values:
            values[key] = float(values[key])/100
        return values
    window.close()

#获取合并单元格相关数据
def Get_merge():
    text = wb.macro('test') #调用VBA中的自定义函数test，获取合并单元格位置及行数
    p = text()
    p = p.replace('$','\t') #删除p中的'$'符号

    words = []  # 建立一个空列表
    index = 0  # 遍历所有的字符
    start = 0  # 记录每个单词的开始位置
    for i in range(len(p)): #将p中的各元素提取出来放入列表words中
        if p[index] == "\t":
            words.append(p[start:index])
            index += 1
            start = index  # start来记录位置
        elif i == len(p)-1:
            index += 1

            words.append(p[start:index])
        else: index += 1

    i = 0
    while i < len(words): #去除列表'words'中所有空元素
        if words[i] == '':
            del words[i]
        else: i += 1

    j = -1
    while words[j].isalpha() == False:
        j = j-1
    j = int((len(words)+j+2)/2)  #计算合并单元格的数量

    columns = []
    rows = []
    high = []
    wide = []
    for i in range(j): #建合并单元格所在的行、列，已经单元格的高、宽提取出来
        columns.append(words[2*i])
        rows.append(int(words[2*i+1]))
        high.append(int(words[i+2*j]))
        wide.append(int(words[i + 3 * j]))

    dict = {'列':columns,'行':rows,'高':high,'宽':wide}
    df = pd.DataFrame(data=dict)
    df = df.sort_values(by='行').reset_index(drop=True) #按“行”的值升序排列，并重设索引
    df["行"] = df["行"]-1
    df["高"] = df["行"].add(df["高"])-1
    return df

#将excel工作簿内容写入pandas
def Get_xls():
    sheet = xw.sheets.active
    info = sheet.used_range
    nrows = info.last_cell.row
    ncols = info.last_cell.column

    sheet.range((4, 10), (nrows, ncols)).clear()  # 删除现有材料汇总，包括格式和内容

    # 将excel工作簿内容写入pandas
    xls = sheet.range((1, 1), (nrows, 8))
    header = sheet.range((4, 1), (4, 8))
    df = pd.DataFrame(xls.value, columns=header.value)  # 列索引为写入的excel表头
    df = df.drop(['每串数量', '每基串数', '铁塔数量'], axis=1)  # 删除'每串数量','每基串数','铁塔数量'列
    return df

#非合并单元格处理
def Group(df,merge):
    for i in merge.index.tolist():
        df = df.drop(labels=range(merge.at[i, '行'], merge.at[i, '高']))  # 删除df中合并单元格所在行，labels为按行索引号选择

    # 删除第'设计用量'列中不包含数字的行，删除不必要的列
    df = df.dropna(subset=['设计用量'])  # 删除‘设计用量’列中的none行
    df = df[df['设计用量'].str.isdigit().isnull()]  # 删除‘设计用量’列中的非数字行
    df = df.fillna("-")  # 将None填充"-"

    duplicate_row = df.duplicated(subset=['名称','型号'],keep=False) #找出'名称','型号'这2列数据完全相同的行
    duplicate_data = df.loc[duplicate_row,:] #取出重复行数据，并重行索引
    duplicate_data_sum = duplicate_data.groupby(['名称','型号']).agg({'设计用量':'sum','序号':'sum','单位':'min'}).reset_index()
    duplicate_data_sum = duplicate_data_sum.loc[:,['序号','名称','型号','单位','设计用量']] #使用切片将列重新排序
    no_duplicate = df.drop_duplicates(subset=['名称','型号'] ,keep=False)#获取不重复的数据，指定列subset=['名称','型号']，不保留重复数据：keep=False

    result = pd.concat([duplicate_data_sum,no_duplicate]).reset_index(drop=True)#拼接”新重复值中的一个”和不重复的数据，并重置行索引

    result["序号"] = result.index.get_level_values(0).values + 1  # 使用行索引+1填充“序号”

    return result

#合并单元格处理
def Merge(df,merge):
    g_list = {} #创建字典
    for i in merge.index.tolist():
        g_list[i] = df.loc[merge.at[i,'行']:merge.at[i,'高']].reset_index(drop=True) #根据合并单元格数量创立对应的二维数据组,df.loc为按行索引号选择

    #两两比较g_list中的二维数组，相同合并，并清除后面那个二维数组的数据
    for i in range(merge.shape[0]-1):
        for j in range(i+1,merge.shape[0]):
            if g_list[i]["型号"].equals(g_list[j]["型号"]) and g_list[i]["名称"].equals(g_list[j]["名称"]):
                g_list[i]["设计用量"] = g_list[i]["设计用量"].add(g_list[j]["设计用量"])
                g_list[j].drop(g_list[j].index, inplace=True)


    result = pd.DataFrame(columns=('序号','名称','型号','单位','设计用量'))
    #把有数值的g_list二维数据，竖向拼接放入result中
    for key in g_list:
        if len(g_list[key].index) != 0:
            g_list[key]["型号"] = g_list[key].loc[0,'型号']
            result = pd.concat([result, g_list[key]], axis=0)

    return result

#计算损耗量
def Wastage_Count(result,wastage1,wastage2):
    #“损耗量”=“设计用量”*损耗率(wastage2)，并向上取整
    result["损耗量"] = result["设计用量"].mul(wastage2).apply(lambda x: math.ceil(x))

    #如果'名称'中包含'绝缘子'或‘耐张线夹’字眼，则损耗率为wastage1
    result["损耗量"] = result["损耗量"].mask(result['名称'].str.contains('绝缘子|耐张线夹'),result["设计用量"].mul(wastage1).apply(lambda x: math.ceil(x)))

    #如果'名称'中包含‘耐张线夹’字眼，则损耗率量+6
    result["损耗量"] = result["损耗量"].mask(result['名称'].str.contains('耐张线夹'),result["损耗量"].add(6))

    #如果'名称'中包含‘接续管’字眼，则损耗率量+3
    result["损耗量"] = result["损耗量"].mask(result['名称'].str.contains('接续管'),result["损耗量"].add(3))

    result["总量"] = result["设计用量"].add(result["损耗量"])
    result["备注"] = None

    #如果'名称'中包含‘耐张线夹’字眼，则‘备注’增加：已考虑试验用6个
    result["备注"] = result["备注"].mask(result['名称'].str.contains('耐张线夹'),"已考虑试验用6个")

    #如果'名称'中包含‘接续管’字眼，则‘备注’增加：已考虑试验用3个
    result["备注"] = result["备注"].mask(result['名称'].str.contains('接续管'),"已考虑试验用3个")

    return result

#引用当前工作簿
app = xw.apps.active
wb = app.books.active
sheets = xw.sheets
sheet = xw.sheets.active


df = Get_xls()      #将表格读取进df
merge = Get_merge() #获取表格中的合并单元格信息

df1 = Group(df,merge)
df2 = Merge(df,merge)

# #将df1'名称'列和‘型号’列中包含"预绞"字眼的行取出放到df2之后
df = df1[df1['名称'].str.contains("预绞|绞丝") | df1['型号'].str.contains("预绞|绞丝")]
df2 = pd.concat([df2,df], axis=0).reset_index(drop=True)
df2 = df2.reset_index(drop=True) #重置行索引
df2["序号"] = df2.index.get_level_values(0).values + 1  # 使用行索引+1填充“序号”

# 删除'名称'列和‘型号’列中包含"预绞"字眼的行
df1 = df1[~df1['名称'].str.contains("预绞|绞丝") & ~df1['型号'].str.contains("预绞|绞丝")]
df1 = df1.reset_index(drop=True) #重置行索引
df1["序号"] = df1.index.get_level_values(0).values + 1  # 使用行索引+1填充“序号”

wastage = InputWastage() #输入损耗量

#计算损耗量
df1 = Wastage_Count(df1,wastage['wastage1'],wastage['wastage2'])
df2 = Wastage_Count(df2,wastage['wastage1'],wastage['wastage2'])

sheet.range(4,10).value = df1  # 将df1输出到excel
sheet.range(df1.shape[0]+5,11).value = '三、预绞丝金具'
sheet.range(df1.shape[0]+6,10).value = df2  # 将df2输出到excel

#调整单元格格式
rng = sheet.range((4,11),(df1.shape[0]+df2.shape[0]+6,18))
for border_id in range(7, 13): #设定边框
    rng.api.Borders(border_id).LineStyle = 1

sheet.range((df1.shape[0]+5,11),(df1.shape[0]+5,18)).api.Merge() #合并单元格
rng.api.HorizontalAlignment = -4108 #水平居中
rng.api.VerticalAlignment = -4130 #垂直居中
sheet.range((df1.shape[0]+5,11),(df1.shape[0]+5,18)).api.HorizontalAlignment = -4131 #水平靠左
sheet.range((df1.shape[0]+5,11),(df1.shape[0]+5,18)).api.Font.Bold = True #加粗





# Border_Set(rng)
sheet.range('J:J').clear_contents()  # 删除行索引







