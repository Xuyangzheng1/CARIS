
#车险续保自动化报表v3
#已完成文件分发20240103，20230103
#firstpart已完成
#表1：2024年第一季度分车种续保率通报表完成
#表2：2024年一季度各省分车种续保率通报表《====SecondPart_all
#员工车续保情况未完成，未完成toExcel
import os
import pandas as pd
import xlrd
import chardet
import streamlit as st
from tabulate import tabulate
import re
import matplotlib.pyplot as plt
from abc import ABC, abstractmethod
import pandas as pd

def modify_index_kindOfCar(index):
    replacements = {
        'A家庭自用车': '家庭自用车',
        'B非营业客车': '非营业客车',
        'C营业客车': '营业客车',
        'F特种车': '特种车'
    }
    return replacements.get(index, index) 

def check_trend(value):
    value = float(value.strip('%'))
    if value < 0:
        return(f"下降{-value}")
    elif value > 0:
        return(f"上升{value}")
    else:
        return("保持不变")

def FirstPart_all(DataFrame):
    
    data1 = DataFrame.loc[('中原农业保险股份有限公司', '汇总'), '应续保单']
    data2 = DataFrame.loc[('中原农业保险股份有限公司', '汇总'), '已续保单']
    data3 = DataFrame.loc[('中原农业保险股份有限公司', '汇总'), '保单续保率']
    data4 = DataFrame.loc[('中原农业保险股份有限公司', '汇总'), '较去年同期']

    check = check_trend(data4)
    # 家用车续保率
    data5 = DataFrame.loc[('中原农业保险股份有限公司', '家庭自用车'), '保单续保率']
    # 非营运客车续保率
    data6 = DataFrame.loc[('中原农业保险股份有限公司', '非营业客车'), '保单续保率']
    # 营运客车续保率
    data7 = DataFrame.loc[('中原农业保险股份有限公司', '营业客车'), '保单续保率']
    # 特种车
    data8 = DataFrame.loc[('中原农业保险股份有限公司', '特种车'), '保单续保率']

    print(f"2024年一季度车险应续保单件数{data1}件，已续保单件数{data2}件，保单续保率{data3}，较去年同期{check}个百分点。"
          f"其中家用车续保率{data5}，非营运客车续保率{data6}；营运客车续保率{data7}，特种车续保率{data8}。")
    result_string = f"2024年一季度车险应续保单件数{data1}件，已续保单件数{data2}件，保单续保率{data3}，较去年同期{check}个百分点。" \
                    f"其中家用车续保率{data5}，非营运客车续保率{data6}；营运客车续保率{data7}，特种车续保率{data8}。"
    return result_string


#总数据索引重命名
def modify_index(index):
    if index == '02150000000-内蒙古分公司':
        return '内蒙古'
    elif index == '02230000000-黑龙江分公司':
        return '黑龙江'
    elif index == '02410000000-河南省分公司':
        return '河南省'
    elif index =='02000000000-中原农业保险股份有限公司':
        return '中原农业保险股份有限公司'
    return index

def FirstTable(DataFrame):
    FirstTable= DataFrame.loc['中原农业保险股份有限公司']
    return FirstTable

def SecondTable(df):#第二部分表格
    df = df.drop('中原农业保险股份有限公司', level='省份')
    return df

def SecondChart(dfs,dfPaths):
    columns = ['一月', '二月', '三月']

# 定义行索引
    index = ['河南省', '内蒙古', '黑龙江']

# 创建一个空的DataFrame
    Table = pd.DataFrame(index=index, columns=columns)
    
    for df,dfp in zip(dfs,dfPaths):
        df=all_data_month(df)
       
        if dfp.endswith('202401.csv'):
            
            Table.at['河南省', '一月']=df.loc[('河南省', '家庭自用车'), '保单续保率']
            Table.at['黑龙江', '一月']=df.loc[('黑龙江', '家庭自用车'), '保单续保率']
            Table.at['内蒙古', '一月']=df.loc[('内蒙古', '家庭自用车'), '保单续保率']
            # st.dataframe(df)
            # st.warning(dfp)
            print('SecondChart使用数据:',dfp)
        elif dfp.endswith('202402.csv'):
            
            Table.at['河南省', '二月']=df.loc[('河南省', '家庭自用车'), '保单续保率']
            Table.at['黑龙江', '二月']=df.loc[('黑龙江', '家庭自用车'), '保单续保率']
            Table.at['内蒙古', '二月']=df.loc[('内蒙古', '家庭自用车'), '保单续保率']
            # st.dataframe(df)
            # st.warning(dfp)
            print('SecondChart使用数据:',dfp)
        elif dfp.endswith('202403.csv'):
            
            Table.at['河南省', '三月']=df.loc[('河南省', '家庭自用车'), '保单续保率']
            Table.at['黑龙江', '三月']=df.loc[('黑龙江', '家庭自用车'), '保单续保率']
            Table.at['内蒙古', '三月']=df.loc[('内蒙古', '家庭自用车'), '保单续保率']
            # st.dataframe(df)
            # st.warning(dfp)
            print('SecondChart使用数据:',dfp)
    return Table

def HandOutYear(files):
    #分发20240103.csv和20230103.csv《============
    for file in files:
        with open(file, 'rb') as f:
            result = chardet.detect(f.read())
            # st.subheader(file)
        if file.endswith('20240103.csv'):
            dataframeNewPath=file
            
            dataframeNew = pd.read_csv(file,encoding=result['encoding'],index_col=[0,1],header=0)
        elif file.endswith('20230103.csv'):
            dataframeOldPath = file
            dataframeOld = pd.read_csv(file,encoding=result['encoding'],index_col=[0,1],header=0)
    
    return dataframeNew, dataframeNewPath, dataframeOld, dataframeOldPath

def HandOutMonth(files):
    #分发20240103.csv和20230103.csv《============

    files_2023 = []
    files_2024 = []
    dataframes_2023 = []
    dataframes_2024 = []

    for file in files:
        with open(file, 'rb') as f:
            result = chardet.detect(f.read())
            
        if file.endswith('202401.csv') or file.endswith('202402.csv') or file.endswith('202403.csv'):
            files_2024.append(file)
            
            dataframes = pd.read_csv(file,encoding=result['encoding'],index_col=[0,1],header=0)
            dataframes_2024.append(dataframes)
        # elif file.endswith('20230103.csv'):
        #     dataframeOldPath = file
        #     dataframeOld = pd.read_csv(file,encoding=result['encoding'],index_col=[0,1],header=0)

    return files_2023, files_2024, dataframes_2023, dataframes_2024

def find_files(folder_path):
    """
    找到指定文件夹中的所有 Excel 和 CSV 文件
    """
    files = []
    for root, _, filenames in os.walk(folder_path):
        for filename in filenames:
            if filename.lower().endswith(('.xls', '.xlsx', '.csv')):
                files.append(os.path.join(root, filename))
    return files

def read_file(file):
    """
    读取 Excel 或 CSV 文件并将其转换为 Pandas DataFrame
    """
    with open(file, 'rb') as f:
        result = chardet.detect(f.read())
    if file.lower().endswith('.xlsx'):
        df = pd.read_excel(file, engine='openpyxl')
    elif file.lower().endswith(('.xls', '.csv')):
        df = pd.read_csv(file, encoding=result['encoding'], index_col=[0,1], header=0)
    return df

def all_data(files):
    dataframeNew, dataframeNewPath, dataframeOld, dataframeOldPath = HandOutYear(files)

    # 检查新旧 DataFrame 列名，确保所需列存在
    print(dataframeNew.columns)
    print(dataframeOld.columns)

    # 创建 Table DataFrame，初始化数据
    All_data_Table = pd.DataFrame({
        '应续保单': dataframeNew['应续保单件数'],
        '已续保单': dataframeNew['已续保单件数'],
        '保单续保率': dataframeNew['累计保单续保率(%)'].apply(lambda x: "{:.2f}%".format(x)),
        '较去年同期': (dataframeNew['累计保单续保率(%)'] - dataframeOld['累计保单续保率(%)']).apply(lambda x: "{:.2f}%".format(x)),
    })

    # 假设 '车辆种类' 和 '机构' 已经是多级索引的一部分
    # 如果不是，您需要先将这些列设置为索引
    if '车辆种类' in dataframeNew.index.names and '机构' in dataframeNew.index.names:
        All_data_Table.index = dataframeNew.index  # 直接使用新 DataFrame 的多级索引
    else:
        print("车辆种类或机构不在索引中，需要调整 DataFrame 结构")

    # 重命名索引名称以符合新的要求
    
    
    All_data_Table.index.names = ['省份', '车种']
    
   
    All_data_Table.index = All_data_Table.index.set_levels(All_data_Table.index.levels[0].map(modify_index), level=0)#省份去除数字

    All_data_Table.index = All_data_Table.index.set_levels(All_data_Table.index.levels[1].map(modify_index_kindOfCar), level=1)

    return All_data_Table
        
def all_data_month(dataframeNew):
    

    # 创建 Table DataFrame，初始化数据
    All_data_Table = pd.DataFrame({
        '应续保单': dataframeNew['应续保单件数'],
        '已续保单': dataframeNew['已续保单件数'],
        '保单续保率': dataframeNew['累计保单续保率(%)'].apply(lambda x: "{:.2f}%".format(x)),
        # '较去年同期': (dataframeNew['累计保单续保率(%)'] - dataframeOld['累计保单续保率(%)']).apply(lambda x: "{:.2f}%".format(x)),
    })

    # 假设 '车辆种类' 和 '机构' 已经是多级索引的一部分
    # 如果不是，您需要先将这些列设置为索引
    if '车辆种类' in dataframeNew.index.names and '机构' in dataframeNew.index.names:
        All_data_Table.index = dataframeNew.index  # 直接使用新 DataFrame 的多级索引
    else:
        print("车辆种类或机构不在索引中，需要调整 DataFrame 结构")

    # 重命名索引名称以符合新的要求
    
    
    All_data_Table.index.names = ['省份', '车种']
    
   
    All_data_Table.index = All_data_Table.index.set_levels(All_data_Table.index.levels[0].map(modify_index), level=0)#省份去除数字

    All_data_Table.index = All_data_Table.index.set_levels(All_data_Table.index.levels[1].map(modify_index_kindOfCar), level=1)

    return All_data_Table
    
    
def main():
    folder_path = r'E:\2024A4A1TEST\车险续保自动化报表'  # 指定文件夹路径
    files = find_files(folder_path)  # 获取文件列表
# 总数据《-------------------------------01
    result = all_data(files)  

    # print(result)
    # st.dataframe(result) 
# 总数据《-------------------------------01
    # st.markdown("---")
#第一张表<-------------------------------02

    firstTable=FirstTable(result)
    # st.header('2024年第一季度分车种续保率通报表')
    # st.dataframe(firstTable)
#第一张表<-------------------------------02

#第一段概括文本<-------------------------03
    FirstTxt=FirstPart_all(result)
    # st.warning(FirstTxt)
#第一段概括文本<-------------------------03

#第二部分2024年一季度各省分车种续保率通报表<-------------------04
    # st.header('2024年第一季度省分车种续保率通报表')
    SeTable=SecondTable(result)
    # st.dataframe(SeTable) 


#第二部分2024年一季度各省分车种续保率通报表<-------------------04

#第二部分二小节2024年一季度各省分车种续保率通报表<-------------------05
    files_2023, files_2024, dataframes_2023, dataframes_2024=HandOutMonth(files)

    print(dataframes_2024)
    print(files_2024)
    SecondTable1=SecondChart(dataframes_2024,files_2024)
    # st.dataframe(SecondTable1)
#第二部分二小节2024年一季度各省分车种续保率通报表<-------------------05
    

#st
# 在边栏中添加内容
    with st.sidebar:
        st.header("边栏区域")
        st.write("这里是边栏的内容...")
        # 边栏中可以添加更多控件
        if st.button("点击我"):
            st.sidebar.write("您点击了按钮！")

    # 使用 columns 在主内容区创建视觉上的均衡
    col1, col2 = st.columns(2)

    with col1:
        st.header("原始数据")

        st.write("原始数据")
        st.dataframe(result)
        st.markdown('---')
    with col2:
        st.header("计算结果")
         
        st.write('2024年第一季度分车种续保率通报表')
        st.dataframe(firstTable)
        st.markdown('---')

        st.write('第一段概括文本')
        st.warning(FirstTxt)
        st.markdown('---')

        st.write('2024年第一季度省分车种续保率通报表')
        st.dataframe(SeTable) 
        st.markdown('---')

        st.write('第二部分二小节2024年一季度各省分车种续保率通报表')
        st.dataframe(SecondTable1)
        st.markdown('---')
   
    
    
    
if __name__ == '__main__':
    main()
