import csv

# 打开 CSV 文件
with open('courses.csv', encoding='utf-8-sig') as csvfile:
    reader = csv.reader(csvfile)
    
    # 读取第一行，通常是列名
    headers = next(reader)
    print("实际读取的列名：", headers)  # 查看实际读取的列名
    
    # 读取第二行
    second_row = next(reader)
    print("第二行的数据：", second_row)  # 查看第二行的数据
