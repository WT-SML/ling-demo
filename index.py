import time
import pandas as pd

# data_a_path = './data/上市公司发明申请专利细分分类号.xlsx' # 真实数据
data_a_path = './data/上市公司发明申请专利细分分类号 - 测试（删减后的小量数据）.xlsx'  # 测试小批量数据
data_b_path = './data/IPC代码.xlsx'  # 目标IPC代码 （数字技术创新类）
output_path = './result/result.xlsx'  # 输出文件路径

print(f"正在读取dataA...")
ts = time.time()
data_a = pd.read_excel(data_a_path)
print(f"读取dataA完毕，耗时{round(time.time() - ts, 2)}秒")

print(f"正在读取dataB Sheet3...")
ts = time.time()
data_b_sheet3 = pd.read_excel(data_b_path, sheet_name='Sheet3')
print(f"读取dataB Sheet3完毕，耗时{round(time.time() - ts, 2)}秒")

target_ipc_list: list[str] = list(data_b_sheet3['IPC代码'])  # 目标IPC代码 （数字技术创新类）

# 处理IPC代码
for i in range(len(target_ipc_list)):
    target_ipc = target_ipc_list[i]
    # 如果IPC代码包含*
    if '*' in target_ipc:
        # 如果长度为5，直接删掉*
        if len(target_ipc) == 5:
            target_ipc_list[i] = target_ipc.replace('*', '')
        else:
            # 否则将*替换成/
            target_ipc_list[i] = target_ipc.replace('*', '/')

print(f"目标IPC代码：{target_ipc_list}")

# 遍历dataA的行
print(f"正在遍历dataA...")
ts = time.time()
output_data = pd.DataFrame(columns=['股票代码', '会计年度', '公司类型', '申请时间', '数字经济专利申请'])
for i, row in data_a.iterrows():
    # 第一行是中文翻译，不是数据，跳过
    if i == 0:
        continue
    count = 0
    # 继续遍历行数据
    for cell in row:
        # 行数据里的内容，按大括号分组成专利列表
        patent_list = str(cell).split(',')
        for patent in patent_list:
            for target_ipc in target_ipc_list:
                # 如果这个专利包含目标IPC代码，则count计数+1
                if target_ipc in patent:
                    count += 1
                    break
    new_row = {'股票代码': row['Scode'], '会计年度': row['Year'], '公司类型': row['Ftyp'], '申请时间': row['Aplctm'],
               '数字经济专利申请': count}
    output_data = output_data._append(new_row, ignore_index=True)
print(f"遍历dataA完毕，耗时{round(time.time() - ts, 2)}秒")

print(f"正在输出统计结果Excel...")
ts = time.time()
output_data.to_excel(output_path, index=False)
print(f"输出统计结果Excel完毕，耗时{round(time.time() - ts, 2)}秒")
