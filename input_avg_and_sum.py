import openpyxl

# Excelファイルの読み込み
work_book = openpyxl.load_workbook('./python問題.xlsx')
# ワークシートの取得
sheet = work_book['問題1']

# 指定したセルに値とフォーマットと3桁区切りを設定する
def set_cell(cell_name, value):
    cell = sheet[cell_name]
    # 3桁ごとのカンマ区切り
    format = '#,##0'

    cell.value = value
    cell.number_format = format

# 各列の合計と平均を計算してセルに値を入れる
def calc_column_avg_and_sum(column_index):

    # 範囲用変数 C3:C7など
    column_range = f'{column_index}3:{column_index}7'
    # 範囲内çのセルをタプルで取得する
    cell_range = sheet[column_range]
    # 合計用変数
    sum = 0

    column_values = [cell.value for row_cells in cell_range for cell in row_cells]
    # 以下とほぼ同じ意味
    # # 各行を取り出す
    # for row_cells in cell_range:
    #     # cellを取り出す
    #     for cell in row_cells:
    #         # C3からC7の値を足し合わせる
    #         sum += cell.value
    
    for cell_value in column_values:
        sum += cell_value

    avg = sum / len(column_values)

    set_cell(f'{column_index}8', sum)
    set_cell(f'{column_index}9', avg)

# 各行の合計と平均を計算してセルに値を入れる
def calc_row_avg_and_sum(row_index):
    # 範囲用変数 C3:C7など    
    row_range = f'C{row_index}:G{row_index}'
    # 範囲内çのセルをタプルで取得する
    cell_range = sheet[row_range]
    # 合計用変数
    sum = 0

    row_values = [cell.value for col_cells in cell_range for cell in col_cells]
    # 以下とほぼ同じ意味
    # for col_cells in cell_range:
    #     for cell in col_cells:
    #         print(cell.value)

    for cell_value in row_values:
        sum += cell_value
    
    avg = sum / len(row_values)

    set_cell(f'H{row_index}', sum)
    set_cell(f'I{row_index}', avg)

# 行の範囲
row_ranges = [3, 4, 5, 6, 7]
# 各列を計算してcellに値を入れる
for row_index in row_ranges:
    calc_row_avg_and_sum(row_index)

# 列の範囲
column_ranges = ['C', 'D', 'E', 'F', 'G', 'H', 'I']
# 各列を計算してcellに値を入れる
for column_index in column_ranges:
    calc_column_avg_and_sum(column_index)

# 保存
work_book.save('./python編集.xlsx')
