import pandas as pd

SHEET = "Game.xlsx"

df = pd.read_excel(SHEET, header=2, usecols = "B:U")
information = {}

i = 0
for column in df.columns[::2]:
    data = list(df.iloc[:, [i]].values.flatten())
    price = list(df.iloc[:, [i + 1]].values.flatten())
    information[column] = {"data": data, "price": price}
    i += 2

def get_lowest_highest_price():
    min_max = {}
    for index in information:
        maximum = max(information[index]["price"])
        minimum = min(information[index]["price"])
        maximumIndex = information[index]["price"].index(maximum)
        minimumIndex = information[index]["price"].index(minimum)
        min_max[index] = {"Minimum": "{} held by {}".format(minimum, information[index]["data"][minimumIndex]), "Maximum": "{} held by {}".format(maximum, information[index]["data"][maximumIndex])}
    return min_max

def get_indices(lst, item):
    return [i for i, x in enumerate(lst) if x == item]

def get_prices_sum():
    summ = {}
    for index, data in information.items():
        summ[index] = {i:[] for i in data["data"]}
        for i in data["data"]:
            indices = get_indices(data["data"], i)
            summ[index][i] = [data["price"][_index] for _index in indices]
    
    for index, data in summ.items():
        for _index, _data in data.items():
            summ[index][_index] = sum(_data)
    return summ

def get_top_five():
    top = {}

    for index, data in information.items():
        _sorted = sorted(data["price"], reverse=True)
        top[index] = []

        for i in range(5):
            _index = data["price"].index(_sorted[i])
            top[index].append("{} held by {}".format(data["price"][_index], data["data"][_index])) 
    return top

min_max = get_lowest_highest_price()
summ = get_prices_sum()
top = get_top_five()

df1 = pd.DataFrame(data=min_max).fillna(0)
df2 = pd.DataFrame(data=summ).fillna(0)
df3 = pd.DataFrame(data=top).fillna(0)
df1.name = "Minimum - Maximum"
df2.name = "Sum of prices"
df3.name = "Top 5"

writer = pd.ExcelWriter('test.xlsx',engine='xlsxwriter')
workbook=writer.book
worksheet=workbook.add_worksheet('Result')
writer.sheets['Result'] = worksheet
worksheet.write_string(0, 4, df1.name)
df1.to_excel(writer,sheet_name='Result',startrow=1 , startcol=0)
worksheet.write_string(df1.shape[0] + 4, 5, df3.name)
df3.to_excel(writer,sheet_name='Result',startrow=df1.shape[0] + 5, startcol=0)
worksheet.write_string(df1.shape[0] + df3.shape[0] + 8, 5, df2.name)
df2.to_excel(writer,sheet_name='Result',startrow=df1.shape[0] + df3.shape[0] + 9, startcol=0)
writer.save()