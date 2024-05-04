import pandas as pd
for i in range(2):
    df = pd.read_excel('mt.xlsx',i)
    df = df.reset_index()  # make sure indexes pair with number of rows

    with open(r'a.cls', 'r', encoding='UTF-8') as file:
        data = file.read()
        for index, row in df.iterrows():
            if(isinstance(row[1],int)):
                x='<name>{}{}</name>'.format(row[1]%100,row[2])
                y='<name>192.168.19.{}</name>'.format(row[1]%100)
                data = data.replace(y, x)
        with open(r'{}en.cls'.format(i+1), 'w', encoding='UTF-8') as file:
            file.write(data)
    print("Text replaced")

