import pandas as pd

fio = []
_date = []
_guid = []
data = pd.read_csv("Move.csv", sep=";", names= list(range(0,20)))

data = data[(data[0] !='Маршрут ХО или/и посетителей')]
data = data[(data[0] !='© ЗАО НВП Болид" 2010г."')]

df_date = data[0].dropna(axis=0, how = "all").to_list()
data = data[(data[0] !='дата/время')]
_fio = data[2].dropna(axis=0, how = "all").to_list()

guid = data[[2, 12]]
guid = guid[(guid[2] == "фамилия")]
guid = guid[12].to_list()

data = data.drop([1,2,4,9, 10, 11, 12], axis=1)
data = data.dropna(how="all", axis=1)
data = data.dropna(how = "all", axis=0)


df = pd.DataFrame(data)
df = df.rename(columns={0:"Дата", 3:"Статус", 7:"Место", 15:"Пояснение"})

date_count = 0
date_count_last = 0
for date in df_date[1:]:
    if date != "дата/время":
        date_count += 1
    else: 
        _date.append(date_count)
        date_count = 0
    date_count_last = date_count
_date.append(date_count_last)

        
get_num_date = 0
for index, item in enumerate(_fio):

    if item == "фамилия":

        surname =  _fio[index+5] if  _fio[index+5] != "фамилия" else ""
        iter_fio = _fio[index+1] + " " + _fio[index+3] + " " + surname
        for item in range(_date[get_num_date]):
            fio.append(iter_fio) 
        get_num_date +=1  
#print(len(fio))

get_num_guid = 0
for gui in guid:

    for item in range(_date[get_num_guid]):
        _guid.append(gui)
    get_num_guid += 1
        
df["ФИО"] = fio
df["guid"] = _guid

first_column = df.pop("ФИО")
df.insert(0, "ФИО", first_column)
df.to_excel("пасматри.xlsx", index = False)

