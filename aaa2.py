import pandas as pd
from pandas import DataFrame

df = pd.read_excel('stock2.xlsx')

inventory_code = df['code'].tolist()
inventory_name= df['name'].tolist()
inventory_amount = df['amount'].tolist()
inventory_unit = df["birim"].tolist()
inventory_location=df["location"].tolist()

inventory_name_new=[]
inventory_amount_new=[]
inventory_unit_new=[]
inventory_location_new=[]

class Inventory:
    def __init__(self,i_code,i_amount,i_unit,i_location):
        self.i_code=i_code
        self.i_amount=i_amount
        self.i_unit=i_unit
        self.i_location=i_location

    def info(self):
        print("Product code:", self.i_code, "Amount:", self.i_amount, "Unit:", self.i_unit, "Place:", self.i_location)

    def add(self, add_amount):
        self.i_amount += add_amount
    def change_location(self,new_location):
        self.i_location=new_location

    def change_unit(self,new_unit):
        self.i_unit=str(new_unit)

    def update(self):
        for i in range(0, len(inventory_code)):
            if inventory_code[i] == self.i_code:
                inventory_unit_new.insert(i,self.i_unit)
                inventory_amount_new.insert(i,self.i_amount)
                inventory_location_new.insert(i,self.i_location)
            else:
                inventory_unit_new.insert(i,5)
                inventory_amount_new.insert(i,5)
                inventory_location_new.insert(i,5)


        p = zip(inventory_name, inventory_amount_new, inventory_unit_new, inventory_code, inventory_location_new)
        df2 = DataFrame(p)
        df2.columns = ["name", "amount", "birim", "code", "location"]

        writer = pd.ExcelWriter('stock2.xlsx', engine='xlsxwriter')
        df2.to_excel(writer, sheet_name='Sheet1')

        writer.save()


for i in range(0,len(inventory_name)):
    globals()[inventory_name[i]]=Inventory(inventory_code[i],inventory_amount[i],inventory_unit[i],inventory_location[i])

