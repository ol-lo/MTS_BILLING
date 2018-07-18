# -*- coding: utf-8 -*-
import xlrd

from datetime import datetime

from constants import SOURCE_FILE, NAME_PAGE
from model import (
    db, HistoryItem, Phone, Subdivision, Region, Legal, Person, Contract)

book_excel = xlrd.open_workbook(SOURCE_FILE)
data_sheet = book_excel.sheet_by_name(NAME_PAGE)
massive_data_from_excel = data_sheet.col_values

region = set(massive_data_from_excel(0)[1:])

legal_subdivision = map(
    lambda x: x.replace(' ', '').split(':'),
    massive_data_from_excel(1)[1:]
)
legal = set(
    map(lambda x: x[0],
        legal_subdivision)
)

subdivision = set(
    map(lambda x: len(x) == 2 and x[1] or '',
        legal_subdivision)
)

telephone = set(massive_data_from_excel(2)[1:])
contract = set(massive_data_from_excel(3)[1:])
name = set(massive_data_from_excel(4)[1:])

# c 5 - 11

db.connect()
db.create_tables(
    [Person, Legal, Region, Subdivision, Phone, Contract,  HistoryItem])
db.commit()


def create_inter(iter_, obj):
    for i in iter_:
        query = obj.select().where(obj.name == i)
        if query:
            continue
            
        if i:
            print i
            v = obj(name=i)
            v.save()


create_inter(region, Region)
print 'region'
create_inter(legal, Legal)
print 'legal'
create_inter(subdivision, Subdivision)
print 'subdivision'
create_inter(telephone, Phone)
print 'telephone'
create_inter(contract, Contract)
print 'contract'
create_inter(name, Person)
print 'name'

months = data_sheet.row_values(0)[6:12]
for month in months:
    first_day_of_month = "1 {}".format(month)
    datetime.strptime(first_day_of_month, '%d %B %Y').date()
    