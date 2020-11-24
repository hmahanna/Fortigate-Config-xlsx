from tqdm import tqdm
from time import sleep
import openpyxl
import os
import sys
import re




conf_file = input("Please Enter config file name : ")
xl_file = input("Please Enter Excel file name : ")

workbook = openpyxl.load_workbook(xl_file)
sheet = workbook.get_sheet_by_name('Sheet1')
print ('#########################################')
print ('Policy export in progress ,Please wait :)')
print ('Policy export in progress ,Please wait :)')
print ('Policy export in progress ,Please wait :)')
print ('Policy export in progress ,Please wait :)')
print ('#########################################')
print ('Policy export in progress ,Please wait :)')
print ('#########################################')


def name_fun():
    if words[1] == 'name':
        name = re.search('set name (".*")', line)
        if name:
            sheet['A' + str(n)] = name.group(1)
            workbook.save(xl_file)


def srcaddr_fun():
    if words[1] == 'srcaddr':
        src = re.search('set srcaddr (".*")', line)
        if src:
            sheet['B' + str(n)] = src.group(1)
            workbook.save(xl_file)



def dstaddr_fun():
    if words[1] == 'dstaddr':
        dst = re.search('set dstaddr (".*")', line)
        if dst:
            sheet['C' + str(n)] = dst.group(1)
            workbook.save(xl_file)


def srcintf_fun():
    if words[1] == 'srcintf':
        src = re.search('set srcintf (".*")', line)
        if src:
            sheet['D' + str(n)] = src.group(1)
            workbook.save(xl_file)            



def dstintf_fun():
    if words[1] == 'dstintf':
        src = re.search('set dstintf (".*")', line)
        if src:
            sheet['E' + str(n)] = src.group(1)
            workbook.save(xl_file)





def service_fun():
    if words[1] == 'service':
        service = re.search('set service (".*")', line)
        if service:
            sheet['F' + str(n)] = service.group(1)
            workbook.save(xl_file)

n=0

with open (conf_file) as f :

    for i in tqdm(range(10000)):
        for line in f :
            if len(line.split()) == 0:
                continue
            line = line.strip()
            words = line.split()

            if words[0] == 'edit' :
                n = n + 1
            if words[0] == 'end' :
                break
            if words[0] == 'next':
                break
            name_fun()
            srcaddr_fun()
            dstaddr_fun()
            service_fun()
            srcintf_fun()
            dstintf_fun()
            break



