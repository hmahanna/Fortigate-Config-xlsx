import openpyxl
import re
import sys


conf_file = input("Please Enter config file name : ")
conf_file = conf_file + '.txt'
xl_file = input("Please Enter Excel file name : ")
xl_file = xl_file + '.xlsx'

workbook = openpyxl.load_workbook(xl_file)
sheet = workbook.get_sheet_by_name('Sheet1')

print ('###################################################')
print ('###################################################')
print ('Policy export in progress ,Please wait ...........)')
print ('###################################################')
print ('###################################################')


def FW_Object(object_var,sheet_no):
    if words[1] == object_var:
        service = re.search('set %s (.*)' % object_var, line)
        if service:
            sheet[str(sheet_no) + str(n)] = service.group(1)
            workbook.save(xl_file)


n=1
count = 0
dic = {sys.argv[1] : 'A' , sys.argv[2] : 'B' , sys.argv[3] : 'C' , sys.argv[4] : 'D' , sys.argv[5] : 'E' , sys.argv[6] : 'F' , sys.argv[7] : 'G' }


with open (conf_file) as f :
    for line in f:
        count = count + 1


with open (conf_file) as f :
    for i in range(10000):
        for line in f :
            if len(line.split()) == 0:
                continue
            line = line.strip()
            words = line.split()
            if words[0] == 'edit': # New Policy Entry
                n = n + 1          # New Policy Entry
            if words[0] == 'end':
                break
            if words[0] == 'next':
                break
            for item in dic.items():
                FW_Object(item[0],item[1])
            break


print ('done')



