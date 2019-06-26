import xlrd
import re
from docx import Document

def salary_peocessing(salary_temp):
    count = 0
    for i in range(salary_temp.__len__()):
        if (salary_temp[i] == ''):
            salary_temp[i] = 'NULL'



    for i in range(salary_temp.__len__()):
        st = str(salary_temp[i])

        if (st != 'NULL'):
            st = st.upper()
            st = st.replace('RS', '')
            st = st.replace(',', '')
            st = st.replace('STIPEND', '')
            st = st.replace('(', '')
            st = st.replace(')', '')
            st = st.replace('UG', 'UG ')
            st = st.replace('PG', 'PG ')
            st = st.replace('MCA', '')
            st = st.replace('/', '')
            st = st.replace('\\', '')
            st = st.replace('K', '000')
            st = st.replace(' K', '000')
            x = st.split(" ")

            st2 = st.replace(' ','')
            st2 = st2.replace('.', '')

            while '' in x:
                x.remove('')




            if 'UG' in x or 'PG' in x:
                if 'LPA' in x:
                    for k3 in range(x.__len__()):
                        if (x[k3] == 'UG'):
                            while (1):
                                if(k3 < x.__len__()-1):
                                    k3 += 1
                                    if (x[k3] == 'LPA'):
                                        st = x[k3 - 1]
                                        break
                                else:
                                    break


                    for k3 in range(x.__len__()):
                        if (x[k3] == 'PG'):
                            while (1):
                                if (k3 < x.__len__() -1):
                                    k3 += 1
                                    if (x[k3] == 'LPA'):
                                        st = st + '/' + x[k3 - 1]
                                        break
                                else:
                                    break


                else:
                    if 'PM' in x:
                        for k3 in range(x.__len__()):
                            if (x[k3] == 'UG'):
                                while (1):
                                    if (k3 < x.__len__()-1):
                                        k3 += 1
                                        if (x[k3] == 'PM'):
                                            st = x[k3 - 1]
                                            num = int(st)
                                            num = num * 12
                                            num = num / 100000
                                            st = str(round(num, 1))
                                            break
                                    else:
                                        break

                        for k3 in range(x.__len__()):
                            if (x[k3] == 'PG'):
                                while (1):
                                    if (k3 < x.__len__()-1):
                                        k3 += 1
                                        if (x[k3] == 'PM'):
                                            st1 = x[k3 - 1]
                                            num = int(st1)
                                            num = num * 12
                                            num = num / 100000
                                            st1 = str(round(num, 1))
                                            st = st + '/' + st1

                                            break

                                    else:
                                        break

                pass
            else:
                if 'LPA' in x:
                    for k1 in range(x.__len__()):
                        if (x[k1] == 'LPA'):
                            st = x[k1 - 1]
                else:
                    if 'PM' in x:
                        for k2 in range(x.__len__()):
                            if (x[k2] == 'PM'):
                                st = x[k2 - 1]
                                num = int(st)
                                num = num * 12
                                num = num / 100000
                                st = str(round(num, 1))

            salary_temp[i] = st
        else:
            st3 = str(comp_info[i]).upper()
            if not st3.__contains__('TOTAL'):
                sl = 'enter salary for comapny '+ str(comp_info[i])
                sal2 = float(input(sl))
                salary_temp[i] = sal2


    return salary_temp


def salary_check(salary_temp):
    for i in range(salary_temp.__len__()):
        s = str(salary_temp[i])
        s = s.replace('.', '')
        s = s.replace('/', '')
        if (((re.search('\D', s)) and s != '0'and s!=0 and s != 'NULL') or s.__len__() > 8):
            print(salary_temp[i])
            # print('It conatins non-readable string please convert it to standard form')
            # print('ex: if string has \'UG 25000 pm and PG 20000 PM\' then calculate LPA and insrt \'3/2.4\' ')
            # print('In above example 3 represent UG salry in LPA and 2.4 represent PG salary in LPA ')
            print('Insert the proper salary in LPA............')
            st = input()
            salary_temp[i] = st

    return salary_temp

stname="2018/1.xlsx"
wb = xlrd.open_workbook(str(stname))
sheet = wb.sheet_by_index(0)

ind_comp = []
temp = []

for i in range(sheet.nrows):
    for j in range(sheet.ncols):
        if (sheet.cell_value(i, j) == "Name of the Company"):
            # print(sheet.cell_value(i,j))
            temp.append(i)
            temp.append(j)
            ind_comp.append(temp)
            temp = []

        if (sheet.cell_value(i, j) == "Total Registered"):
            end_row = i

start_row = ind_comp[0][0] + 1

cse_ind = 0
date_ind = 0
salary_ind = 0
total_ind = 0
mcse_ind = 0
mcne_ind = 0
ise_ind=0
ece_ind = 0
tce_ind = 0
eie_ind = 0
mle_ind = 0
me_ind = 0
iem_ind = 0
che_ind = 0
cve_ind = 0
bte_ind = 0
mca_ind = 0
mba_ind = 0

for i in range(sheet.ncols):
    str1 = sheet.cell_value(start_row, i)
    # print(str1)
    str1 = str(str1)
    str1 = str1.upper()

    if (str1.__contains__("M.TECH-CSE")):
        mcse_ind = i
    if (str1.__contains__("CSE") and (not str1.__contains__("TECH"))):
        cse_ind = i
    if (str1.__contains__("TOTAL")):
        total_ind = i
    if (str1.__contains__("SALARY")):
        salary_ind = i
    if (str1.__contains__("DATE")):
        date_ind = i
#print(ind_comp)

comp_info = []
ind_info=[]
salary_info = []
date_info = []

for i in range(start_row, end_row + 1):  # for rows
    str1 = sheet.cell_value(i, 1)
    str1 = str(str1).upper()
    if ((not str1.__contains__("COMPANY")) and str1 != ''):
        num_ind=[]
        comp_info.append(sheet.cell_value(i, 1))
        for j in range(2,date_ind+1):
            val=sheet.cell_value(i,j)
            val=str(val).upper()
            if(val != 'NA' and j<mcse_ind-1):
                num_ind.append(j-1)
            if(j == date_ind):
                date_info.append(val)
            if(j == salary_ind):
                salary_info.append(val)
        if(len(num_ind) == 0):
            ind_info.append([0])
        else:
            ind_info.append(num_ind)
print(comp_info)
print(salary_info)
print(ind_info)
print(date_info)

print(len(comp_info))
print(len(salary_info))
print(len(ind_info))
print(len(date_info))

salary_info = salary_peocessing(salary_info)
salary_info = salary_check(salary_info)

fp=open('comp_ind.txt','w')
# branch = ['cs','is','ec','ee','ei','im','ch','bt','te','cv','ml','at','me']
for k in range(0,len(comp_info)):
    indd = ""
    fp.write(comp_info[k])
    fp.write('\n')
    fp.write(str(salary_info[k]))
    fp.write('\n')
    fp.write(date_info[k])
    fp.write('\n')
    for j in range(len(ind_info[k])):
        k1 = str(ind_info[k][j])
        d = {'0':'0','1':'1','2':'2','3':'3','4':'4','5':'9','6':'5','7':'11','8':'12','9':'6','10':'7','11':'10','12':'8','13':'13'}
        if(j==len(ind_info[k])-1):
            indd = indd + d[k1]
        else:
            indd = indd + d[k1] + ','
    fp.write(indd)
    fp.write('\n')


fp.close()