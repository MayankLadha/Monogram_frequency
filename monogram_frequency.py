###########################
# Written by Mayank Ladha #
###########################

import xlsxwriter

workbook = xlsxwriter.Workbook('Monogram_Frequency.xlsx')
worksheet = workbook.add_worksheet()
bold = workbook.add_format({'bold': 1})

fo = open("sample2.txt", "r")
b = True
alpha = list('abcdefghijklmnopqrstuvwxyz')
freq = {}
total = 0

for x in alpha:
    freq[x] = 0;

while b:
    hs = fo.read(100)
    if len(hs) != 0:
        for x in alpha:
            freq[x] += hs.count(x)
    else:
        b = False

print ("letter : frequency")

row=1
col=0
worksheet.set_row(0, 20, bold)
worksheet.set_row(27, 20, bold)
worksheet.set_column('A:A', 20)
worksheet.set_column('B:B', 20)
worksheet.write(0,0,"Letter")
worksheet.write(0,1,"Frequency")

for x in alpha:
    col=0
    print (x+"      : "+str(freq[x]))
    worksheet.write(row,col,x)
    col+=1
    worksheet.write(row,col,freq[x])
    row+=1

for x in alpha:
    total += freq[x]

worksheet.write(row,0,"Total Character Count")
worksheet.write(row,1,total)

sorted_freq = sorted(freq.items(), key=lambda x: x[1], reverse=True)

row=1
col=5
worksheet.set_column('F:F', 20)
worksheet.set_column('G:G', 20)
worksheet.write(0,5,"Letter")
worksheet.write(0,6,"Frequency")
for x in range(0,26):
    col=5
    worksheet.write(row,col,sorted_freq[x][0])
    col+=1
    worksheet.write(row,col,sorted_freq[x][1])
    row+=1

worksheet.write(row,5,"Total Character Count")
worksheet.write(row,6,total)

print ("\nTotal Character Count: "+str(total))

workbook.close()

print ("Report successfully writtern to Single_Letter_Frequency.xlsx")
