#import libraries for specific use functions
import pandas as pd
import numpy as np
import math
from pandas import ExcelWriter
from pandas import ExcelFile

#reads in intitial data from excel file
#includes all sheets or "TABS"
df = pd.read_excel('C:/Users/alyss/Desktop/Spring 2020/Honors Thesis/Ez/EZall.xlsx', sheet_name=None)

#rows and colums count to set initial array size
#make new variables for the other 2 data set groups
rows, cols = (34, 34)

# create new arrays for other data sets using new row and column count
arr = [[0 for i in range(cols)] for j in range(rows)]
arr1 = [[0 for i in range(cols)] for j in range(rows)]
arr2 = [[0 for i in range(cols)] for j in range(rows)]

#math function one
def logit(p):
    return math.log(p/(1-p))
#math function two
def ezdiff(subject_id,MRT,VRT,p, Variable):

    if p == 0:
        print("Oops, only errors for subject " + subject_id + "!")
    elif p == 0.5:
        print("Oops, chance performance for " + subject_id + "!")
    elif p == 1:
        print("Oops,  only correct responses for " + subject_id + "!")
    
    s = 0.1
    s2 = s*s # scaling parameter squared
    L = logit(p)
    x1 = (L*p*p - L*p + p - 0.5)
    x = L * x1 / VRT
    v = np.sign(p - 0.5) * s*math.pow(x,0.25)
    a = s2*logit(p)/v
    y = -v*a/s2
    MDT = (a/(2*v))*(1-math.exp(y))/(1+math.exp(y))
    Ter = MRT - MDT
    if Variable == 0:
        return v
    elif Variable == 1:
        return a
    elif Variable == 2:
        return Ter
    #return([v,a,Ter])
#IGNORE
#print(df['W-MRT'].iloc[20,10]/1000)
#print(df['W-VRT'].iloc[20,10])
#print(df['W-Pc'].iloc[20,10])
#print(ezdiff("0000", df['W-MRT'].iloc[20,10]/1000, df['W-VRT'].iloc[20,10]/1000000, df['W-Pc'].iloc[20,10]))

#loop to itterate through each row
#duplicate for other two data groups
for i in range(0,34):
    #asign id, group, sex, and age independently
    arr[i][0] = df['MRT'].iloc[i,0]
    arr[i][1] = df['MRT'].iloc[i,1]
    arr[i][2] = df['MRT'].iloc[i,2]
    arr[i][3] = df['MRT'].iloc[i,3]

    arr1[i][0] = df['MRT'].iloc[i,0]
    arr1[i][1] = df['MRT'].iloc[i,1]
    arr1[i][2] = df['MRT'].iloc[i,2]
    arr1[i][3] = df['MRT'].iloc[i,3]

    arr2[i][0] = df['MRT'].iloc[i,0]
    arr2[i][1] = df['MRT'].iloc[i,1]
    arr2[i][2] = df['MRT'].iloc[i,2]
    arr2[i][3] = df['MRT'].iloc[i,3]
    #loop to itterate through each column
    #data cells start at column 4(5) in the excel sheet
    for j in range(4,34):
        #perform mathmatical calculations, division is done in method for ease
        arr[i][j] = ezdiff("0000", df['MRT'].iloc[i,j]/1000, df['VRT'].iloc[i,j]/1000000, df['PC'].iloc[i,j], 0)
        arr1[i][j] = ezdiff("0000", df['MRT'].iloc[i,j]/1000, df['VRT'].iloc[i,j]/1000000, df['PC'].iloc[i,j], 1)
        arr2[i][j] = ezdiff("0000", df['MRT'].iloc[i,j]/1000, df['VRT'].iloc[i,j]/1000000, df['PC'].iloc[i,j], 2)
#displays data in easy to read format in consol for data verification
#for row in arr:
#    print(row)
#assigns data array to dataframe for later transfer to excel sheet
#accomplishes this by a shorthand version of a loop for each column
mydf = pd.DataFrame({'Group':[i[0] for i in arr],
                   'Subject':[i[1] for i in arr],
                   'Age':[i[2] for i in arr],
                   'Sex':[i[3] for i in arr],
                   'n90(1)':[i[4] for i in arr],
                   'n45(1)':[i[5] for i in arr],
                   '0(1)':[i[6] for i in arr],
                   'p45(1)':[i[7] for i in arr],
                   'p90(1)':[i[8] for i in arr],
                   'n90(2)':[i[9] for i in arr],
                   'n45(2)':[i[10] for i in arr],
                   '0(2)':[i[11] for i in arr],
                   'p45(2)':[i[12] for i in arr],
                   'p90(2)':[i[13] for i in arr],
                   'n90(3)':[i[14] for i in arr],
                   'n45(3)':[i[15] for i in arr],
                   '0(3)':[i[16] for i in arr],
                   'p45(3)':[i[17] for i in arr],
                   'p90(3)':[i[18] for i in arr],
                   'n90(4)':[i[19] for i in arr],
                   'n45(4)':[i[20] for i in arr],
                   '0(4)':[i[21] for i in arr],
                   'p45(4)':[i[22] for i in arr],
                   'p90(4)':[i[23] for i in arr],
                   'n90(5)':[i[23] for i in arr],
                   'n45(5)':[i[25] for i in arr],
                   '0(5)':[i[26] for i in arr],
                   'p45(5)':[i[27] for i in arr],
                   'p90(5)':[i[28] for i in arr],
                   'n90(6)':[i[29] for i in arr],
                   'n45(6)':[i[30] for i in arr],
                   '0(6)':[i[31] for i in arr],
                   'p45(6)':[i[32] for i in arr],
                   'p90(6)':[i[33] for i in arr]})

mydf1 = pd.DataFrame({'Group':[i[0] for i in arr1],
                   'Subject':[i[1] for i in arr1],
                   'Age':[i[2] for i in arr1],
                   'Sex':[i[3] for i in arr1],
                   'n90(1)':[i[4] for i in arr1],
                   'n45(1)':[i[5] for i in arr1],
                   '0(1)':[i[6] for i in arr1],
                   'p45(1)':[i[7] for i in arr1],
                   'p90(1)':[i[8] for i in arr1],
                   'n90(2)':[i[9] for i in arr1],
                   'n45(2)':[i[10] for i in arr1],
                   '0(2)':[i[11] for i in arr1],
                   'p45(2)':[i[12] for i in arr1],
                   'p90(2)':[i[13] for i in arr1],
                   'n90(3)':[i[14] for i in arr1],
                   'n45(3)':[i[15] for i in arr1],
                   '0(3)':[i[16] for i in arr1],
                   'p45(3)':[i[17] for i in arr1],
                   'p90(3)':[i[18] for i in arr1],
                   'n90(4)':[i[19] for i in arr1],
                   'n45(4)':[i[20] for i in arr1],
                   '0(4)':[i[21] for i in arr1],
                   'p45(4)':[i[22] for i in arr1],
                   'p90(4)':[i[23] for i in arr1],
                   'n90(5)':[i[24] for i in arr1],
                   'n45(5)':[i[25] for i in arr1],
                   '0(5)':[i[26] for i in arr1],
                   'p45(5)':[i[27] for i in arr1],
                   'p90(5)':[i[28] for i in arr1],
                   'n90(6)':[i[29] for i in arr1],
                   'n45(6)':[i[30] for i in arr1],
                   '0(6)':[i[31] for i in arr1],
                   'p45(6)':[i[32] for i in arr1],
                   'p90(6)':[i[33] for i in arr1]})

mydf2 = pd.DataFrame({'Group':[i[0] for i in arr2],
                   'Subject':[i[1] for i in arr2],
                   'Age':[i[2] for i in arr2],
                   'Sex':[i[3] for i in arr2],
                   'n90(1)':[i[4] for i in arr2],
                   'n45(1)':[i[5] for i in arr2],
                   '0(1)':[i[6] for i in arr2],
                   'p45(1)':[i[7] for i in arr2],
                   'p90(1)':[i[8] for i in arr2],
                   'n90(2)':[i[9] for i in arr2],
                   'n45(2)':[i[10] for i in arr2],
                   '0(2)':[i[11] for i in arr2],
                   'p45(2)':[i[12] for i in arr2],
                   'p90(2)':[i[13] for i in arr2],
                   'n90(3)':[i[14] for i in arr2],
                   'n45(3)':[i[15] for i in arr2],
                   '0(3)':[i[16] for i in arr2],
                   'p45(3)':[i[17] for i in arr2],
                   'p90(3)':[i[18] for i in arr2],
                   'n90(4)':[i[19] for i in arr2],
                   'n45(4)':[i[20] for i in arr2],
                   '0(4)':[i[21] for i in arr2],
                   'p45(4)':[i[22] for i in arr2],
                   'p90(4)':[i[23] for i in arr2],
                   'n90(5)':[i[24] for i in arr2],
                   'n45(5)':[i[25] for i in arr2],
                   '0(5)':[i[26] for i in arr2],
                   'p45(5)':[i[27] for i in arr2],
                   'p90(5)':[i[28] for i in arr2],
                   'n90(6)':[i[29] for i in arr2],
                   'n45(6)':[i[30] for i in arr2],
                   '0(6)':[i[31] for i in arr2],
                   'p45(6)':[i[32] for i in arr2],
                   'p90(6)':[i[33] for i in arr2]})

#creates writer object to facilitate excel transfer    
writer = ExcelWriter('EZ WxL.xlsx')
#transfers excel sheet
mydf.to_excel(writer,'V')
mydf1.to_excel(writer,'A')
mydf2.to_excel(writer,'TER')
#saves the actual excell sheet to file location in python
writer.save()
