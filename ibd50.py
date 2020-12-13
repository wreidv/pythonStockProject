import pandas as pd
import yfinance as yf
import datetime as dt
from pandas_datareader import data as pdr
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import csv

root = Tk()
ftypes = [("Excel files", ".xlsx .xls")]
ttl = "Title"
dir1 = 'C:\\'
filePath = askopenfilename(filetypes = ftypes, initialdir = dir1, title = ttl)

df = pd.read_excel(filePath)

yf.pdr_override()

eps = []
rating = []
names = []
rank = []
rsrating = []
emasUsed = [3,5,8,10,12,15,30,35,40,45,50,60]
isYellow = False
isRWB = False

startyear = 2019
startmonth = 1
startday = 1

for x in range(8,58):

    if(pd.isna(df.iloc[x,15]) == False):
        tempEPS = int(df.iloc[x,15])
    else:
        tempEPS = 0

    if(pd.isna(df.iloc[x,9]) == False):
        tempRating = int(df.iloc[x,9])
    else:
        tempRating = 0

    if(pd.isna(df.iloc[x,10]) == False):
        tempRSRating = int(df.iloc[x,10])
    else:
        tempRSRating = 0

#    print("EPS: ",tempEPS, " Rating: ",tempRating," RS: ",tempRSRating)

    if tempEPS > 40:
        if tempRating > 95:
            if tempRSRating > 95:
                rsrating.append(tempRSRating)
                eps.append(tempEPS) #[x,15]
                rating.append(tempRating) #[x,9]
                rank.append(int(df.iloc[x,2])) #[x,2]
                names.append(df.iloc[x,0]) #[x,0]


start = dt.datetime(startyear, startmonth, startday)
now = dt.datetime.now()

with open('ibd50output.csv', mode='w') as output:
    output_writer = csv.writer(output, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
    output_writer.writerow(["Symbol", "Rank", "Composite Rating", "RS Rating", "EPS Change(Latest Qtr)",  "IsYellow?", "Days in row", "RWB?", "Days in row"])


    for x in range(0, len(names)):
        df = pdr.get_data_yahoo(names[x], start, now)

        ma50 = df.iloc[:,4].rolling(window = 50).mean()
        ma150 = df.iloc[:,4].rolling(window = 150).mean()
        numYellow = 0
        numRWB = 0

        for i in emasUsed:
            df["Ema_"+str(i)] = round(df.iloc[:,4].ewm(span=i, adjust = False).mean(), 2)

        for i in df.index:
            cmin = min(df["Ema_3"][i], df["Ema_5"][i],df["Ema_8"][i], df["Ema_10"][i],df["Ema_12"][i], df["Ema_15"][i])
            cmax = max(df["Ema_30"][i], df["Ema_35"][i],df["Ema_40"][i], df["Ema_45"][i],df["Ema_50"][i], df["Ema_60"][i])

            if((df["Adj Close"][i] > cmin) and (cmin > cmax)):
                numRWB += 1
            else:
                numRWB = 0

            if((df["Adj Close"][i] > ma50[i]) and (ma50[i] > ma150[i])):
                numYellow += 1
            else:
                numYellow = 0


        isRWB = (df["Adj Close"][len(df.index)-1] > cmin) and (cmin > cmax)
        isYellow = (df["Adj Close"][len(df.index)-1] > ma50[len(ma50)-1]) and (ma50[len(ma50)-1] > ma150[len(ma150)-1])

        output_writer.writerow([names[x], rank[x], rating[x], rsrating[x], eps[x], isYellow, numYellow, isRWB, numRWB])


print("written to ibd50output.csv")
