#import four excel into dataframe
import pandas as pd

df1 = pd.read_excel("1734_history.xlsm",index_col = "Date")
df2 = pd.read_excel("4105_history.xlsm",index_col = "Date")
df3 = pd.read_excel("4142_history.xlsm",index_col = "Date")
df4 = pd.read_excel("6547_history.xlsm",index_col = "Date")

plt1 = df1["Close"].plot()
plt1.set_xlabel("Date")
plt1.set_ylabel("Price")
plt1.set_title("1734 Trending")

plt2 = df2["Close"].plot()
plt2.set_xlabel("Date")
plt2.set_ylabel("Price")
plt2.set_title("4105 Trending")

plt3 = df3["Close"].plot()
plt3.set_xlabel("Date")
plt3.set_ylabel("Price")
plt3.set_title("4142 Trending")

plt4 = df4["Close"].plot()
plt4.set_xlabel("Date")
plt4.set_ylabel("Price")
plt4.set_title("6547 Trending")

##計算 1.從2019/1~2019/12期間投資股票 （買低賣高）的總獲利
######2.從2019/1~2021/1期間投資股票 （同樣買低賣高）的總獲利
######總獲利2-總獲利1= 疫情開始後可獲利金額

##2019/1~2020/1
df11 = df1.loc["2019-12-31":"2019-01-01"]

df21 = df2.loc["2019-12-31":"2019-01-01"]

df31 = df3.loc["2019-12-31":"2019-01-01"]

df41 = df4.loc["2019-12-31":"2019-01-01"]

### First do 2019/1~2019/12 profit

import numpy as np
df11["max20d"] = df11["Close"].rolling(20).max().shift(1)
df11["min20d"] = df11["Close"].rolling(20).min().shift(1)
df11["mean20d"] = ((df11["max20d"]+df11["min20d"])/2).shift(1)
df11["signal"] = ""
df11["profit"] = np.nan

df21["max20d"] = df21["Close"].rolling(20).max().shift(1)
df21["min20d"] = df21["Close"].rolling(20).min().shift(1)
df21["mean20d"] = ((df21["max20d"]+df21["min20d"])/2).shift(1)
df21["signal"] = ""
df21["profit"] = np.nan

df31["max20d"] = df31["Close"].rolling(20).max().shift(1)
df31["min20d"] = df31["Close"].rolling(20).min().shift(1)
df31["mean20d"] = ((df31["max20d"]+df31["min20d"])/2).shift(1)
df31["signal"] = ""
df31["profit"] = np.nan

df41["max20d"] = df41["Close"].rolling(20).max().shift(1)
df41["min20d"] = df41["Close"].rolling(20).min().shift(1)
df41["mean20d"] = ((df41["max20d"]+df41["min20d"])/2).shift(1)
df41["signal"] = ""
df41["profit"] = np.nan

##since there are four stocks, write a function to apply four on
def stock_signal_detect(date):
    signal = False
    buy = False
    sell = False
    balance = 0
    share = 1000
    
    for index, row in date.iterrows():
        if signal == False:
            if row["Close"]<=row["min20d"]:
                date.loc[index,"signal"]="buy"
                balance-=row["Close"]*share
                date.loc[index,"profit"] = balance
                signal = True
                buy = True
            elif row["Close"]>=row["max20d"]:
                date.loc[index,"signal"]="sell"
                balance+=row["Close"]*share
                date.loc[index,"profit"] = balance
                signal = True
                sell = True
            else:
                date.loc[index,"signal"] = "---"
                date.loc[index,"profit"] = balance
        elif signal == True:
            if row["Close"]>=row["mean20d"] and buy == True:
                date.loc[index,"signal"]="Offset"
                balance+=row["Close"]*share
                date.loc[index,"profit"]= balance
                signal = False
                buy = False
                balance = 0
            elif row["Close"]<=row["mean20d"] and sell == True:
                date.loc[index,"signal"]="Offset"
                balance-=row["Close"]*share
                date.loc[index,"profit"]=balance
                signal=False
                sell=False
                balance=0
            else:
                date.loc[index,"signal"]="---"
                date.loc[index,"profit"]= balance

##從2019-1-1到2019-12-31買4支股票的總獲利
stock_signal_detect(df11)
result11 = df11[df11["signal"]=="Offset"]
profit_11 =result11["profit"][-1]

stock_signal_detect(df21)
result21 = df21[df21["signal"]=="Offset"]
profit_21 =result21["profit"][-1]

stock_signal_detect(df31)
result31 = df31[df31["signal"]=="Offset"]
profit_31 =result31["profit"][-1]

stock_signal_detect(df41)
result41 = df41[df41["signal"]=="Offset"]
profit_41 =result41["profit"][-1]

profit_11,profit_21,profit_31,profit_41

profit_before_covid = profit_11+profit_31+profit_41
profit_before_covid

import numpy as np
df1["max20d"] = df1["Close"].rolling(20).max().shift(1)
df1["min20d"] = df1["Close"].rolling(20).min().shift(1)
df1["mean20d"] = ((df1["max20d"]+df1["min20d"])/2).shift(1)
df1["signal"] = ""
df1["profit"] = np.nan

df2["max20d"] = df2["Close"].rolling(20).max().shift(1)
df2["min20d"] = df2["Close"].rolling(20).min().shift(1)
df2["mean20d"] = ((df2["max20d"]+df2["min20d"])/2).shift(1)
df2["signal"] = ""
df2["profit"] = np.nan

df3["max20d"] = df31["Close"].rolling(20).max().shift(1)
df3["min20d"] = df31["Close"].rolling(20).min().shift(1)
df3["mean20d"] = ((df31["max20d"]+df11["min20d"])/2).shift(1)
df3["signal"] = ""
df3["profit"] = np.nan

df4["max20d"] = df4["Close"].rolling(20).max().shift(1)
df4["min20d"] = df4["Close"].rolling(20).min().shift(1)
df4["mean20d"] = ((df4["max20d"]+df4["min20d"])/2).shift(1)
df4["signal"] = ""
df4["profit"] = np.nan

##從2019-1-1到2021-1-1買4支股票的總獲利
stock_signal_detect(df1)
result1 = df1[df1["signal"]=="Offset"]
profit_1 =result1["profit"][-1]

stock_signal_detect(df2)
result2 = df2[df2["signal"]=="Offset"]
profit_2 =result2["profit"][-1]

stock_signal_detect(df3)
result3 = df3[df3["signal"]=="Offset"]
profit_3 =result3["profit"][-1]

stock_signal_detect(df4)
result4 = df4[df4["signal"]=="Offset"]
profit_4 =result4["profit"][-1]

profit_1,profit_2,profit_3,profit_4

profit_after_covid=profit_1+profit_3+profit_4
profit_after_covid

#疫情開始後3支股票相較於疫情開始前的獲利
#東洋未納入計算，因為趨勢圖上並未在疫情後有明顯起伏
profit_after_covid-profit_before_covid
