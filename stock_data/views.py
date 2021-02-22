#import openpyxl
#from django.core.files.storage import FileSystemStorage
from django.shortcuts import render

# Create your views here.
import pandas as pd
import numpy as np
import datetime as dt





def xyz(myFile):
    df = pd.read_excel(myFile,index_col='Datetime')
    #df = pd.read_excel('./QQQ_OHLC_Data_Hourly.xlsx', index_col='Datetime')
    set_df = setdf(df)
    return getresult(set_df)


def setdf(df):
    straddle_check = dt.time(16, 0)
    df['seven_hour_ma'] = df['Close'].rolling(7).mean()
    df['twenty_hour_ma'] = df['Close'].rolling(20).mean()
    df['return'] = df['Close'].pct_change()
    df['four_pm_straddle'] = np.where((df['Time'] == straddle_check) & (df['High'] > df['twenty_hour_ma']) & (df['Low'] < df['twenty_hour_ma']), 0, 1)
    df['buy_signal'] = np.where(df['seven_hour_ma'] > df['seven_hour_ma'].shift(), 1, 0)
    df['portfolio_return'] =df['return'].shift(-1) * df['buy_signal'] * df['four_pm_straddle']
    df['mid'] = (df['High'] + df['Low']) / 2
    df['seven_hour_ma_mid'] = df['mid'].rolling(7).mean()
    df['twenty_hour_ma_mid'] = df['mid'].rolling(20).mean()
    df['return_mid'] = df['mid'].pct_change()
    df['four_pm_straddle_mid'] = np.where((df['Time'] == straddle_check) & (df['High'] > df['twenty_hour_ma_mid']) & (
            df['Low'] < df['twenty_hour_ma_mid']), 0, 1)
    df['buy_signal_mid'] = np.where(df['seven_hour_ma_mid'] > df['seven_hour_ma_mid'].shift(), 1, 0)
    df['portfolio_return_mid'] = df['return_mid'].shift(-1) * df['buy_signal_mid'] * df['four_pm_straddle_mid']
    return df


def getresult(df):
    result = []
    buy_and_hold = np.cumprod(1 + df['return'])
    strategy = np.cumprod(1 + df['portfolio_return'])
    buy_and_hold_monthly_value = buy_and_hold.resample("M").last()
    strategy_monthly_value = strategy.resample("M").last()

    buy_and_hold_monthly_return = (buy_and_hold_monthly_value / buy_and_hold_monthly_value.shift()) - 1
    strategy_monthly_return = (strategy_monthly_value / strategy_monthly_value.shift()) - 1

    buy_and_hold_sharpe = (buy_and_hold_monthly_return.mean() / buy_and_hold_monthly_return.std()) * np.sqrt(12)
    strategy_sharpe = (strategy_monthly_return.mean() / strategy_monthly_return.std()) * np.sqrt(12)
    buy_and_hold_mid = np.cumprod(1 + df['return_mid'])
    strategy_mid = np.cumprod(1 + df['portfolio_return_mid'])
    buy_and_hold_monthly_value_mid = buy_and_hold_mid.resample("M").last()
    strategy_monthly_value_mid = strategy_mid.resample("M").last()

    buy_and_hold_monthly_return_mid = (buy_and_hold_monthly_value_mid / buy_and_hold_monthly_value_mid.shift()) - 1
    strategy_monthly_return_mid = (strategy_monthly_value_mid / strategy_monthly_value_mid.shift()) - 1

    buy_and_hold_sharpe_mid = (
                                          buy_and_hold_monthly_return_mid.mean() / buy_and_hold_monthly_return_mid.std()) * np.sqrt(
        12)
    strategy_sharpe_mid = (strategy_monthly_return_mid.mean() / strategy_monthly_return_mid.std()) * np.sqrt(12)
    result.append(buy_and_hold_monthly_return.mean())
    result.append(strategy_monthly_return.mean())
    result.append(buy_and_hold_sharpe)
    result.append(strategy_sharpe)
    result.append(buy_and_hold_monthly_return_mid.mean())
    result.append(strategy_monthly_return_mid.mean())
    result.append(buy_and_hold_sharpe_mid)
    result.append(strategy_sharpe_mid)
    return result


def home(request):
    if request.method == 'GET':
        result = None
    else:
        excel_file = request.FILES["excel_file"]
        result = xyz(excel_file)
    return render(request, 'Home.html', {'result': result})
