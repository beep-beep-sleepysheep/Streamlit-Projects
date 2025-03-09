# streamlit run "D:/lectures/python/final project/project draft/FINALS/Visualization with streamlit.py"
# pip install alpha_vantage
# pip install stocknews
# conda install spyder-terminal -c spyder-ide
from alpha_vantage.fundamentaldata import FundamentalData
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from sklearn.linear_model import LinearRegression
from sklearn.model_selection import train_test_split   
import yfinance as yf
import streamlit as st
import plotly.express as px
from stocknews import StockNews
import os
import win32com.client 
import pythoncom
#%%

# Get current working directory
current_directory = os.getcwd()
print(current_directory)
# Open new terminal in current working directory
# streamlit run "D:/lectures/python/final project/project draft/FINALS/Visualization with streamlit.py"
#%%
#create sidebars on streamlit
tickers = st.sidebar.text_input('tickers')
ticker_list = [ticker.strip() for ticker in tickers.split(',')]

Start_date = st.sidebar.date_input('Start_date')
End_date = st.sidebar.date_input('End_date')
pricing_data_df = {} 
#%%
# Define input widgets and tabs using Streamlit
# Input widgets for ticker, start date, end date
pricing_data, momentum_strategy, OLS_price_prediction, combined, simulated_account, email_attachments, news, fundamental_data = st.tabs(["Pricing Data", "Momentum Strategy", "OLS Price Prediction", "Combined", "Simulated Account", "Email & Attachments", "Top 10 News", "Fundamental Data"])
#https://docs.streamlit.io/library/api-reference/layout/st.tabs
for ticker in ticker_list:
    data = yf.download(ticker, start=Start_date, end=End_date)
    if 'Close' in data.columns:
            fig = px.line(data, x=data.index, y=data['Close'], title=f"{ticker}-Close price, {Start_date}-{End_date}")
            st.plotly_chart(fig)
            # Storing processed data for each stock
            data2 = data
            data2['%change'] =data['Close']/data['Close'].shift(1)-1
            pricing_data_df[ticker] = data2
    else:
        st.warning(f"No 'Close' data available for {ticker}")
#%%
# pricing data
# Fetch and display pricing data for each ticker
with pricing_data:
    st.header('Price Movements')
# Display price movements and calculate annual return & standard deviation
    for ticker, data2 in pricing_data_df.items():
        st.subheader(f'Price Movements From Yahoo for {ticker} from {Start_date} to {End_date}')
        st.write(data2)
        annual_return=data2['%change'].mean()*252*100 #252 Trading days a year
        st.write('Annual Return is ',annual_return,'%')
        standard_deviation=np.std(data2['%change'])*np.sqrt(252)
        st.write('Standard Deviation is', standard_deviation*100,'%')
#%%
#momentum strategy
with momentum_strategy:
    st.header('Momentum Strategy')
    st.write("Step 1: Data Processing and Signal Generation")
#Process our data by defnition
    def momentum_processed_data(ticker_list):
        momentum_processed_data = {}
        for ticker in ticker_list:
            data = yf.download(ticker)
            day = np.arange(1, len(data) + 1)
            data['day'] = day
            #close always equals to adjusted close and we dont need volume
            data.drop(columns=['Adj Close', 'Volume'], inplace=True) 
            data = data[['day', 'Open', 'High', 'Low', 'Close']]
#add rolling average, 5days as short term and 30days for long term
            data['5-day'] = data['Close'].rolling(5).mean().shift()
            data['30-day'] = data['Close'].rolling(30).mean().shift()
#add empty column for signal, if short term > long term, we generate buy signal
            data['signal'] = 0
            for i in range(len(data)):
                if data['5-day'].iloc[i] > data['30-day'].iloc[i]:
                    data.loc[data.index[i], 'signal'] = 1
                elif data['5-day'].iloc[i] < data['30-day'].iloc[i]:
                    data.loc[data.index[i], 'signal'] = -1
                else:
                    data.loc[data.index[i], 'signal'] = 0
#drop rows with empty elements
            data.dropna(inplace=True)
            data['return'] = np.log(data['Close']) - np.log(data['Close'].shift(1))
            data['system_return'] = data['signal'] * data['return']
            data['entry'] = data['signal'] - data['signal'].shift(1)
            data['entry'].fillna(0, inplace=True)
            momentum_processed_data[ticker] = data

        return momentum_processed_data
#%%
#call the function
    momentum_data = momentum_processed_data(ticker_list)
#%%
#generte coresponding plots for each stock ,Graph 1
    for stock, df in momentum_data.items():
        st.subheader(f"{stock} Momentum Data")
        st.dataframe(df)

        st.write("Step 2: Plotting")
        plt.plot(df.iloc[-252:]['Close'], label='Close')
        plt.plot(df.iloc[-252:]['5-day'], label='5-day')
        plt.plot(df.iloc[-252:]['30-day'], label='30-day')
        plt.plot(df[-252:].loc[df.entry == 2].index, df[-252:]['5-day'][df.entry == 2], '^', color='g', label='Buy')
        plt.plot(df[-252:].loc[df.entry == -2].index, df[-252:]['30-day'][df.entry == -2], 'v', label='Sell', color='r')
        plt.legend()
        plt.title(f"{stock} - Momentum Strategy Chart 1")
        current_directory = os.getcwd()
        save_path = os.path.join(current_directory, f"{stock} - Momentum Strategy Chart 1.png")
        print(save_path)
        plt.savefig(save_path)
        st.pyplot()

        st.write("Step 3: Strategy Summary 1")
        st.write("Total number of trades:", len(df[df['entry'] != 0]))
        st.write("Total return:", df['system_return'].sum())
#%%
#OLS price prediction tab            
with OLS_price_prediction:
    #this part of data is same as previous dataframe
    for ticker in ticker_list:
        data = yf.download(ticker, start=Start_date, end=End_date)#import data from yahoo finance for each stock
        day = np.arange(1, len(data) + 1)#from the day 1 to today
        data['day'] = day
        st.header('Regression Part Of'+ ' ' +ticker)
        st.write("Step 1: Data Processing and Signal Generation")
        #Create rolling average for short term and long term momentum indicator
        #rolling method from pandas documentation https://pandas.pydata.org/pandas-docs/stable/user_guide/window.html#rolling-apply
        data['5-day'] = data['Close'].rolling(5).mean()  
        data['30-day'] = data['Close'].rolling(30).mean()
        data.dropna(inplace=True)
        x = data[['Open', 'High', 'Low', 'Close', '5-day', '30-day']]  
        y = data['Close']
        #we use linear regression as our model
        model = LinearRegression()
        x_train = x.iloc[:-126]
        x_test = x.iloc[-126:]
        y_train = y.iloc[:-126]
        y_test = y.iloc[-126:]
        model.fit(x_train, y_train)
        train_predicted_close = model.predict(x_train)
        test_predicted_close = model.predict(x_test)
        data['predicted_close'] = 0
        data.loc[x_train.index, 'predicted_close'] = train_predicted_close
        data.loc[x_test.index, 'predicted_close'] = test_predicted_close
#%%        
#generate signals for this model, if predicted stock is bigger than the day before, we generate buy signal
        data['signal_sklearn'] = 0
        for i in range(1, len(data)):
            if data['predicted_close'].iloc[i-1] < data['predicted_close'].iloc[i]:
                data['signal_sklearn'].iloc[i] = 1
            else:
                data['signal_sklearn'].iloc[i] = -1

        data['entry_ols'] = (data['signal_sklearn'] - data['signal_sklearn'].shift(1))
        data['entry_ols'].fillna(0, inplace=True)
#display dataframe on streamlit
        st.dataframe(data)
#%%        
#generate coresponding plots for each stock ,Graph 2
        st.write("Step 2: Plotting")
        plt.plot(data.iloc[-252:]['Close'], label='Close')
        plt.plot(data.iloc[-252:]['5-day'], label='5-day')
        plt.plot(data.iloc[-252:]['30-day'], label='30-day')
        plt.plot(data.iloc[-252:]['predicted_close'], label='Predicted Close')
        plt.plot(data[-252:].loc[data.entry_ols == 2].index, data[-252:]['predicted_close'][data.entry_ols == 2], '^', color='b',label='Buy')
        plt.plot(data[-252:].loc[data.entry_ols == -2].index, data[-252:]['predicted_close'][data.entry_ols == -2], 'v', color='orange',label='Sell')
        plt.legend()
        plt.title(f"{ticker} - Scikit-Learn Signals Chart 2")
        plt.savefig(os.path.join(current_directory, f"{ticker} - Scikit-Learn Signals Chart 2.png"))
#display on streamlit
        st.pyplot()
#generate summary
        st.write("Step 3: Strategy Summary")
        st.write("Total number of trades:", len(data[data['entry_ols'] != 0]))
#%%
#Combined tab        
with combined:
    st.header("Final Plotting")
#these is the combination of our code for previous 2 parts, just repeat it again as this is another individual tab
    for ticker in ticker_list:
        data = yf.download(ticker, start=Start_date, end=End_date)
        data['day'] = day
        data = data[['day', 'Open', 'High', 'Low', 'Close']]
        data['5-day'] = data['Close'].rolling(5).mean().shift()
        data['30-day'] = data['Close'].rolling(30).mean().shift() 
        data['signal'] = 0
        for i in range(len(data)):
            if data['5-day'].iloc[i] > data['30-day'].iloc[i]:
                data.loc[data.index[i], 'signal'] = 1
            elif data['5-day'].iloc[i] < data['30-day'].iloc[i]:
                data.loc[data.index[i], 'signal'] = -1
            else:
                data.loc[data.index[i], 'signal'] = 0
        data.dropna(inplace=True)
        data['return'] = np.log(data['Close']) - np.log(data['Close'].shift(1))
        data['system_return'] = data['signal'] * data['return']
        data['entry'] = data['signal'] - data['signal'].shift(1)
        data['entry'].fillna(0, inplace=True)
        x = data[['Open', 'High', 'Low', 'Close', '5-day', '30-day']]  
        y = data['Close']
        model = LinearRegression()
        x_train = x.iloc[:-126]
        x_test = x.iloc[-126:]
        y_train = y.iloc[:-126]
        y_test = y.iloc[-126:]
        model.fit(x_train, y_train)
        train_predicted_close = model.predict(x_train)
        test_predicted_close = model.predict(x_test)
        data['predicted_close'] = 0
        data.loc[x_train.index, 'predicted_close'] = train_predicted_close
        data.loc[x_test.index, 'predicted_close'] = test_predicted_close
        data['signal_sklearn'] = 0
        for i in range(1, len(data)):
            if data['predicted_close'].iloc[i-1] < data['predicted_close'].iloc[i]:
                data['signal_sklearn'].iloc[i] = 1
            else:
                data['signal_sklearn'].iloc[i] = -1
        data['entry_ols'] = (data['signal_sklearn'] - data['signal_sklearn'].shift(1))
        data['entry_ols'].fillna(0, inplace=True)
#generate the plot
        plt.plot(data.iloc[-252:]['Close'], label='Close')
        plt.plot(data.iloc[-252:]['5-day'], label='5-day')
        plt.plot(data.iloc[-252:]['30-day'], label='30-day')
        plt.plot(data[-252:].loc[(data.entry == -2) & (data.entry_ols == -2)].index, data[-252:]['30-day'][(data.entry == -2) & (data.entry_ols == -2)], 'v', color='r', markersize=10, label='Sell')
        plt.plot(data[-252:].loc[(data.entry == 2) & (data.entry_ols == 2)].index, data[-252:]['5-day'][(data.entry == 2) & (data.entry_ols == 2)], '^', color='g', markersize=10, label='Buy')
        plt.legend()
        plt.title(f"{ticker} - Final decision -  Strategies Combination Chart 3")
        plt.savefig(os.path.join(current_directory, f"{ticker} - Final decision - Strategies Combination Chart 3.png"))
        st.pyplot()
        st.write("Strategy Summary")
        total_trades = len(data[(data.entry_ols == 2) | (data.entry_ols == -2)])
        st.write("Total number of trades:", total_trades)
        
#%%
# define a class as our simulated account,include execute trade, export balance to excel and display the outcome on streamlit
class SimulatedAccount:
    # Initialize attributes for the simulated account
    def __init__(self, stock_name , initial_cash=100000):
        self.stock_name = stock_name
        self.cash = initial_cash   
        self.holdings = 0          
        self.last_price = None   
        self.account_info = pd.DataFrame(columns=['Date', 'Holdings', 'Cash Balance', 'Last Close Price', 'entry', 'entry_ols', 'last_price'])
        
     # Method to execute trades based on given signals
    def execute_trade(self, i, row):
        # Determine the price for the trade
        price = row['Close']
     # Execute buy or sell actions based on signal conditions
        if (row['entry'] == 2) & (row['entry_ols'] == 2):
            self.holdings = self.holdings + self.cash // price
            self.cash = self.cash - self.holdings * price 
            st.write(f"Buying: {self.holdings} holdings at {price}. Cash: {self.cash}")
        elif (row['entry'] == -2) & (row['entry_ols'] == -2):
            self.cash = self.cash + self.holdings * price
            self.holdings = 0
            st.write(f"Selling: All holdings at {price}. Cash: {self.cash}")
            # Update account information with trade details
        self.last_price = price
        self.account_info.loc[len(self.account_info)] = [row.name, self.holdings, self.cash, price, row['entry'], row['entry_ols'], price]
#%%
# Method to export account information to an Excel file
    def get_account_info(self):
        filename = f'{self.stock_name}_account_info.xlsx'
        self.account_info.to_excel(filename, index=False)
        return filename
#%%
# Method to generate a combination chart based on trade signals
    def generate_plots(self, data):
        plt.plot(data['Close'], label='Close')
        plt.plot(data['5-day'], label='5-day')
        plt.plot(data['30-day'], label='30-day')
        plt.plot(data.loc[(data.entry == -2) & (data.entry_ols == -2)].index, data['30-day'][(data.entry == -2) & (data.entry_ols == -2)], 'v', color='r', markersize=10, label='Sell')
        plt.plot(data.loc[(data.entry == 2) & (data.entry_ols == 2)].index, data['5-day'][(data.entry == 2) & (data.entry_ols == 2)], '^', color='g', markersize=10, label='Buy')
        plt.legend()
        plt.title(f"{self.stock_name} - Final decision: Strategies Combination Chart")
        return plt

# Streamlit section for simulated results and account details
with simulated_account:
    # Streamlit section for simulated results and account details
    st.title("Simulated Results")
    st.header("Introduction")
    st.write("This is a simulated trading account that executes buy and sell trades based on the given strategy. It keeps track of the cash balance, holdings, and the last close price. Trades are executed when both the 'entry' and 'entry_ols' signals are 2 (buy) or -2 (sell).")
    st.write("The 'execute_trade' function performs the buying and selling actions based on the signals and updates the account information.")
    st.write("The 'get_account_info' function saves the account information to an Excel file.")
    st.write("The 'generate_plots' function generates the final combination chart.")
    account_instances = {}
    
    # Loop through ticker_list to simulate trading for each stock
    for ticker in ticker_list:
        # Loop through ticker_list to simulate trading for each stock
        data = yf.download(ticker, start=Start_date, end=End_date)
        data['day'] = day
        data = data[['day', 'Open', 'High', 'Low', 'Close']]
        data['5-day'] = data['Close'].rolling(5).mean().shift()
        data['30-day'] = data['Close'].rolling(30).mean().shift() 
        data['signal'] = 0
        for i in range(len(data)):
            if data['5-day'].iloc[i] > data['30-day'].iloc[i]:
                data.loc[data.index[i], 'signal'] = 1
            elif data['5-day'].iloc[i] < data['30-day'].iloc[i]:
                data.loc[data.index[i], 'signal'] = -1
            else:
                data.loc[data.index[i], 'signal'] = 0
        data.dropna(inplace=True)
        data['return'] = np.log(data['Close']) - np.log(data['Close'].shift(1))
        data['system_return'] = data['signal'] * data['return']
        data['entry'] = data['signal'] - data['signal'].shift(1)
        data['entry'].fillna(0, inplace=True)
        x = data[['Open', 'High', 'Low', 'Close', '5-day', '30-day']]  
        y = data['Close']
        model = LinearRegression()
        x_train = x.iloc[:-126]
        x_test = x.iloc[-126:]
        y_train = y.iloc[:-126]
        y_test = y.iloc[-126:]
        model.fit(x_train, y_train)
        train_predicted_close = model.predict(x_train)
        test_predicted_close = model.predict(x_test)
        data['predicted_close'] = 0
        data.loc[x_train.index, 'predicted_close'] = train_predicted_close
        data.loc[x_test.index, 'predicted_close'] = test_predicted_close
        data['signal_sklearn'] = 0
        for i in range(1, len(data)):
            if data['predicted_close'].iloc[i-1] < data['predicted_close'].iloc[i]:
                data['signal_sklearn'].iloc[i] = 1
            else:
                data['signal_sklearn'].iloc[i] = -1
        data['entry_ols'] = (data['signal_sklearn'] - data['signal_sklearn'].shift(1))
        data['entry_ols'].fillna(0, inplace=True)
        st.header(f"Simulated Account - {ticker}")
        st.write(f"This simulation is for the {ticker} stock, initial cash is 100000")
        
        # Create a new SimulatedAccount instance
        account = SimulatedAccount(ticker, initial_cash=100000)
        
        # Iterate over the data and execute trades
        for i, row in data.iterrows():
            account.execute_trade(i, row)
            
        # Display and allow download of account information
        st.subheader("Account Information")
        filename = account.get_account_info()
        st.write(account.account_info)

        # Generate and display the combination chart in a new figure
        st.subheader("Strategies Combination Chart")
        plot = account.generate_plots(data)
        st.pyplot(plot)
        # Clear previous plot
        plt.clf()#This line is from GPT
#%%
# Get current working directory

# functions to send emails
current_directory = os.getcwd()
print(current_directory)
def send_email(receiver, subject, body):
    try:
        # launch outlook
        outlook = win32com.client.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        
        # set email parameters
        mail.To = receiver
        mail.Subject = subject
        mail.Body = body
        
        # use join() to add attachments from current working directory
        for ticker in ticker_list:
            attachment_1 = os.path.join(current_directory, f"{stock} - Momentum Strategy Chart 1.png")
            attachment_2 = os.path.join(current_directory, f"{ticker} - Scikit-Learn Signals Chart 2.png")
            attachment_3 = os.path.join(current_directory, f"{ticker} - Final decision - Strategies Combination Chart 3.png")
            attachment_4 = os.path.join(current_directory, f"{ticker}_account_info.xlsx")
            mail.Attachments.Add(attachment_1)
            mail.Attachments.Add(attachment_2)
            mail.Attachments.Add(attachment_3)
            mail.Attachments.Add(attachment_4)
        
        # send email
        mail.Send()
        st.success("Email sent successfully!")
    except Exception as e:
        st.error(f"Error sending email: {e}")

# email tab
with email_attachments:
    pythoncom.CoInitialize()#this line is from Gpt
    st.header("Send Email with Attachments")
    receiver = st.text_input("Receiver")
    subject = st.text_input("Subject")
    body = st.text_area("Body")
    
    if st.button("Send Email"):
        if receiver and subject and body:
            send_email(receiver, subject, body)
        else:
            st.warning("Please fill in all fields!")

#%%
# News section
with news:
    sn = StockNews(ticker_list, save_news=True, wt_key='328798da8ca8345d60a117816c98e9f7')
    for ticker in ticker_list:
        df_news = sn.read_rss()
        for i in range(1,11):
            st.subheader(f'News of {ticker},{i} of 10')
            st.write(df_news['published'][i])
            st.write(df_news['title'][i])
            st.write(df_news['summary'][i])
            title_sentiment = df_news['sentiment_title'][i]
            st.write(f'Title sentiment: {title_sentiment}')
            news_sentiment = df_news['sentiment_summary'][i]
            st.write(f'News Sentiment: {news_sentiment}')
#%%
# Fundamental data section
with fundamental_data:
    # Fundamental data section, output will be a dataframe
    # key = 'BFHGMS3WCQSF3U3C'
    key = 'CNG3QHYISPDMUX9R'
    fd = FundamentalData(key, output_format='pandas')
#%%
    for ticker in ticker_list:
        try:
            # Fundamental data section
            st.subheader(f'Income Statement From Alpha Vantage for {ticker}')
            income_statement = fd.get_income_statement_annual(ticker)
            if income_statement is None or len(income_statement) == 0:  # Check if the tuple is empty
                st.warning(f"No income statement available for {ticker}")
            else:
                income_statement = income_statement[0]  # Access the first element of the tuple
                st.write(income_statement)
            
            # Fundamental data section
            st.subheader(f'Cash flow From Alpha Vantage for {ticker}')
            cash_flow_quarterly = fd.get_cash_flow_annual(ticker)
            if cash_flow_quarterly is None or len(cash_flow_quarterly) == 0:  # Check if the tuple is empty
                st.warning(f"No cash flow statement available for {ticker}")
            else:
                cash_flow_quarterly = cash_flow_quarterly[0]  # Access the first element of the tuple
                st.write(cash_flow_quarterly)
                
            # Display balance sheet data
            st.subheader(f'Balance sheet From Alpha Vantage for {ticker}')
            balance_sheet = fd.get_balance_sheet_annual(ticker)
            if balance_sheet is None or len(balance_sheet) == 0:  # Check if the tuple is empty
                st.warning(f"No balance sheet available for {ticker}")
            else:
                balance_sheet = balance_sheet[0]  # Access the first element of the tuple
                st.write(balance_sheet)
        
        # Handle ValueErrors
        except ValueError as value_error:
            st.warning(f"Error getting data for {ticker}: {value_error}")
    
