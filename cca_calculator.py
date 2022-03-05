#Imports
import pandas as pd
import numpy as np
from yahooquery import Ticker
from datetime import datetime

#Initialize dataframe
list_of_columns = ['COMPANY NAME','SHARE PRICE ($/share)','OUTSTANDING SHARES','MARKET CAP ($M)',
                                'TOTAL DEBT ($M)','TOTAL CASH ($M)','DILUTED EPS ($/share)','ENTERPRISE VALUE ($)',
                                'REVENUE ($)','QUARTERLY REVENUE GROWTH (%)','EBITDA ($M)','EBIT ($M)',
                                'EBITDA/EBIT MARGIN (%)','EBITDA/EBIT PROJ GROWTH (%)','EV/REVENUE (x)',
                                'EV/EBITDA (x)','EV/EBIT (x)','PEG 5Y Expected(x)']

df_main = pd.DataFrame(columns = list_of_columns)


#Get input by typing or from a csv
csv_input_option = input('Input companies via CSV?: (type only "yes" or "no")')
if csv_input_option == 'yes':
    csv_dir = input('Paste full file path here: ')
    print("Entered directory: ", csv_dir)
    try:
        csv_list = pd.read_csv(csv_dir, header=None)
        companies_list = list(csv_list[0])
    except:
        print("File not found, please restart application")
elif csv_input_option == 'no':
    companies_string = input('Type in companies stock ticker concatenanted by comma: ')
    companies_list = list(companies_string.split(","))
else:
    print("Invalid input, please restart application")
    
print("Companies to be calculated: ", companies_list)

#define functions
def get_price_marketCap(stock):
    tick = Ticker(stock)
    return tick.price[stock]['regularMarketPrice'], tick.price[stock]['marketCap']

def get_outstandingShares_enterpriseValue_peg(stock):
    tick = Ticker(stock)
    shares_outstanding = tick.key_stats[stock]['sharesOutstanding']
    enterprise_val = tick.key_stats[stock]['enterpriseValue']
    peg = tick.key_stats[stock]['pegRatio']
    return shares_outstanding,enterprise_val,peg

def get_totalDebt_totalCash_EBITDA(stock):
    tick = Ticker(stock)
    debt = tick.financial_data[stock]['totalDebt']
    cash = tick.financial_data[stock]['totalCash']
    EBITDA = tick.financial_data[stock]['ebitda']
    return debt, cash, EBITDA

def get_dilutedEps_revenue_EBIT(stock):
    tick = Ticker(stock)
    data = tick.all_financial_data()
    index_last = len(data)-1
    diluted_eps = data.iloc[index_last]['DilutedEPS']
    revenue = data.iloc[index_last]['TotalRevenue']
    ebit = data.iloc[index_last]['EBIT']
    return diluted_eps, revenue, ebit

def get_quarterlyRevenueGrowth(stock):
    tick = Ticker(stock)
    data = tick.earnings_trend[stock]['trend'][0]['growth']
    return data
    
def express_in_MM(number):
    return number/1_000_000

def change_to_dictionary(data_list):
    return_dict = {}
    for i, col_name in enumerate(list_of_columns):
        return_dict[col_name] = data_list[i]
    return return_dict

def get_all_data(stock):
    company_name = stock
    share_price, market_cap = get_price_marketCap(stock)
    outstanding_shares, ev, peg = get_outstandingShares_enterpriseValue_peg(stock)
    total_debt, total_cash, ebitda = get_totalDebt_totalCash_EBITDA(stock)
    diluted_eps, revenue, ebit = get_dilutedEps_revenue_EBIT(stock)
    quarterly_revenue_growth = get_quarterlyRevenueGrowth(stock)
    ebitda_ebit_margin = ebitda/revenue*100
    ev_revenue = ev/revenue
    ev_ebitda = ev/ebitda
    ev_ebit = ev/ebit
    ebitda_ebit_proj_growth = ''
    ordered_list = [company_name, share_price, outstanding_shares, express_in_MM(market_cap),
                       express_in_MM(total_debt), express_in_MM(total_cash), diluted_eps, ev, 
                       revenue, quarterly_revenue_growth, express_in_MM(ebitda), express_in_MM(ebit), 
                       ebitda_ebit_margin, ebitda_ebit_proj_growth, ev_revenue,
                       ev_ebitda, ev_ebit, peg]
    return ordered_list

#Loop through original list and perform the shit
for company in companies_list:
    print("Getting data for: ", company)
    company_data_dict = change_to_dictionary(get_all_data(company))
    df_result_temp = pd.DataFrame(company_data_dict, index=[0])
    df_main = pd.concat([df_main,df_result_temp], ignore_index=True)

#Calculate Averages
print("Calculating averages")
average_row_dict = {}
for col_name in list_of_columns:
    if col_name == 'COMPANY NAME':
        average_row_dict[col_name] = 'Average/Median'
    elif col_name == 'EBITDA/EBIT MARGIN (%)':
        average_row_dict[col_name] = df_main["EBITDA/EBIT MARGIN (%)"].mean()
    else:
        average_row_dict[col_name] = 0

df_result_temp = pd.DataFrame(average_row_dict, index=[0])
df_main = pd.concat([df_main,df_result_temp], ignore_index=True)


#save the file
today = str(datetime.now())
today = today.replace(":", ".")

#Export to excel
result_save_folder = input("Enter result save folder, or leave blank for same location: ")
if result_save_folder == '':
    df_main.to_excel(today+".xlsx", sheet_name='CCA_Calculator_results')
    print("Results saved to: ", today+".xlsx")
else:
    df_main.to_excel(result_save_folder+'//'+today+".xlsx", sheet_name='CCA_Calculator_results')
    print("Results saved to: ", result_save_folder+'//'+today+".xlsx")