#Imports
import pandas as pd
import numpy as np
from yahooquery import Ticker
from datetime import datetime
import time
import traceback

#define functions
def check_existence(stock):
    skip = False
    result = Ticker(stock).price[stock]
    if "Quote not found for ticker symbol:" in result:
        print(stock, " is not found/recognized, SKIPPING")
        skip = True
    return skip

def get_price_marketCap(stock):
    tick = Ticker(stock)
    ticker = tick.price[stock]
    price = ticker['regularMarketPrice']
    marketCap = ticker['marketCap']
    return price, marketCap

def get_outstandingShares_enterpriseValue_peg(stock):
    tick = Ticker(stock)
    ticker = tick.key_stats[stock]
    shares_outstanding = ticker['sharesOutstanding']
    enterprise_val = ticker['enterpriseValue']
    peg = ticker['pegRatio']
    return shares_outstanding,enterprise_val,peg

def get_totalDebt_totalCash_EBITDA(stock):
    tick = Ticker(stock)
    ticker = tick.financial_data[stock]
    debt = ticker['totalDebt']
    cash = ticker['totalCash']
    EBITDA = ticker['ebitda']
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
    data = tick.earnings[stock]['financialsChart']['quarterly']
    total_data_num = len(data)
    if total_data_num >= 2:
        latest_quarter = data[len(data)-1]['revenue']
        quarter_prior = data[len(data)-2]['revenue']
        quarterly_revenue_growth = (latest_quarter-quarter_prior)/quarter_prior*100
    else:
        quarterly_revenue_growth = 'No Data'
    return quarterly_revenue_growth
    
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


if __name__ == "__main__":
    try:
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

        #Loop through original list and perform the shit
        for company in companies_list:
            if check_existence(company):
                continue
            print("Getting data for: ", company)
            company_data_dict = change_to_dictionary(get_all_data(company))
            df_result_temp = pd.DataFrame(company_data_dict, index=[0])
            df_main = pd.concat([df_main,df_result_temp], ignore_index=True)


        #replace empty strings with nan for mean calculation
        df_main = df_main.replace(r'^\s*$', np.nan, regex=True)

        #Calculate Averages
        print("Calculating averages")
        list_of_columns_to_calculate_average = ['EBITDA/EBIT MARGIN (%)','EBITDA/EBIT PROJ GROWTH (%)','EV/REVENUE (x)',
                                                    'EV/EBITDA (x)','EV/EBIT (x)','PEG 5Y Expected(x)']

        average_row_dict = {}
        for col_name in list_of_columns:
            if col_name == 'COMPANY NAME':
                average_row_dict[col_name] = 'Average/Median'
            elif col_name in list_of_columns_to_calculate_average:
                print(col_name)
                average_row_dict[col_name] = df_main[col_name].mean()
            else:
                average_row_dict[col_name] = None

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

        time.sleep(5)
    except Exception:
        print(traceback.format_exc())
        print("If you don't understand the error please show it to the developer")
        time.sleep(100)