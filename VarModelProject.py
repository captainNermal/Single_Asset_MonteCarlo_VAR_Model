"""
Name: Smith Joshua 
Date: April, 2023
Project: Var Model - Monte Carlo 
"""

"""
Description: Single Asset Monte Carlo VAR (Value at Risk Model)
    - VAR model: risk management model/ tool  utalized by institutions regarding equity portfolios, pension fund applications (equity, debt or MM/ Cash), or Hedge funds
        -> Using Monte Carlo Method over parametric and historical: parametric (RELIES ON ASSUMPTIONS - IN THIS CASE NORMAL DATA) simplistic but utalizes normalized data...
        as we know in the real world DATA IS NOT NORMALLY DISTRIBUTED, to avoid this flaw we shall use Monte Carlo and tolerate a higher degree of complexity to obtain greater "accuracy"
    - VAR model essentially states that we are X% certain that we will not loose more than $V amt of dollars by/or at time T 
        -> X percent certain: an established confidence level as a product of probability distribution for losses or gains at time T.
        (losses are negpative gains and gains are negative losses), if 97% confident, we look to losses at 3rd percentile of gains at right tail of distribution 
        -> V dollar amt 
        -> T time horizon
    - Limitations of the Model: VAR Models, although used in industry, are fundementally broken and must be taken with a grain of salt
    - CVAR: a measure beyond VAR that is more robust. Always >= VAR. Accounts for tail risk
    as a "gauge" opposed to an absolute accurate figure
        -> Limitations unique to Monte Carlo Model: time consuming and complicated 
        -> Advantages to Monte Calro Model: not as influenced by extreme BLACK SWANS like historical or parametric as the model does not solely rely on historical data 
    - Single Asset: we shall use single asset for for the sake of simplicity
    - Historical Data: Although MonteCarlo does not rely solely on historical data like other models, we will be using a 5 year period of historical data to comply 
    with industry standards (5 year rolling beta etc) to ensure that we are in the most accurate snapshot of relevant market trends and conditions 
"""

"""
External libraries and Sources used: A variety of libraries and external websites shall be used to manipuate and obtain/ scrape data to avoid the simplicity of File I/O
    Libraries:
        -> Numpy: used for statistical data manipulation
        -> Datetime: to assist with date interaction and manipulation
        -> Pandas: used for statistical data manipulation
        -> yfinance: utalized to extract data from websitees and online sources via scraping
        -> Pandas_makret_calendars: utalized to assist in finding the closest trading dates exempt of holidays and market shutdowns to aid in scraping for pricing data 
            -o: https://www.learndatasci.com/tutorials/python-finance-part-yahoo-finance-api-pandas-matplotlib/
        -> Matplot: used for statistical data manipulation and graphical respresentation
        -> xlwings: library that writes to excel and utalizes excel as an interface for python 
        -> urllib.request: utalized to get URL quests to help obtain URL tags for webscraping of live (slight time delay) prices and rf rate 
        -> beautiful soup: library used for webscraping and webscraping features/ function
        -> sortedcontainers: replacing a traditional sorting algorithm such as insertsort (O(n^2)) to avoid recursion limit error. It is possible to turn off recursion limit but highly cautioned against.
           thus as we are working with 10000 data points we shall import a method. 
    Website(s):
        -> Yahoo Finance: utalized as a proxy for a conventional platform often found within financial institutions such as bloomberg terminal or Capital IQ
"""

"""
Getting started:
    xlwings installation and excel xlwings compatability:
        -> from terminal: pip install xlwings --> python library installation
        -> from command prompt: xlwings addin install --> excel adds vba reference add in and creates xlwings tab
    Directories:
        -> open command prompt
            -o ensure your working directory is in desired location: check; cd | change (if required); cd ___________ <- desired path
        -> create xslm and py file pair (we can do this manually but it is a much slower process): xlwings quickstart _____ <- desired name of folder containing xlsm and py file
    Py file: as mentioned above, this entire "getting started configuration can be done manually however the auto set up through command promp is easier
        -> def main():
            wb = xw.Book.caller()
            desired_sheet = wb.sheets("Sheet1")
           if __name__ == "__main__":
            xw.Book("VarModelProject.xlsm").set_mock_caller()
            main()

            ^ main() shall serve as a wrapper function 
            ^ "if __name__ == "__main__"" and "wb = xw.Book.caller()" allow dynamic movement of py file and xlsm such that they do not have to be in the same folder or same directory
                -o very useful when we want to transport or share the model itself
    UDF connection (connect excel to python script): 
        -> always refernce your main wrapper function as function you call from py file 
            -o RunPython (" import VarModelProject; VarModelProject.main()") --> RunPython (" import script/file name; script/file name.main()")
                -* never pass variables through the function you are referencing
"""

#library implementation 
import xlwings as xw
import numpy as np
import datetime
import pandas as pd 
import pandas_market_calendars as mcal 
import yfinance as yf
import matplotlib.pyplot as plt
from bs4 import BeautifulSoup
from urllib.request import urlopen
import re 
from sortedcontainers import SortedList


"""
User input to define perameters of the model itself prior to computation
    -Confidence Interval: Default 95, can be manually adjusted
    -Value of Portfolio: No defualt - specification required
    -Number of Iterations: Default 10,000; standard per monte carlo methodology
    -Underlying Asset Ticker: can be equity commodity or particular index 
"""
#user input
def user_model_input(wb):
    #specify sheet 
    dataplug_Scraper_sheet = wb.sheets("DataPlug|ScrapeDriver")
    
    #initiate VBA Input box & append input with list index
    input_Command_list = ["Confidence Interval", "Value of Portfolio", "Number of Iterations", "Underlying Asset Ticker", "Specify Exchange"]
    input_Command_list_feedback = []
    for i in range(len(input_Command_list)):
        rng = wb.app.api.InputBox(f"{input_Command_list[i]}")
        input_Command_list_feedback.append(rng)
    
    dataplug_Scraper_sheet.range("A1").value = input_Command_list
    dataplug_Scraper_sheet.range("A2").value = input_Command_list_feedback

    return input_Command_list_feedback



#grab dates for respective tickers
def ticker_historicalData_dateLocation(wb,input_Command_list_feedback):
    #establish displayed infromation 
    #desired sheet
    var_Model_sheet = wb.sheets("VARModel")
    #ticker input 
    ticker = input_Command_list_feedback[3]
    exchange = input_Command_list_feedback[4]
    #returnModeltype & asset ticker
    var_Model_sheet.range("A1").value = " Single Asset Monte Carlo VAR"
    var_Model_sheet.range("A2").value = ticker
    var_Model_sheet.range("B2").value = exchange

    #retrieve todays date and the date five years ago to the day
    #end date - get todays date as a basepoint for the closest trading date for specified exchange to specify an endpouint of data series
    todays_date = datetime.datetime.today()
    #retrieve todays date - 5 years - start date
    fiveYearsago_date = todays_date - datetime.timedelta(days=365*5)

    #reteive closest trading dates the dates described above 
    #grab exchange for sake for retrieving trading day
    exchangename_date_retrieve = f"{input_Command_list_feedback[4]}"
    #get market calendar from pandas librray for specified exchange
    exchange_dates = mcal.get_calendar(exchangename_date_retrieve)
    #generate schedule of trading dates between now and 5 years ago on a daily increment 
    tradingRange_dates = exchange_dates.schedule(start_date=fiveYearsago_date, end_date=todays_date)
    #retrieve closest trading dates to today and date 5 years ago 
    if not tradingRange_dates.empty:
        closestToday_date = tradingRange_dates.index[-1].date()
        closestFiveyearsago_date = tradingRange_dates.index[0].date()
    else:
        closestToday_date = "error"
        closestFiveyearsago_date = "error"
    #dictionary of position: date
    time_Range_dict = {"Start Date:":closestToday_date, "End Date:": closestFiveyearsago_date}
    for i, (key, value) in enumerate(time_Range_dict.items()):
        var_Model_sheet.range(f'A{i+4}').value = key
        var_Model_sheet.range(f'B{i+4}').value = value

    return closestToday_date, closestFiveyearsago_date, var_Model_sheet



#retreive historical data
def ticker_historicalData_retrieve(closestToday_date, closestFiveyearsago_date, var_Model_sheet, input_Command_list_feedback):          
    #convert closest exchange date to string to assist in retreiveal process
    startDate_retrieve = f'{closestFiveyearsago_date}'
    endDate_retrieve = f'{closestToday_date}'
    #convert ticker to string to assit in retreival
    ticker_Retrieve = f'{input_Command_list_feedback[3]}'
    ticker_Retrieve_yfformat = yf.Ticker(ticker_Retrieve)

    #retrieve call
    historical_data = ticker_Retrieve_yfformat.history(period="1d", start=startDate_retrieve, end=endDate_retrieve)

    #manipulate data - yfinance already places it in dataframe format quicker than CSV file IO 
    historical_data_Close = historical_data[["Close", "Volume"]]

    #Place data
    var_Model_sheet.range("A7").value = historical_data_Close

    return historical_data_Close


"""
we use log normal returns as they are better to use cumulatively, taking into account compounded returns. Moreover, Lognormal returns account for fat tail and positive skew distribution (more common than expected under normal distribution) 
accounting for black swan outliers

finaincal data is right skewed as high concentration of small retunrs positive and lower concentration of high returns, so log of the data mkes the distru=ibution more symetrical by compressing large returns and expanding small returns 
then fat tails, log gives fatter tails which account for extreme events 
"""
def logNormaldaily_returns(var_Model_sheet, historical_data_Close):
    #define column 
    var_Model_sheet.range("E7").value = "LogNormal Returns"
    
    #log normal return list 
    logNormal_return_daily_list = []
    #log returns of each datapoint
    for i in range(len(historical_data_Close)-1):
        
        logNormal_return_daily = np.log((historical_data_Close['Close'][i]) / (historical_data_Close['Close'][i-1]))
        logNormal_return_daily_list.append(logNormal_return_daily)
        var_Model_sheet.range(f'E{i+8}').value = logNormal_return_daily

    return logNormal_return_daily_list



#get 10 year risk free rate 
def retrieve_rf_rate():
    #fed website URL 
    rf_url = "https://home.treasury.gov/resource-center/data-chart-center/interest-rates/TextView?type=daily_treasury_real_yield_curve&field_tdr_date_value=2023"
    html=urlopen(rf_url)
    soup=BeautifulSoup(html)
    
    tbill_list = []

    allrows = soup.find_all("tr")
    for i in allrows:
        row_list = i.find_all("td")
    for i in row_list:
        tbill_list.append(i.text)
    
    tenyear_rf = tbill_list[3]

    tenyear_rf = float(tenyear_rf)

    return tenyear_rf



#basic stats
def basic_stats(var_Model_sheet, historical_data_Close,logNormal_return_daily_list, tenyear_rf):
    #define column
    var_Model_sheet.range("G7").value = 'Metric(s)'

    #length 
    number_of_observations = len(logNormal_return_daily_list)
    #min 
    min_val = logNormal_return_daily_list[0]
    for i in logNormal_return_daily_list:
        if i < min_val:
            min_val = i

    #max 
    max_val = logNormal_return_daily_list[0]
    for i in logNormal_return_daily_list:
        if i > max_val:
            max_val = i 
    # #average
    average = 0
    for i in range(len(logNormal_return_daily_list)):
        average = average + logNormal_return_daily_list[i]
    
    average = average / number_of_observations
    #std dev 
        #square abs val list
    var_list = []
    for i in range(len(logNormal_return_daily_list)):
        var_val = (logNormal_return_daily_list[i] - average)**2
        var_list.append(var_val)
        #sum values above to retrieve numerator 
    var_list_sum = 0
    for i in range(len(var_list)):
        var_list_sum = var_list_sum + var_list[i]
        #fial calc 
    std_dev = (var_list_sum / number_of_observations)**(1/2)

    #annualized data 
    number_trading_days_per_year = 252
    
    time_increment = 1 / number_trading_days_per_year

    annualized_Mean_return = ((number_trading_days_per_year)**(1/2)) * average

    annualized_Stddev = (number_trading_days_per_year**(1/2)) * std_dev

    annualized_expected_return = annualized_Mean_return - ((annualized_Stddev**2)/2)

    annualized_sharpe_ratio = ((average - (tenyear_rf/100)) / std_dev) * (number_trading_days_per_year)**(1/2)

    #placement 
    var_Model_sheet.range("H8").value = number_of_observations
    var_Model_sheet.range("H9").value = f'{min_val*100} %'
    var_Model_sheet.range("H10").value = f'{max_val*100} %'
    var_Model_sheet.range("H11").value = f'{average*100} %'
    var_Model_sheet.range("H12").value = f'{std_dev*100} %'

    var_Model_sheet.range("G8").value = 'Number of Obs.'
    var_Model_sheet.range("G9").value = f'Minimum return'
    var_Model_sheet.range("G10").value = f'Maximum return'
    var_Model_sheet.range("G11").value = f'Average'
    var_Model_sheet.range("G12").value = f'Std Dev.'

    #annualzied data placement
    var_Model_sheet.range("H14").value = number_trading_days_per_year
    var_Model_sheet.range("H15").value = time_increment
    var_Model_sheet.range("H16").value = f'{annualized_Mean_return*100} %'
    var_Model_sheet.range("H17").value = f'{annualized_Stddev*100} %'
    var_Model_sheet.range("H18").value = f'{annualized_expected_return*100} %'
    var_Model_sheet.range("H19").value = f'{annualized_sharpe_ratio}'
    

    var_Model_sheet.range("G14").value = 'Number of Trading Days'
    var_Model_sheet.range("G15").value = 'Time Increment'
    var_Model_sheet.range("G16").value = 'Annualized Mean Return'
    var_Model_sheet.range("G17").value = 'Annualized Std Dev'
    var_Model_sheet.range("G18").value = 'Annualized Expected Return'
    var_Model_sheet.range("G19").value = 'Annaulized Sharpe Ratio'
    
    return time_increment, annualized_expected_return, annualized_Stddev
    


#seed value is arbitrary 
#modulus (2**31) - 1  below is a mersenne prime number (can e expressed as 2^n-1), using modulus prime number helps ensure long and uniform cycle of random numbers, reducing pattern or repetition in sequence: modulus used in range generation for random numbers 
#multiplier: bwlow number 7^5 leads to a the same value as the modulus value ensuring no pattern reoccurance or repetition
def simulation_parameters(var_Model_sheet, input_Command_list_feedback):
    #set values 
    seed_value = 4352545
    modulus = (2**31) - 1 
    multiplier = 7**5
    number_Iterations = input_Command_list_feedback[2]

    #placements
    var_Model_sheet.range("H21").value = seed_value
    var_Model_sheet.range("H22").value = modulus
    var_Model_sheet.range("H23").value = multiplier
    var_Model_sheet.range("H24").value = input_Command_list_feedback[2]

    var_Model_sheet.range("G21").value = 'Seed-Value'
    var_Model_sheet.range("G22").value = 'Modulus'
    var_Model_sheet.range("G23").value = 'Multiplier'
    var_Model_sheet.range("G24").value = 'Number of Iterations'

    return seed_value, modulus, multiplier, number_Iterations


#run simulation
def simulation(seed_value, modulus, multiplier, number_Iterations, var_Model_sheet, time_increment, annualized_expected_return, annualized_Stddev):
    #iteration type conversion
    number_Iterations_conversion = int(number_Iterations)
    
    #list setup
    preliminary_Number = []
    random_number = []
    standard_Normal_cummulative_inverse_distribution = []
    simulation_return = []

    #title establishment
    var_Model_sheet.range('H26').value = 'Preliminary Number'
    var_Model_sheet.range('I26').value = 'Random Probability'
    var_Model_sheet.range('J26').value = 'Cumm Norm Inverse Dis'
    var_Model_sheet.range('K26').value = 'Simulation Return'

    # prelim number initialization 
    preliminary_Number_start = (multiplier * seed_value) % modulus
    preliminary_Number.append(preliminary_Number_start)

    #establish Cumm norm inverse dis parameters 
    sim_mean = 0 
    sim_std = 1

    #loop to establish parameters
    # inverse normal CDF is used to convert a probability into a zscore which is used to estimate impact of on returns: z scores represent how many std devs a given value is from mean (in this case 0)
    #simulation return makes sense as we take time increment and multiply by expected return, from which point we add the standard deviation component to represent the expected impact on the return
        #we square time on std dev because: we see volatilty increase as periods get longer, square reflects cummulative impact of small price changes over time - if we didnt square for time with the way vvarince acts it woudlnt add up properly
        #we add zscore in formula to account for how many std deviations we are away from the mean, then scaling it by time. We are essentially adjsuting annual std dev by zscore (based off probability from Random Number) to adjust annual std dev
    for i in range(number_Iterations_conversion):
        #preliminary number
        preliminary_Number_nextval = (preliminary_Number[i] * multiplier) % modulus
        preliminary_Number.append(preliminary_Number_nextval)
        var_Model_sheet.range(f'H{i+27}').value = preliminary_Number[i]

        #random number - random probability
        random_Number_nextval = preliminary_Number[i] / modulus
        random_number.append(random_Number_nextval)
        var_Model_sheet.range(f'I{i+27}').value = random_number[i]

        #Inverse value of standard normal cummulative distribution - deploy random normal to get z score to see how many std devs we are away from mean
        cumm_Norm_inverse_dis_nextval = np.percentile(np.random.normal(sim_mean,sim_std,number_Iterations_conversion), 100 * random_number[i])
        standard_Normal_cummulative_inverse_distribution.append(cumm_Norm_inverse_dis_nextval)
        var_Model_sheet.range(f'J{i+27}').value = standard_Normal_cummulative_inverse_distribution[i] 

        #Simulation return 
        simualtion_Return_nextval = (time_increment * annualized_expected_return) + (annualized_Stddev*standard_Normal_cummulative_inverse_distribution[i]*((time_increment)**(1/2)))
        simulation_return.append(simualtion_Return_nextval)
        var_Model_sheet.range(f'K{i+27}').value = simulation_return[i] 

    return preliminary_Number, random_number, standard_Normal_cummulative_inverse_distribution, simulation_return



# could use quicksort recursively - fast method great for large datasets with asymptotic complexity of nlogn. Only issue is that 10000 datapoint exceeds python recursion limit
# thus we need to use a non recursive method that is computationally efficient --> insert sort best case O(n) worst case O(n^2) | directly size of dataset or dataset sqaured (which is bad)
    #see example below of sorting algorithm, we could logially turn off recursion limit but not wise
# use pre-existing testing sort method   sroted containers sort list
def sort_sim(simulation_return, var_Model_sheet):
    #     if len(simulation_return) < 1:
    #         sorted_Return_list = simulation_return
    #     else: 
    #         pivot = simulation_return[len(simulation_return) // 2]
    #         left = [x for x in simulation_return if x < pivot]
    #         middle = [x for x in simulation_return if x == pivot]
    #         right = [x for x in simulation_return if x > pivot]

    #         sorted_Return_list = sort_sim(left) + sort_sim(middle) + sort_sim(right)

    #     return sorted_Return_list

    #placement
    var_Model_sheet.range("K26").value = "Sorted Simulation Returns"
    #sort simulation returns 
    sorted_Sim_return = SortedList(simulation_return)
    sorted_Sim_return = list(sorted_Sim_return)
    #place sorted list
    for i in range(len(sorted_Sim_return)):
        var_Model_sheet.range(f'K{i+27}').value = sorted_Sim_return[i]

    return sorted_Sim_return
    

#var and cvar value
def var_Cvar_Calc(input_Command_list_feedback, var_Model_sheet, sorted_Sim_return):
    #establish variables
    confidence_Interval = int(input_Command_list_feedback[0])
    portfolio_Value = int(input_Command_list_feedback[1])
    number_Of_iterations = int(input_Command_list_feedback[2])
    data_Point_value = int(number_Of_iterations * (1-(confidence_Interval/100)))
    #placement
    var_Model_sheet.range("N4").value = confidence_Interval
    var_Model_sheet.range("N5").value = portfolio_Value
    var_Model_sheet.range("N6").value = data_Point_value

    #var calculation - x percent certain that we will not loose more than y amt iof dollar in T time 
    #look to sorted data and then take the kth value of sorted data as how much you will loose a the %th percentile (95% confident means 5% uncertainty) so data point at 5% would be max how much you will loose
    #FROM EXAMPLE ABOVE WE WOULD BE LOOKING FOR THE LOSS AT THE 95TH PERCENTILE
    kth_sortedSim_return_varvalue = sorted_Sim_return[data_Point_value-1]
    var_Model_sheet.range("N8").value = f'{kth_sortedSim_return_varvalue * 100} %'
    var_Dollar_value = kth_sortedSim_return_varvalue * portfolio_Value
    var_Model_sheet.range("N9").value = var_Dollar_value

    #cvar calculation - more robust metric than var, cvar demonstrates what we are capable of loosing after var (always >= var). Accounts for tail risk (losses after var)
        #essentially an average to capture tail risk 
    kth_sortedSim_return_varvalue_denominator = (1/(data_Point_value - 1)) 

    kth_sortedSim_return_varvalue_numerator = 0
    for i in range(data_Point_value - 1):
        kth_sortedSim_return_varvalue_numerator = kth_sortedSim_return_varvalue_numerator + sorted_Sim_return[i]

    cvar_Percent_value = kth_sortedSim_return_varvalue_numerator * kth_sortedSim_return_varvalue_denominator
    cvar_Dollar_value = cvar_Percent_value * portfolio_Value

    var_Model_sheet.range("N11").value = f'{cvar_Percent_value * 100} %'
    var_Model_sheet.range("N12").value = cvar_Dollar_value

    #title placement 
    var_Model_sheet.range("M4").value = 'Confidence Interval'
    var_Model_sheet.range("M5").value = 'Value of Portfolio'
    var_Model_sheet.range("M6").value = 'Kth Target'
    var_Model_sheet.range("M8").value = 'VAR Percent Value'
    var_Model_sheet.range("M9").value = 'VAR Dollar Value'
    var_Model_sheet.range("M11").value = 'CVAR Percent Value'
    var_Model_sheet.range("M12").value = 'CVAR Dollar Value'


#genrate chart
def chart(historical_data, var_Model_sheet):
    #chart generation
    fig = plt.figure()
    #obtain data
    plt.plot(historical_data['Close'])
    #place chart in excel file
    plt.title("5 Year Historical Performance")
    plt.ylabel("Price($)")
    plt.xlabel("Year(N)")
    var_Model_sheet.pictures.add(fig, name="chart1", update=True)
    

#main wrapper
def main():
    #establosh workbook connection adn create dynamic tracking caller
    wb = xw.Book.caller()
    #call above
    input_Command_list_feedback = user_model_input(wb)
    closestToday_date, closestFiveyearsago_date, var_Model_sheet = ticker_historicalData_dateLocation(wb,input_Command_list_feedback)
    historical_data_Close = ticker_historicalData_retrieve(closestToday_date, closestFiveyearsago_date, var_Model_sheet, input_Command_list_feedback)
    logNormal_return_daily_list = logNormaldaily_returns(var_Model_sheet, historical_data_Close)
    tenyear_rf = retrieve_rf_rate()
    time_increment, annualized_expected_return, annualized_Stddev = basic_stats(var_Model_sheet, historical_data_Close, logNormal_return_daily_list, tenyear_rf)
    seed_value, modulus, multiplier, number_Iterations = simulation_parameters(var_Model_sheet, input_Command_list_feedback)
    preliminary_Number, random_number, standard_Normal_cummulative_inverse_distribution, simulation_return = simulation(seed_value, modulus, multiplier, number_Iterations, var_Model_sheet, time_increment, annualized_expected_return, annualized_Stddev)
    sorted_Sim_return = sort_sim(simulation_return, var_Model_sheet)
    var_Cvar_Calc(input_Command_list_feedback, var_Model_sheet, sorted_Sim_return)
    chart(historical_data_Close, var_Model_sheet)

    
#mock caller
if __name__ == "__main__":
    xw.Book("VarModelProject.xlsm").set_mock_caller()
    main()
