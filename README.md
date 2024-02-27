# Single Asset MonteCarlo VAR Model
##### Joshua Smith 
## Project Description &#x1F4D3;

Some of the risk attempts to measure numerical,  more specifically monetary risk, and thus the lossess attributed to a portfolio are attributed to Francis Edgeworth as far back as 1888:
<p align="center">
  <b>"We are X% certain that we will not lose more than Y amount of dollars in T amount of time"</b>
</p>

A VAR model (Value at Risk): risk management model/ tool  utalized by institutions regarding equity portfolios, pension fund applications (equity, debt or MM/ Cash), or Hedge funds to state "we are X% Certain that we will not lose more than Y amount of dollars in T amount of time". Two popular methods exist: parametric (assuming normal data and relying solely on historical information); and Monte Carlo, in which data does not have to be normal and we do not rely solely on historical data (due to its simulation style nature).

<p align="center">
  <b>Most "real world" data, especially financial data is not normal; it is very right skewed. This is due to the fact that with financial markets we see a large concentration of small positive returns and a handful of aggregiously high postive returns</b>
</p>

To avoid the "normalicy flaw": MonteCarlo, although more complex is often used to obtain greater accuracy. 

Despite seeming promising, each VAR model variant, including MonteCarlo does possess flaws. VAR models although utalized in industry possess a magnitude of flaws and are at their core fundamentally broken and should be used a "guage" opposed to an absolute value. A collection of pros and cons exist to MonteCarlo Variants: 

* Advantages: not as influenced by extreme BLACK SWANS like historical or parametric models as the model does not solely rely on historical data 
* Draw Backs: time consuming and complicated

## Why this Project &#x2753;
Most VAR models are constructed and built via python to utalize the workabilty of ther language with large datasets through the base language and its respective libraries such as $Numpy$ or $Pandas$. However, these models are sometimes passed off to and maintained by individuals who do not possess sufficient programming knowledge (many financial type analysts are familiar with excel and limited VBA; nothing more, nothing less). Through constructing a VAR model that utalizes the python library $xlwings$, one can write a macro that calls a python script and returns the results into a xlsm or xlsx file. These results can then be worked with or altered (Essentially integratign the usability of python into excel and transforming excel into a GUI). This process drastically reduces the barrier to entry for VAR model comprehension, usability and can reduce the risk of broken VAR models resulting in potential financial loss among other hardships. 

## How to Use Model &#x2611;
1. Ensure the xlsm file is in the same directory as your py file and ensure that you have $macros$ enabled in your excel file
2. Once in sheet click macro $Initiate Data Entry$ on $DataPlug|ScrapeDriver$ tab.
3. A dialogue box will appear that will prompt you to enter five values:
<p align="left">
  <b> Confidence Interval: Default 95% </b>
</p>
<p align="left">
  <b> Value of Portfolio: Enter the notional Value of the portfolio in dollars (any currency is valid) </b>
</p>
<p align="left">
  <b> Number of iterations: Less iterations will be less time consuming but will provide a more innacurate figure. Default MonteCarlo is 10,000 simulations </b>
</p>
<p align="left">
  <b> Underlying Asset: Enter teh ticker value of a publically traded equity or asset: for example for Tesla you would enter TSLA </b>
</p>
<p align="left">
  <b> Specify Exchange: enter the exchange the asset trades on. for instance tesla (TSLA) would be the New York Stock Exchange so enter NYSE </b>
</p>
4. close the window and the simulation will take a few minutes. After such time you will be provided with crucial key metrics, a VAR value and a CVAR value

## Project - Other considerations and Birds Eye View &#x1F426;

##### Other Considerations 
An additonal metric of CVAR (Conditional Value at Risk) will be shown. CVAR is a measure beyond VAR that is more robust and its always greater than VAR as it accounts for tail risk. 

Historical Data: Although MonteCarlo does not rely solely on historical data like other models, we will be using a 5 year period of historical data to comply with industry standards (5 year rolling beta etc) to ensure that we are in the most accurate snapshot of relevant market trends and conditions 

##### Birds Eye View 
* Enter your paramaters above.
* 5 year of historical pricing data from yesterday backwards (with 252 trading days per year) will be gathered from Yahoo Finance via the $yfinance$ library. Holidays will be accounted for via $pandas_market_calendar$ library to avoid innacuracies. Note* A more appropriate data source such as BBGT (Bloomberg terminal) would be utalized in an industry setting, however Yahoo Finance is a sufficient proxy for the time being.
* The log normal value of these returns will be taken (see reasoning above). By doing so we are condensiing larger returns and stretching smaller returns to make the distribution more semetric and "normal" to fit the model appropriately.
* The simulation will then take place within a dataframe with 4 columns and 10000 rows. Each row a simulation.
<p align="left">
  <b> The first column a random number determiend via a modulus, multiplier and seed value. The modulus and multiplier are mersenne prime numbers ensuring low repetition or pattern sequencing within random number generation </b>
</p>
<p align="left">
  <b> The second column a random probability between 0 and 1 (0 < x < 1) derived by dividing the random number by the modulus </b>
</p>
<p align="left">
  <b> The third column a zscore (the number of standard deviations we are away from the mean). This can be done by going from probabilty to z score via the inverse CDF of the probability</b>
</p>
<p align="left">
  <b> Lastly a simulation expected return via formula </b> $$E(r) = \text{Annual Return} \cdot \left(\frac{1}{252}\right) + \text{Annual Std Dev} \cdot \text{ZScore} \cdot \left(\frac{1}{\sqrt{252}}\right)$$ 
</p>
* These simulated returns will then be sorted in ascenting order. If we are 95% confident, we look to lossess in the 5th percentisle (or the kth value being the 5th value in the sorted expected returns). If the 5th vlauer is 8.70% we would state we are 95% confident that we will not loose more $87,000 (8.70% x 1,000,000 notiional portfolio) on a 1,000,000 portfolio of asset X after a timeframe of 5 Years. $CVAR$ will provide a value greater than this.

## Library and Web Usage &#x1F4F6;
Libraries:
* Numpy: used for statistical data manipulation
* Datetime: to assist with date interaction and manipulation
* Pandas: used for statistical data manipulation
* yfinance: utalized to extract data from websitees and online sources via scraping
* Pandas_makret_calendars: utalized to assist in finding the closest trading dates exempt of holidays and market shutdowns to aid in scraping for pricing data 
* Matplot: used for statistical data manipulation and graphical respresentation
* xlwings: library that writes to excel and utalizes excel as an interface for python 
* urllib.request: utalized to get URL quests to help obtain URL tags for webscraping of live (slight time delay) prices and rf rate 
* beautiful soup: library used for webscraping and webscraping features/ function
* sortedcontainers: replacing a traditional sorting algorithm such as insertsort (O(n^2)) to avoid recursion limit error. It is possible to turn off recursion limit but highly cautioned against. Thus as we are working with 10000 data points we shall import a method. 

Website(s):
* Yahoo Finance: utalized as a proxy for a conventional platform often found within financial institutions such as bloomberg terminal or Capital IQ



