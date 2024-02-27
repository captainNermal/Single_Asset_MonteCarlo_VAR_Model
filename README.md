# Single Asset MonteCarlo VAR Model
##### Joshua Smith 
## Project Description 

Some of the risk attempts to measure numerical,  more specifically monetary risk, and thus the lossess attributed to a portfolio are attributed to Francis Edgeworth as far back as 1888:
<p align="center">
  <b>"We are X% certain that we will not lose more than Y amount of dollars in T amount of time"</b>
</p>

A VAR model (Value at Risk): risk management model/ tool  utalized by institutions regarding equity portfolios, pension fund applications (equity, debt or MM/ Cash), or Hedge funds to state "we are X% Certain that we will not lose more than Y amount of dollars in T amount of time". Two popular methods exist: parametric (assuming normal data and relying solely on historical information); and Monte Carlo, in which data does not have to be normal and we do not rely solely on historical data (due to its simulation style nature).

<p align="center">
  <b>Most "real world" data, especially financial data is not normal; it is very right skewed. This is due to the fact that with financial markets we see a large concentration of small positive returns and a handful of aggregiously high postive returns</b>
</p>

To avoid the "normalicy flaw": MonteCarlo, although more complex is often used to obtain greater accuracy. 

Despite seeming promising, each VAR model variant, including MonteCarlo does possess flaws. VAR models although utalized in industry possess a magnitude of flaws and are at their core fundamentally broken. A collection of pros and cons exist to MonteCarlo Variants: 

* Advantages: not as influenced by extreme BLACK SWANS like historical or parametric models as the model does not solely rely on historical data 
* Draw Backs: time consuming and complicated

## Why this Project 
Most VAR models are constructed and built via python to utalize the workabilty of ther language with large datasets through the base language and its respective libraries such as $Numpy$ or $Pandas$. However, these models are sometimes passed off to and maintained by individuals who do not possess sufficient programming knowledge (many financial type analysts are familiar with excel and limited VBA; nothing more, nothing less). Through constructing a VAR model that utalizes the python library $xlwings$, one can write a macro that calls a python script and returns the results into a xlsm or xlsx file. These results can then be worked with or altered (Essentially integratign the usability of python into excel and transforming excel into a GUI). This process drastically reduces the barrier to entry for VAR model comprehension, usability and can reduce the risk of broken VAR models resulting in potential financial loss among other hardships. 

## Project - Other considerations and Birds Eye View

##### Other Considerations


