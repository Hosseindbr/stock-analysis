#VBA Stock Analysis
##Project Summary
###Purpos
In this challenge a VBA code was used to evaluate the stock investments in the year 2017 and 2018 to assist with future investment in the stock market. In this challenge we aimed to make the code run faster through refactoring the existing code.  
##Results
###Data Analysis
The data wee presented in two sheets each of which contain information regarding ticker, the issuance date, price at the opening, maximum, minimum,  closing, and adjusted closing, as well as, the stock volume in 2017 and 2018. Using this code, ticker, total daily volume, and return value were retrieved. To this end As the first step in refactoring the code, the script needed to create the input box, chart headers, ticker array, and activate the worksheet was added to the Microsoft Visual Basic editor. Afterwards, the refactoring steps were laid out to create the structure. 

>'1a) Create a ticker Index
    >For i = 0 To 11
    >tickerIndex = tickers(i)
    
    >'1b) Create three output arrays
    >Dim tickerVolumes As Long
    >Dim tickerStartingPrices As Single
    >Dim tickerEndingPrices As Single
    
    
    >''2a) Create a for loop to initialize the tickerVolumes to zero.
   > Worksheets(yearValue).Activate
    >tickerVolume = 0
        
    >''2b) Loop over all the rows in the spreadsheet.
    >For j = 2 To RowCount
    
        >'3a) Increase volume for current ticker
        >If Cells(j, 1).Value = tickerIndex Then
        >tickerVolumes = tickerVolume + Cells(j, 8)
        >End If
        
        
        >'3b) Check if the current row is the first row with the selected tickerIndex.
        >'If  Then
        >If Cells(j, 1).Value = tickerIndex And Cells(j - 1, 1).Value <> tickerIndex Then
        >tickerStartingPrices = Cells(j, 6).Value
        
            
            
        >'End If
        >End If
        
        >'3c) check if the current row is the last row with the selected ticker
         >'If the next row‚Äôs ticker doesn‚Äôt match, increase the tickerIndex.
        >'If  Then
        >If Cells(j, 1) = tickerIndex And Cells(j + 1, 1).Value <> tickerIndex Then
        >tickerEndingPrices = Cells(j, 6).Value
        >End If
        >Next j
        
            

            >'3d Increase the tickerIndex.
            >If Cells(j, 1).Value = tickerIndex And Cells(j + 1, 1) <> tickerIndex Then
            >tickerIndex = tickerIndex + 1
            
        >'End If
    >End If
    

    
    >'4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
       
        >Worksheets("All Stocks Analysis").Activate
        >Cells(4 + i, 1).Value = tickerIndex
        >Cells(4 + i, 2).Value = tickerVolume
        >Cells(4 + i, 3).Value = ((tickerEndingPrices / tickerStartingPrices) - 1)
        
    
        
        
    >Next i

##Summary

###Advantages
*Coding is better designed when refactored
*coding can be more easily understood through refactoring
*Bugs can be found through refactoring
*Programming becomes faster when refactoring is used
###disadvantage
*It is risky when the code is big
*There is a risk when there are no proper test cases for the existing code
*There is a risk when the programmer has lack of understanding what the code is designed to do.

It is always easier to change, understand, and maintain a clean and well-organized code. By paying attention to the code refactoring process earlier, you can avoid facing difficulties later. By refactoring, macro run time was decreased significantly. Our new analysis took approximately 0.7 seconds to run, as compared with the original analysis that took approximately one second.

