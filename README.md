# VBA-challenge

**Notes**
  - I did not provide every piece of code in this ReadMe but pointed out how I solved the assignment problems. Full copy of copy is also provided

**Module 2 Objectives**
Create a script that runs through all the stocks for one year and outputs the following information
  1. The Ticker Symbol
        - I first created a variable for this called Ticker
          
              Dim Ticker As String
          
        - I was able to pull the ticker symbol by using a for loop to iterate through the rows and find where i + 1 was not equal to i.
          
              If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                  Ticker = ws.Cells(i, 1).Value
        - Whenever I found this value I used the following code to add it to my table I was creating
          
              ws.Range("I" & Summary_Table_Row).Value = Ticker
          **Code for this part of the assignment was found by using activities we performed in class to find the solution to this step. Activities were Credit Card and Star Counter**
  
  2. Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
       - I started this by declaring a few variables
         
             Dim Yearly_Change As Double
             Dim Close_Price As Double
       - I also set a variable for open price and set it equal to the open price of the first cell
         
             Dim Open_Price As Double
             Open_Price = ws.Cells(2, 3).Value
       - I then set the close price within the for loop by declaring it whenever i was not equal to i + 1.
         
             Close_Price = ws.Cells(i, 6).Value
       - I then determined the yearly change by subtracting the opening price from the close price
         
             Yearly_Change = Close_Price - Open_Price
       - I then returned the yearly change value in the summary table I had created
         
             ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
       - I then made sure to reset the open price before iterating to the next row
     
             ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
        ** Code for this was mostly taken from the Credit Card Activity performed in class where we were shown how to reset variables within loops before the next iteration.**  
  
  3. The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
       - I started this by declaring a variable for percent change
         
             Dim Percent_Change As Double
       - Within the for loop, after I had calculated the yearly change above, I set Percent_Change = to yearly change divided by open price for that stock
         
             Percent_Change = Yearly_Change / Open_Price
       - I then returned the results in the summary table I had created
         
             ws.Range("K" & Summary_Table_Row).Value = Percent_Change
         **Code for this was gathered from the principles we had learned in class**     
  
  4. The total stock volume of the stock. The result should match the following image:
       - I started by set the variable for total stock volume and setting it equal to 0. This was so that I could add to it when iterating through the rows
        
             Dim Total_Stock_Volume As Double
             Total_Stock_Volume = 0
       - Within the for loop I said when the row is not equal to the next row I would add the row volume to the total stock volume
         
             Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
       - I also reset the Total_Stock_Volume to 0 before going to the else part of my statement. This was so I could reset the stock volume to calculate properly for the next stock
     
             Total_Stock_Volume = 0
       - I also said by using an Else statement that I would add the volume row total to the total stock volume if the stock on the row was the same as the next one. This is how I got the sum for the rows of the same stock value
         
             Else

                 (Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
         **Code for this part of the assignment was taken from activities we performed in class. Mainly used the Credit Card activity performed in class for this one.**
  
  5. Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". The solution should match the following image:
      - I started this off by declaring a few variables I would use for the max value, min value, and max volume value, and setting the variables to the cell I wanted them to start at
        
              Dim Max_Value As Double
              Max_Value = ws.Cells(2, 11).Value
    
              Dim Min_Value As Double
              Min_Value = ws.Cells(2, 11).Value
    
              Dim Max_Volume As Double
              Max_Volume = ws.Cells(2, 12).Value
        
      - I then created another for loop to iterate through my summary table looking for where the next row was greater than (for greatest increase and total volume), lesser than (for greatest decrease) and pasted that value into my 2nd summary table to highlight these three values. I also had to bring the ticker into the 2nd summary table.
        
            If ws.Cells(i + 1, 11).Value > Max_Value Then
        
              Max_Value = ws.Cells(i + 1, 11).Value
            
              ws.Cells(2, 17).Value = Max_Value
            
              ws.Cells(2, 16).Value = ws.Cells(i + 1, 9).Value
            
        'If Statement to return the greatest % decrease
        ElseIf ws.Cells(i + 1, 11).Value < Min_Value Then
        
                Min_Value = ws.Cells(i + 1, 11).Value
            
                ws.Cells(3, 17).Value = Min_Value
            
                ws.Cells(3, 16).Value = ws.Cells(i + 1, 9).Value
                
        End If
        
        'If Statement to return the greatest
        If ws.Cells(i + 1, 12).Value > Max_Volume Then
        
                Max_Volume = ws.Cells(i + 1, 12).Value
            
                ws.Cells(4, 17).Value = Max_Volume
            
                ws.Cells(4, 16).Value = ws.Cells(i + 1, 9).Value
        
        End If
        
      - This part stumped me for a while so I used an article I found on WallStreetMojo to start working through this part of the homework
      - (https://www.wallstreetmojo.com/vba-max/#:~:text=As%20the%20name%20suggests%2C%20Max,an%20array%20as%20an%20argument.)
        
      - That helped me get started on the code, but I couldn't get it to work using the information I found on that page. Eventually, I ended up on a page on stack overflow that was talking about Max and Min functions and provided some code for how it worked. Specifically, they started talking about using loops to iterate through the spreadsheet cell by cell. It was then I made the jump to checking if the value of the next cell was greater or lesser than the current cell and then keep performing that while iterating through each row, which is the route I ended up going. That way I would check each cell to see if it was greater or less than the previous cell, and if so, then I would set that as the new max or min value
      - Here's the stackoverflow link (https://stackoverflow.com/questions/45072650/finding-max-value-of-a-loop-with-vba)
      - I also used ChatGTP to test my code when I couldn't get it to work properly. ChatGPT helped clean up some of the syntax and get everything working in the proper format
  
  8. Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.
     - I looked at the class assignments for this one. The Credit Card activity specifically provides some code on how to iterate through each sheet in the workbook
     - I started with the following code to specify I wanted these changes done on each worksheet
    
           For Each ws In Worksheet
     - I also overlooked the fact that I needed to use ws. in front of each referenced cell or range. I came to this conclusion after running my code through chatgpt to figure out what was wrong with it.

  **Conditional Formatting**
  
  9. Highlight positive Yearly Change in green and negative change in red
      - For this part of the assignment I used the code we learned in class to highlight a cell a certain color. I paired that with an if statement within the for loop I created for the summary table to highlight the cells red or green depending on whether or not they were greater than or less than o
    
        'Statement to Color the cells depending on their value
        
            If ws.Cells(i, 10).Value > 0 Then
            
              'Set the Cell Colors to Green
              ws.Cells(i, 10).Interior.ColorIndex = 4
            
            ElseIf ws.Cells(i, 10).Value < 0 Then
        
              'Set the Cell Colors to Red
              ws.Cells(i, 10).Interior.ColorIndex = 3
            
          End If


  10. The assignment wasn't really clear on what was supposed to be done here, but I assumed we were supposed to use VBA to set the format of the percent change column to percent
        - For this assignment I originally tried to use similar code to what we learned in class as to how to set the currency, but I was unable to make that work. So I found some code online in order to turn the whole column to be formatting as percentage with 2 decimal places
     
                'Format the Percent Change Column to be Percentage with the desired decimal places
                ws.Columns("K").NumberFormat = "0.00%"
                ws.Range("Q2:Q3").NumberFormat = "0.00%"


        - I found this code on MrExcel.com (https://www.mrexcel.com/board/threads/vba-change-number-of-decimal-places-of-a-percentage.521221/)
  
    


**Resources Used**
