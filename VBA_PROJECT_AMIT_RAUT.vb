
' code to run the script on every shoot


Sub Stock_data()
    
    Dim xSh As Worksheet
    
    Application.ScreenUpdating = False
    
    For Each xSh In Worksheets
        xSh.Select
        
        Call Stockfile
    
    Next
    
    Application.ScreenUpdating = True

End Sub

'main script to run on everysheet.

Sub Stockfile()

'declare variable

 Dim lastrow As Double
 
 Dim column As Integer
 
 Dim tickernumberheading As Integer
 
 Dim totalvolume As Double
 
 Dim lastcolumn As Integer
 
 Dim yearstartprice As Double
 
 Dim yearendprice As Double
 
 Dim changeprice As Double
 
 Dim changepricepercent1 As Double
 
 Dim changepricepercent As Double
 
 Dim maximumyearlychange As Double
 

' initial variable

 totalvolume = 0
 lastcolumn = 0
 yearstartprice = 0
 yearendprice = 0
 changeprice = 0
 changepricepercent = 0
 
 
column = 1

tickernumberheading = 2

totalvolume = 0

lastrow = Cells(Rows.Count, 2).End(xlUp).Row

lastcolumn = Cells(1, Columns.Count).End(xlToLeft).column

'MsgBox (lastcolumn)

Range("I1").Value = "Ticker"

Range("j1").Value = "Yearly Change"
 
Range("k1").Value = "Percent Change"

Range("l1").Value = "Total Stock Volume"

yearstartprice = Cells(2, 6).Value
yearendprice = 0


'loop through last row with column 2

For i = 2 To lastrow

'check if two consecutive rows having differnt value


 
  If Cells(i + 1, column).Value <> Cells(i, column).Value Then
  
     yearendprice = Cells(i, 6).Value
     
     'code to find out the yearly price change of each individual stock
     
     changeprice = yearendprice - yearstartprice
     
     ' check if yearstartprice is 0
      
      If yearstartprice > 0 Then
      
     'code to find out the percentage change
     
         changepricepercent = (yearendprice - yearstartprice) / yearstartprice
     
       Else: changepricepercent = 0
       
       End If
       
     'code to find out the total volume
     
     totalvolume = totalvolume + Cells(i, lastcolumn).Value
     
     ' fill up the column with stock name Yearly Change  ,Percent Change , Total Stock Volume

     Cells(tickernumberheading, 9).Value = Cells(i, column).Value
     
     Cells(tickernumberheading, 10).Value = changeprice
     
     Cells(tickernumberheading, 11).Value = Format(changepricepercent, "Percent")
     
          
     Cells(tickernumberheading, 12).Value = totalvolume
          
     tickernumberheading = tickernumberheading + 1
     
     totalvolume = 0
     
     yearstartprice = Cells(i + 1, 6).Value
      
     
' add  volumes if the consecutive cells are not same

     Else: totalvolume = totalvolume + Cells(i, lastcolumn).Value
      
           yearstartprice = yearstartprice
            
        
        
         
   
     
   End If
    
     
Next i


 
 
 'change the color format red/green for percent update
 
 lastrow = Cells(Rows.Count, "J").End(xlUp).Row
 
 For i = 2 To lastrow
 
   If Cells(i, 10).Value < 0 Then
     
      Cells(i, 10).Interior.ColorIndex = 3
 
    Else: Cells(i, 10).Interior.ColorIndex = 4
    
    End If
    
  Next i
  
  




    
    
     
   
     
    
 
  
  



End Sub


