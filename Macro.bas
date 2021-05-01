Attribute VB_Name = "Module1"
Sub Multipleyearstockdata()

Dim Ticker As String

Dim Summay_Table_Row As Long

Summary_Table_Row = 2

Dim Year_End_Price As Double
  
Dim Beginning_Open_Price As Double
  
Dim Yearly_Change As Double

Dim Percent_Change As Double

Dim Total_Volumn As Double

Last_Row = Cells(Rows.Count, 1).End(xlUp).Row




For i = 2 To Last_Row

    Range("I1").Value = "Ticker"
    
    Range("J1").Value = "Yearly Change"
    
    Range("K1").Value = "Percent Change"
    
    Range("L1").Value = "Total Volumn"
    
    Range("O2").Value = "Greatest % Increase"
    
    Range("O3").Value = "Greatest % Decrease"
    
    Range("O4").Value = "Greatest Total Volumn"
    
    Range("P1").Value = "Ticker"
    
    Range("Q1").Value = "Value"
    
    
    
    If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
    
       Beginning_Open_Price = Cells(i, 3).Value
       
       Total_Volumn = Cells(i, 7).Value
    
    
    
    ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    
       Ticker = Cells(i, 1).Value
    
       Year_End_Price = Cells(i, 6).Value
    
       Total_Volumn = Total_Volumn + Cells(i, 7).Value
    
       Yearly_Change = Year_End_Price - Beginning_Open_Price
    
       Percent_Change = (Year_End_Price - Beginning_Open_Price) / Beginning_Open_Price
    
    Columns("K").NumberFormat = "0.00%"

    Range("I" & Summary_Table_Row).Value = Ticker
    
    Range("J" & Summary_Table_Row).Value = Yearly_Change
    
    Range("K" & Summary_Table_Row).Value = Percent_Change
    
    Range("L" & Summary_Table_Row).Value = Total_Volumn
    
    Summary_Table_Row = Summary_Table_Row + 1
    
    
    Else
    
     Total_Volumn = Total_Volumn + Cells(i, 7).Value
     
    End If
    
    
    If Cells(i, 10).Value > 0 Then
    
    Cells(i, 10).Interior.ColorIndex = 4
    
    ElseIf Cells(i, 10).Value < 0 Then
    
    Cells(i, 10).Interior.ColorIndex = 3
    
    
    End If
    
    
 
    
   
    


Next i
   
   

   





End Sub



