# VBA-Challenge_EyasuYilma
Challenge # 2
Sub Stock_Analysis()

' i is a row where the analysis starts and ends on the last row
Dim i As Double
Dim p As Double
Dim k As Double
Dim j As Double
Dim D As Double
Dim R As Integer
Dim Ticker As String
Dim Total_Stock_Volume As Double
'R is a row where the result appears and increases for every new result
R = 2
'D is a row where the new Ticker starts
D = 2
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row

         If Cells(i, 1) <> Cells(i + 1, 1) Then
         Ticker = Cells(i, 1)
        'Total_Stock_Volume
         Cells(R, 9) = Ticker
         'Yearly Change last row of the Ticker minus start a row of the Ticker
         Cells(R, 10) = (Cells(i, 6) - Cells(D, 3))
         'Percentage change yearly change divided by start value
          Cells(R, 11) = (Cells(R, 10) / Cells(D, 3)) * 100
          
          'Stock volume the sum of all volume by Ticker
            
            Cells(R, 12) = Application.WorksheetFunction.Sum(Range(Cells(D, 7), Cells(i, 7)))
           If Cells(R, 10) < 0 Then
           Cells(R, 10).Interior.ColorIndex = 3
           Else
           Cells(R, 10).Interior.ColorIndex = 4
           End If
        R = R + 1
        D = i + 1
        Else
                                  
      End If
      
    Next i
              
       'calculate the greatest increase, decrease, and volume
        Cells(2, 19) = Application.WorksheetFunction.Max(Range(Cells(2, 11), Cells(Cells(Rows.Count, 1).End(xlUp).Row, 11)))
        Cells(3, 19) = Application.WorksheetFunction.Min(Range(Cells(2, 11), Cells(Cells(Rows.Count, 1).End(xlUp).Row, 11)))
        Cells(4, 19) = Application.WorksheetFunction.Max(Range(Cells(2, 12), Cells(Cells(Rows.Count, 1).End(xlUp).Row, 12)))
        
        'This is a formula to find a Ticker that attributed to the values greatest increase, decrease, and volume
        
            For j = 2 To Cells(Rows.Count, 1).End(xlUp).Row
            If Cells(j, 11) = Cells(2, 19) Then
                Cells(2, 18) = Cells(j, 1)
                Else
                 End If
                 Next j
                 
            For p = 2 To Cells(Rows.Count, 1).End(xlUp).Row
            If Cells(p, 11) = Cells(3, 19) Then
                 Cells(3, 18) = Cells(p, 1)
                    Else
                    End If
                    Next p
                    
             For k = 2 To Cells(Rows.Count, 1).End(xlUp).Row
             If Cells(k, 12) = Cells(4, 19) Then
                Cells(4, 18) = Cells(k, 1)
                Else
                End If
                Next k
        
                       
End Sub

