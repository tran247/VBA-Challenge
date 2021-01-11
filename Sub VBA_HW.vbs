Sub VBA_HW()

        'set cloumn headers
        Cells(1, 9).Value = "ticker"
        Cells(1, 10).Value = "yearly change"
        Cells(1, 11).Value = "percentage change"
        Cells(1, 12).Value = "total stock volume"
  
        
        'assign column headers a data type
        Dim tick_name As String
        Dim yearly_change As Double
        Dim percentage_change As Double
        Dim total_vol As Long
        
        
        
        
        'assign lastrow formula
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        
    
        
        'Start for loop
        For i = 2 To lastrow
            
            'create variable for counter for volume
            Dim volume_count As Long
            

            'set data type and variable for column
            Dim Column As Integer
            Column = 1
        

    
            'set counter as 0
            volume_count = 0
            
       
        
            'search if ticker is different
            If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
            
              'make formula for ticker volume
            volume_count = volume_count + Cells(i, 7).Value
            
                 'set ticker name as a variable
                tick_name = Cells(i, Column).Value
        
              'set column of cells "ticker" column to be similar to 1,1 column through variable
              Cells(2, 9).Value = tick_name
              
              'set total volume in volume cloumn
              Cells(2, 12).Value = total_vol
              
              'reset the total_vol after the total has been printed
              total_vol = 0
              
              
              
              
              ' I just can't seem to get past this part.
              
        
            
            End If
            
            Next i
            
            
        
        
        
        
        
        
        

End Sub

