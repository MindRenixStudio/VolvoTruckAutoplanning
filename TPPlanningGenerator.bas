Attribute VB_Name = "Module1"
Sub Button1_Plan()

    Dim InputCounter As Integer
    Dim TPCounter As Integer
    
    Dim SPZ As String
    Dim LN As String
    
    Dim SPZ1 As String
    Dim LN1 As String
    
    Dim TOs As String
    Dim TOs1 As String
    
    Dim CollInet As Integer
        
    Dim TPLineCounter As Integer
    
    Dim WordString As String
        
    TPCounter = 4
    
    TPLineCounter = 3

    '_____________________________________________

    'Creating TrailerList_Printable sheet
    Worksheets("TP_Template").Copy After:=Worksheets("Input")
    Sheets("TP_Template (2)").Name = "TrailerPlanning"

    'Processing of first line
    'Trailer Plate
    Worksheets("TrailerPlanning").Cells(3, 9) = Worksheets("Input").Cells(3, 16)
    'Time
    Worksheets("TrailerPlanning").Cells(3, 10) = Worksheets("Input").Cells(3, 17)
    'Supplier Country
    Worksheets("TrailerPlanning").Cells(3, 8) = Worksheets("Input").Cells(3, 7)
    'Carrier
    Worksheets("TrailerPlanning").Cells(3, 7) = Worksheets("Input").Cells(3, 9)
    'TO
    Worksheets("TrailerPlanning").Cells(3, 6) = Worksheets("Input").Cells(3, 2)
    'LN
    Worksheets("TrailerPlanning").Cells(3, 1) = Worksheets("Input").Cells(3, 1)
    'CollInet
    Worksheets("TrailerPlanning").Cells(3, 13) = Worksheets("Input").Cells(3, 12)
    
    For x = 3 To 500 'RowCount
        
        'CollInet
        CollInet = Worksheets("Input").Cells(TPCounter, 12)
        
        'SPZ
        SPZ = Worksheets("Input").Cells(TPCounter, 16)
        SPZ1 = Worksheets("Input").Cells(x, 16)
        
        'LN
        LN = Worksheets("Input").Cells(TPCounter, 1)
        LN1 = Worksheets("Input").Cells(x, 1)
        
        'TOs
        TOs = Worksheets("Input").Cells(TPCounter, 2)
        TOs1 = Worksheets("Input").Cells(x, 1)
                
        If SPZ <> SPZ1 Then
            'MsgBox ("nerovno")
            
            TPLineCounter = TPLineCounter + 1
            
            TPCounter = TPCounter + 1
            
            'Trailer Plate
            Worksheets("TrailerPlanning").Cells(TPLineCounter, 9) = Worksheets("Input").Cells(x + 1, 16)
            'Time
            Worksheets("TrailerPlanning").Cells(TPLineCounter, 10) = Worksheets("Input").Cells(x + 1, 17)
            'Supplier Country
            Worksheets("TrailerPlanning").Cells(TPLineCounter, 8) = Worksheets("Input").Cells(x + 1, 7)
            'Carrier
            Worksheets("TrailerPlanning").Cells(TPLineCounter, 7) = Worksheets("Input").Cells(x + 1, 9)
            'TOs
            Worksheets("TrailerPlanning").Cells(TPLineCounter, 6).Value = Worksheets("TrailerPlanning").Cells(TPLineCounter, 6) & " " & TOs
            'INET Colli
            Worksheets("TrailerPlanning").Cells(TPLineCounter, 13) = Worksheets("TrailerPlanning").Cells(TPLineCounter, 13).Value + CollInet
            
            If LN <> LN1 Then
                'LN
                Worksheets("TrailerPlanning").Cells(TPLineCounter, 1).Value = Worksheets("TrailerPlanning").Cells(TPLineCounter, 1) & "/" & LN
            End If
        
        ElseIf SPZ = SPZ1 Then
            
            'Trailer Plate
            Worksheets("TrailerPlanning").Cells(TPLineCounter, 9) = Worksheets("Input").Cells(x, 16)
            'Time
            Worksheets("TrailerPlanning").Cells(TPLineCounter, 10) = Worksheets("Input").Cells(x, 17)
            'Supplier Country
            Worksheets("TrailerPlanning").Cells(TPLineCounter, 8) = Worksheets("Input").Cells(x, 7)
            'Carrier
            Worksheets("TrailerPlanning").Cells(TPLineCounter, 7) = Worksheets("Input").Cells(x, 9)
            'TOs
            Worksheets("TrailerPlanning").Cells(TPLineCounter, 6).Value = Worksheets("TrailerPlanning").Cells(TPLineCounter, 6) & " " & TOs
            'INET Colli
            Worksheets("TrailerPlanning").Cells(TPLineCounter, 13) = Worksheets("TrailerPlanning").Cells(TPLineCounter, 13).Value + CollInet
            
            If LN <> LN1 Then
                'LN
                Worksheets("TrailerPlanning").Cells(TPLineCounter, 1).Value = Worksheets("TrailerPlanning").Cells(TPLineCounter, 1) & "/" & LN
            End If
                        
            If SPZ <> SPZ1 Then
                'MsgBox (LN1)
                
                'TPLineCounter = TPLineCounter + 1
            End If
                        
            TPCounter = TPCounter + 1
            
        ElseIf SPZ <> SPZ1 Then
            
            TPLineCounter = TPLineCounter + 1
        
        End If
        
    Next
    
    'CHECK FUNCTION FOR / and " "
    Dim LNCheck As String
    Dim TOCheck As String
    
    For x = 4 To 200
        
        LNCheck = Worksheets("TrailerPlanning").Cells(x, 1).Value
        Worksheets("TrailerPlanning").Cells(x, 1).Value = Replace(LNCheck, "/", "", 1, 1)
        
        TOCheck = Worksheets("TrailerPlanning").Cells(x, 6).Value
        Worksheets("TrailerPlanning").Cells(x, 6).Value = Replace(TOCheck, " ", "", 1, 1)
        
    Next
End Sub



















