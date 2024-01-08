Attribute VB_Name = "Module1"
Sub VBA_challenge_loop()

    ' Defining WS as a worksheet variable
    Dim WS As Worksheet
    
    ' Looping through each worksheet
    For Each WS In Worksheets
    
        ' Activating the worksheet
        WS.Activate
      
        ' Running the macro below on the worksheet
        Call VBA_challenge
    
    ' Moving on to the next worksheet
    Next WS

End Sub
