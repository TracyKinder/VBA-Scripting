Attribute VB_Name = "Module4"
Sub Thaw_RawData_Transform()
Attribute Thaw_RawData_Transform.VB_Description = "Take raw data and transform it into a format for a report based on fruit data received into cooler"
Attribute Thaw_RawData_Transform.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Thaw_RawData_Transform Macro
' Take raw data and transform it into a format for a report based on fruit data received into cooler
'
'______________________________________________________________________________
'Author:  Tracy R Kinder
' c is cell G3 that has the today() function that will be used as a reference to compare to the daily load plan date
' Two counters are used with a loop to transfer the desired keys over to the scratchpad
'Date 12/15/19
'
'_______________________________________________________________________________

' This how I unprotect the master Tabulation sheet in order to process new load planning
Sheet6.Unprotect "Odundbemru@123456"

Application.ScreenUpdating = False

'Declaration of Local Variables

Dim c As Date
Dim counter As Double
Dim destcount As Double
Dim rng As Range
Dim lastRow As Long

lastRow = Sheet2.Range("B7").End(xlDown).Row

'Setting variables to initial values

c = Sheet22.Range("E6").Value
counter = 14
destcount = 6
Set rng = Sheet2.Range("B7").CurrentRegion

'
'Only want to show dates that will have product in the rows by using todays date for thaw date
'______________________________________________________________________________________________________________

For counter = 14 To lastRow

If Sheet22.Range("C" & counter) = "" Then
Sheet22.Range("C" & counter) = Sheet22.Range("C" & counter).Offset(-1, 0)
Else

End If
Next counter

counter = 14
destcount = 6

For counter = 14 To lastRow
If c >= Sheet22.Range("A" & counter) Then
Sheet12.Range("A" & destcount) = Sheet22.Range("C" & counter)
Sheet12.Range("B" & destcount) = Sheet22.Range("D" & counter)

Sheet12.Range("C" & destcount) = Sheet22.Range("H" & counter)
Sheet12.Range("D" & destcount) = Sheet22.Range("P" & counter)
Sheet12.Range("E" & destcount) = Sheet22.Range("T" & counter)

destcount = destcount + 1
Else
End If
Next counter



    '________________________________________________________
    'For Full Loads only of a single FG #
    
    'counter = 14
    'destcount = 6
    
    'For counter = 14 To lastRow
    'If Sheet22.Range("B" & counter) = c And Sheet22.Range("B" & counter).Offset(0, 11) >= 1 Then
    'destcount = destcount + 1
    'Sheet12.Range("A" & destcount + 90) = Sheet22.Range("B" & counter).Offset(0, -1)
    'Else
    'End If
    'Next counter
    
    'It is a good practice to password protect the sheet in order to prevent tampering or inadvertent manipulation of information
    'located in the cells of the spreadsheet

Sheet6.Protect "Odundbemru@123456"
Application.ScreenUpdating = True
End Sub
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                