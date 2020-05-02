Attribute VB_Name = "Module1"



Sub TransferToScratch()

'______________________________________________________________________________
  'Author:  Tracy R Kinder
  'c is cell G3 that has the today() function that will be used as reference  'to compare to the daily load plan date
  'Two counters are used with a loop to transfer the desired keys over to the 'scratchpad  Date 12/15/19
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

c = Sheet2.Range("G3").Value
counter = 7
destcount = 6
Set rng = Sheet2.Range("B7").CurrentRegion

'Main program to transfer Key #'s to scratchpad for daily load planning of partial FG # loads
' counters will have to be reprogramed from using hard code to lower and upper bounds in the for next loop
'______________________________________________________________________________________________________________

For counter = 7 To lastRow
If c = Sheet2.Range("B" & counter) Then
destcount = destcount + 1
Sheet6.Range("A" & destcount) = Sheet2.Range("B" & counter).Offset(0, -1)
Else
End If
Next counter

'________________________________________________________
'For Full Loads only of a single FG #

counter = 7
destcount = 6

For counter = 7 To lastRow
If Sheet2.Range("B" & counter) = c And Sheet2.Range("B" & counter).Offset(0, 11) >= 1 Then
destcount = destcount + 1
Sheet6.Range("A" & destcount + 90) = Sheet2.Range("B" & counter).Offset(0, -1)
Else
End If
Next counter

'It is a good practice to password protect the sheet in order to prevent tampering or inadvertent manipulation of information
'located in the cells of the spreadsheet

Sheet6.Protect "Odundbemru@123456"
Application.ScreenUpdating = True

End Sub

Sub ScratchPad()


'__________________________________________________________________________________________________________________________________
'Author:  Tracy R Kinder
' Date 12/15/2019
' This program will load the FG #'s by key into each partial load with utililzing a math function to maximize each load
' Also, each full load will be automatically loaded into the full load section near the bottom
'
'__________________________________________________________________________________________________________________________________
 
'Unlock protected Sheet

Sheet6.Unprotect "Odundbemru@123456"
Application.ScreenUpdating = False

' >> Define local variables <<

Dim lastRow As Long
Dim PlannedProduct() As Variant
Dim i As Double
Dim j As Double
Dim n As Double

' Dynamic arrays must have a way to find matrix dimensions each time the code is run

lastRow = Sheet6.Range("A7").End(xlDown).Row
lastCol = Sheet6.Range("A7").End(xlToRight).Column
n = 7

' Dynamic Arrays must be redimensioned after each successful run of code
' Use the Preserve command to keep your data from being erased after a successful run of code

ReDim Preserve PlannedProduct(0 To lastRow, 1 To lastCol)

'>> Assign Values to the Planned Product array <<
' arrary when i = 12 is the pallet number and i = 14 is pallet weight
For i = 7 To lastRow
    For j = 1 To lastCol
        PlannedProduct(i, j) = Sheet6.Cells(i, j).Value
        Debug.Print PlannedProduct(i, j)
    Next j
Next i

'>> Now that you have a working dynamic array filled with data, it needs to be manipulated in such a way as to fill individual
'     loads with an exact weight, and only using 26 pallets or under.  Use a math function and pass variables for maximing
'     full loads <<

Call Math(PlannedProduct)
Debug.Print "Hey this is the main scratchpad routine just post the call math statement "

'It is a good practice to password protect the sheet in order to prevent tampering or inadvertent manipulation of information
'located in the cells of the spreadsheet

Sheet6.Protect "Odundbemru@123456"
Application.ScreenUpdating = True

End Sub

Sub Math(ByRef PlannedProductArray() As Variant)
'>>> This is a subroutine called Math that will manipulate the data from the Array before going back to finish the scratchpad
'       routine.

'Variable Declaration

Dim AlreadyDone As Integer
Dim destcount As Integer
Dim i As Integer

Debug.Print "Hey this is the First Math routine saying hello"
  
 'Initialize variables
 
  AlreadyDone = 0
  destcount = 7
   i = 7
   
   'Main Sorting Algorithm Do While Loop
   
   Do While True
   
   If PlannedProductArray(i, 6) < 26 And PlannedProductArray(i, 8) < 44500 Then
   Debug.Print "Key 1 needs to be transferred to Load 1"
   Sheet6.Range("J" & destcount).Value = PlannedProductArray(i, 1)
   Else:
   MsgBox "This is the first FG and should not be a full load"
   Exit Do
   End If
   
        If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) < 44500 Then
        Debug.Print "Next Key needs to be transferred to Load 1"
        Sheet6.Range("J" & destcount + 1).Value = PlannedProductArray((i + 1), 1)
        Else:
        MsgBox "This is the second FG that was added to previous FG and it did not meet the requirements to be added to load 1"
        Exit Do
        End If
   
            If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) + PlannedProductArray(i + 2, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) + PlannedProductArray(i + 2, 8) < 44500 Then
            Debug.Print "Next Key needs to be transferred to Load 1"
            Sheet6.Range("J" & destcount + 2).Value = PlannedProductArray((i + 2), 1)
            Else:
            MsgBox "This is the third FG that was added to previous FG and it did not meet the requirements to be added to load 1"
            Exit Do
            End If
        
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) + PlannedProductArray(i + 2, 6) + PlannedProductArray(i + 3, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) + PlannedProductArray(i + 2, 8) + PlannedProductArray(i + 3, 8) < 44500 Then
   Debug.Print "Next Key needs to be transferred to Load 1"
   Sheet6.Range("J" & destcount + 3).Value = PlannedProductArray((i + 3), 1)
   Else:
   MsgBox "This is the fourth FG reviewed but does not meet the requirements to be added to load 1.  Exiting 1st Math Routine"
   Exit Do
   End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) + PlannedProductArray(i + 2, 6) + PlannedProductArray(i + 3, 6) + PlannedProductArray(i + 4, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) + PlannedProductArray(i + 2, 8) + PlannedProductArray(i + 3, 8) + PlannedProductArray(i + 4, 8) < 44500 Then
   Debug.Print "Next Key needs to be transferred to Load 1"
   Sheet6.Range("J" & destcount + 4).Value = PlannedProductArray((i + 4), 1)
   Else:
   MsgBox "This is the fifth FG that was added to previous FG and it did not meet the requirements to be added to load 1"
   Exit Do
   End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) + PlannedProductArray(i + 2, 6) + PlannedProductArray(i + 3, 6) + PlannedProductArray(i + 4, 6) + PlannedProductArray(i + 5, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) + PlannedProductArray(i + 2, 8) + PlannedProductArray(i + 3, 8) + PlannedProductArray(i + 4, 8) + PlannedProductArray(i + 5, 8) < 44500 Then
   Debug.Print "Next Key needs to be transferred to Load 1"
   Sheet6.Range("J" & destcount + 5).Value = PlannedProductArray((i + 5), 1)
   Else:
   MsgBox "This is the sixth FG that was added to previous FG and it did not meet the requirements to be added to load 1"
   Exit Do
   End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) + PlannedProductArray(i + 2, 6) + PlannedProductArray(i + 3, 6) + PlannedProductArray(i + 4, 6) + PlannedProductArray(i + 5, 6) + PlannedProductArray(i + 6, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) + PlannedProductArray(i + 2, 8) + PlannedProductArray(i + 3, 8) + PlannedProductArray(i + 4, 8) + PlannedProductArray(i + 5, 8) + PlannedProductArray(i + 6, 8) < 44500 Then
   Debug.Print "Next Key needs to be transferred to Load 1"
   Sheet6.Range("J" & destcount + 6).Value = PlannedProductArray((i + 6), 1)
   Else:
   MsgBox "This is the seventh FG that was added to previous FG and it did not meet the requirements to be added to load 1"
   Exit Do
   End If
   Exit Do
   Loop
   i = i + 1
   Debug.Print "Hey the math1 routine finished and calling Math1 to begin Load 2"
   Call Math1(PlannedProductArray)
   
   
   
End Sub

Sub Math1(ByRef PlannedProductArray() As Variant)
'>> This Second Math Routine is to ensure each load is full as possible <<
'     This is for Load 2
       
                
 'This is where variable AlreadyDone comes into play.  I cannot view AlreadyDone variable from the previous routine from which this one is called.
 'so I have to view what was in the array that has been saved.  Then I add what was AlreadyDone to the i variable to keep from repeated
 'product that has already been planned to ship
 
 
                destcount = 7
                AlreadyDone = 0
            
              'For each Load or Math Routine, another DO WHILE loop is need to add all AlreadyDone keys
              
               Do While Sheet6.Range("J" & destcount).Value <> 0
               destcount = destcount + 1
               AlreadyDone = (AlreadyDone) + 1
               Loop

Do While True
i = 7 + (AlreadyDone)
destcount = 27
    If PlannedProductArray(i, 6) < 26 And PlannedProductArray(i, 8) < 44500 Then
    Debug.Print "Key" & destcount; " needs to be transferred to Load 1"
    Sheet6.Range("J" & destcount).Value = PlannedProductArray(i, 1)
    Else:
    MsgBox "This should not be considered a partial load"
    Exit Do
    End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) < 44500 Then
   Debug.Print "Key 2 needs to be transferred to Load 2"
   Sheet6.Range("J" & destcount + 1).Value = PlannedProductArray((i + 1), 1)
   Else:
   MsgBox "This is the second FG that was added to previous FG and it did not meet the requirements to be added to load 2"
   Exit Do
   End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) + PlannedProductArray(i + 2, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) + PlannedProductArray(i + 2, 8) < 44500 Then
   Debug.Print "Next Key needs to be transferred to Load 2"
   Sheet6.Range("J" & destcount + 2).Value = PlannedProductArray((i + 2), 1)
   Else:
   MsgBox "This is the third FG that was added to previous FG and it did not meet the requirements to be added to load 1"

   Exit Do
   End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) + PlannedProductArray(i + 2, 6) + PlannedProductArray(i + 3, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) + PlannedProductArray(i + 2, 8) + PlannedProductArray(i + 3, 8) < 44500 Then
   Debug.Print "Next Key needs to be transferred to Load 2"
   Sheet6.Range("J" & destcount + 3).Value = PlannedProductArray((i + 3), 1)
   Else:
   MsgBox "This is the fourth FG that was added to previous FG and it did not meet the requirements to be added to load 2"
   Exit Do
   End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) + PlannedProductArray(i + 2, 6) + PlannedProductArray(i + 3, 6) + PlannedProductArray(i + 4, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) + PlannedProductArray(i + 2, 8) + PlannedProductArray(i + 3, 8) + PlannedProductArray(i + 4, 8) < 44500 Then
   Debug.Print "Next Key needs to be transferred to Load 2"
   Sheet6.Range("J" & destcount + 4).Value = PlannedProductArray((i + 4), 1)
   Else:
   MsgBox "This is the fifth FG that was added to previous FG and it did not meet the requirements to be added to load 2"
   Exit Do
   End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) + PlannedProductArray(i + 2, 6) + PlannedProductArray(i + 3, 6) + PlannedProductArray(i + 4, 6) + PlannedProductArray(i + 5, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) + PlannedProductArray(i + 2, 8) + PlannedProductArray(i + 3, 8) + PlannedProductArray(i + 4, 8) + PlannedProductArray(i + 5, 8) < 44500 Then
   Debug.Print "Next Key needs to be transferred to Load 2"
   Sheet6.Range("J" & destcount + 5).Value = PlannedProductArray((i + 5), 1)
   Else:
   MsgBox "This is the sixth FG that was added to previous FG and it did not meet the requirements to be added to load 2"
   Exit Do
   End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) + PlannedProductArray(i + 2, 6) + PlannedProductArray(i + 3, 6) + PlannedProductArray(i + 4, 6) + PlannedProductArray(i + 5, 6) + PlannedProductArray(i + 6, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) + PlannedProductArray(i + 2, 8) + PlannedProductArray(i + 3, 8) + PlannedProductArray(i + 4, 8) + PlannedProductArray(i + 5, 8) + PlannedProductArray(i + 6, 8) < 44500 Then
   Debug.Print "Next Key needs to be transferred to Load 2"
   Sheet6.Range("J" & destcount + 6).Value = PlannedProductArray((i + 6), 1)
   Else:
   MsgBox "This is the seventh FG that was added to previous FG and it did not meet the requirements to be added to load 2"
   Exit Do
   End If
     
   Exit Do
   Loop
   
   i = i + 1
   Debug.Print "Hey the math1 routine finished and calling Math 2 to begin Load 3"
   Call Math2(PlannedProductArray)

End Sub

Sub Math2(ByRef PlannedProductArray() As Variant)
'>> This Second Math Routine is to ensure each load is full as possible <<
'     This is for Load 3
             
                destcount = 7
                AlreadyDone = 0
            
              'For each Load or Math Routine, another DO WHILE loop is need to add all AlreadyDone keys
              
               Do While Sheet6.Range("J" & destcount).Value <> 0
               destcount = destcount + 1
               AlreadyDone = (AlreadyDone) + 1
               Loop
               
               destcount = 27
               Do While Sheet6.Range("J" & destcount).Value <> 0
               destcount = destcount + 1
               AlreadyDone = (AlreadyDone) + 1
               Loop

Do While True
i = 7 + (AlreadyDone)
destcount = 49
    If PlannedProductArray(i, 6) < 26 And PlannedProductArray(i, 8) < 44500 Then
    Debug.Print "Key" & destcount; " needs to be transferred to Load 1"
    Sheet6.Range("J" & destcount).Value = PlannedProductArray(i, 1)
    Else:
    MsgBox "This should not be considered a partial load"
    Exit Do
    End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) < 44500 Then
   Debug.Print "Key 2 needs to be transferred to Load 2"
   Sheet6.Range("J" & destcount + 1).Value = PlannedProductArray((i + 1), 1)
   Else:
   MsgBox "This is the second FG that was added to previous FG and it did not meet the requirements to be added to load 2"
   Exit Do
   End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) + PlannedProductArray(i + 2, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) + PlannedProductArray(i + 2, 8) < 44500 Then
   Debug.Print "Next Key needs to be transferred to Load 2"
   Sheet6.Range("J" & destcount + 2).Value = PlannedProductArray((i + 2), 1)
   Else:
   MsgBox "This is the third FG that was added to previous FG and it did not meet the requirements to be added to load 1"

   Exit Do
   End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) + PlannedProductArray(i + 2, 6) + PlannedProductArray(i + 3, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) + PlannedProductArray(i + 2, 8) + PlannedProductArray(i + 3, 8) < 44500 Then
   Debug.Print "Next Key needs to be transferred to Load 2"
   Sheet6.Range("J" & destcount + 3).Value = PlannedProductArray((i + 3), 1)
   Else:
   MsgBox "This is the fourth FG that was added to previous FG and it did not meet the requirements to be added to load 2"
   Exit Do
   End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) + PlannedProductArray(i + 2, 6) + PlannedProductArray(i + 3, 6) + PlannedProductArray(i + 4, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) + PlannedProductArray(i + 2, 8) + PlannedProductArray(i + 3, 8) + PlannedProductArray(i + 4, 8) < 44500 Then
   Debug.Print "Next Key needs to be transferred to Load 2"
   Sheet6.Range("J" & destcount + 4).Value = PlannedProductArray((i + 4), 1)
   Else:
   MsgBox "This is the fifth FG that was added to previous FG and it did not meet the requirements to be added to load 2"
   Exit Do
   End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) + PlannedProductArray(i + 2, 6) + PlannedProductArray(i + 3, 6) + PlannedProductArray(i + 4, 6) + PlannedProductArray(i + 5, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) + PlannedProductArray(i + 2, 8) + PlannedProductArray(i + 3, 8) + PlannedProductArray(i + 4, 8) + PlannedProductArray(i + 5, 8) < 44500 Then
   Debug.Print "Next Key needs to be transferred to Load 2"
   Sheet6.Range("J" & destcount + 5).Value = PlannedProductArray((i + 5), 1)
   Else:
   MsgBox "This is the sixth FG that was added to previous FG and it did not meet the requirements to be added to load 2"
   Exit Do
   End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) + PlannedProductArray(i + 2, 6) + PlannedProductArray(i + 3, 6) + PlannedProductArray(i + 4, 6) + PlannedProductArray(i + 5, 6) + PlannedProductArray(i + 6, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) + PlannedProductArray(i + 2, 8) + PlannedProductArray(i + 3, 8) + PlannedProductArray(i + 4, 8) + PlannedProductArray(i + 5, 8) + PlannedProductArray(i + 6, 8) < 44500 Then
   Debug.Print "Next Key needs to be transferred to Load 2"
   Sheet6.Range("J" & destcount + 6).Value = PlannedProductArray((i + 6), 1)
   Else:
   MsgBox "This is the seventh FG that was added to previous FG and it did not meet the requirements to be added to load 2"
   Exit Do
   End If
     
   Exit Do
   Loop
   
   i = i + 1
Debug.Print "Hey the math2 routine finished and this is back in the math routine"
Call Math3(PlannedProductArray)
End Sub

Sub Math3(ByRef PlannedProductArray() As Variant)
'This is for Load 4
                destcount = 7
                AlreadyDone = 0
            
              'For each Load or Math Routine, another DO WHILE loop is need to add all AlreadyDone keys
              
               Do While Sheet6.Range("J" & destcount).Value <> 0
               destcount = destcount + 1
               AlreadyDone = (AlreadyDone) + 1
               Loop
               
               destcount = 27
               Do While Sheet6.Range("J" & destcount).Value <> 0
               destcount = destcount + 1
               AlreadyDone = (AlreadyDone) + 1
               Loop
               
               destcount = 49
               Do While Sheet6.Range("J" & destcount).Value <> 0
               destcount = destcount + 1
               AlreadyDone = (AlreadyDone) + 1
               Loop

Do While True
i = 7 + (AlreadyDone)
destcount = 70
    If PlannedProductArray(i, 6) < 26 And PlannedProductArray(i, 8) < 44500 Then
    Debug.Print "Key" & destcount; " needs to be transferred to Load 1"
    Sheet6.Range("J" & destcount).Value = PlannedProductArray(i, 1)
    Else:
    MsgBox "This should not be considered a partial load"
    Exit Do
    End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) < 44500 Then
   Debug.Print "Key 2 needs to be transferred to Load 2"
   Sheet6.Range("J" & destcount + 1).Value = PlannedProductArray((i + 1), 1)
   Else:
   MsgBox "This is the second FG that was added to previous FG and it did not meet the requirements to be added to load 2"
   Exit Do
   End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) + PlannedProductArray(i + 2, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) + PlannedProductArray(i + 2, 8) < 44500 Then
   Debug.Print "Next Key needs to be transferred to Load 2"
   Sheet6.Range("J" & destcount + 2).Value = PlannedProductArray((i + 2), 1)
   Else:
   MsgBox "This is the third FG that was added to previous FG and it did not meet the requirements to be added to load 1"

   Exit Do
   End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) + PlannedProductArray(i + 2, 6) + PlannedProductArray(i + 3, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) + PlannedProductArray(i + 2, 8) + PlannedProductArray(i + 3, 8) < 44500 Then
   Debug.Print "Next Key needs to be transferred to Load 2"
   Sheet6.Range("J" & destcount + 3).Value = PlannedProductArray((i + 3), 1)
   Else:
   MsgBox "This is the fourth FG that was added to previous FG and it did not meet the requirements to be added to load 2"
   Exit Do
   End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) + PlannedProductArray(i + 2, 6) + PlannedProductArray(i + 3, 6) + PlannedProductArray(i + 4, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) + PlannedProductArray(i + 2, 8) + PlannedProductArray(i + 3, 8) + PlannedProductArray(i + 4, 8) < 44500 Then
   Debug.Print "Next Key needs to be transferred to Load 2"
   Sheet6.Range("J" & destcount + 4).Value = PlannedProductArray((i + 4), 1)
   Else:
   MsgBox "This is the fifth FG that was added to previous FG and it did not meet the requirements to be added to load 2"
   Exit Do
   End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) + PlannedProductArray(i + 2, 6) + PlannedProductArray(i + 3, 6) + PlannedProductArray(i + 4, 6) + PlannedProductArray(i + 5, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) + PlannedProductArray(i + 2, 8) + PlannedProductArray(i + 3, 8) + PlannedProductArray(i + 4, 8) + PlannedProductArray(i + 5, 8) < 44500 Then
   Debug.Print "Next Key needs to be transferred to Load 2"
   Sheet6.Range("J" & destcount + 5).Value = PlannedProductArray((i + 5), 1)
   Else:
   MsgBox "This is the sixth FG that was added to previous FG and it did not meet the requirements to be added to load 2"
   Exit Do
   End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) + PlannedProductArray(i + 2, 6) + PlannedProductArray(i + 3, 6) + PlannedProductArray(i + 4, 6) + PlannedProductArray(i + 5, 6) + PlannedProductArray(i + 6, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) + PlannedProductArray(i + 2, 8) + PlannedProductArray(i + 3, 8) + PlannedProductArray(i + 4, 8) + PlannedProductArray(i + 5, 8) + PlannedProductArray(i + 6, 8) < 44500 Then
   Debug.Print "Next Key needs to be transferred to Load 2"
   Sheet6.Range("J" & destcount + 6).Value = PlannedProductArray((i + 6), 1)
   Else:
   MsgBox "This is the seventh FG that was added to previous FG and it did not meet the requirements to be added to load 2"
   Exit Do
   End If
     
   Exit Do
   Loop
   
   i = i + 1

Call Math4(PlannedProductArray)
'Debug.Print "Hey the math3 routine finished and this is back in the math routine"
End Sub

Sub Math4(ByRef PlannedProductArray() As Variant)

                destcount = 7
                AlreadyDone = 0
            
              'For each Load or Math Routine, another DO WHILE loop is need to add all AlreadyDone keys
              
               Do While Sheet6.Range("J" & destcount).Value <> 0
               destcount = destcount + 1
               AlreadyDone = (AlreadyDone) + 1
               Loop
               
               destcount = 27
               Do While Sheet6.Range("J" & destcount).Value <> 0
               destcount = destcount + 1
               AlreadyDone = (AlreadyDone) + 1
               Loop
               
               destcount = 49
               Do While Sheet6.Range("J" & destcount).Value <> 0
               destcount = destcount + 1
               AlreadyDone = (AlreadyDone) + 1
               Loop
               
               destcount = 70
               Do While Sheet6.Range("J" & destcount).Value <> 0
               destcount = destcount + 1
               AlreadyDone = (AlreadyDone) + 1
               Loop

Do While True
i = 7 + (AlreadyDone)
destcount = 92
    If PlannedProductArray(i, 6) < 26 And PlannedProductArray(i, 8) < 44500 Then
    Debug.Print "Key" & destcount; " needs to be transferred to Load 1"
    Sheet6.Range("J" & destcount).Value = PlannedProductArray(i, 1)
    Else:
    MsgBox "This should not be considered a partial load"
    Exit Do
    End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) < 44500 Then
   Debug.Print "Key 2 needs to be transferred to Load 2"
   Sheet6.Range("J" & destcount + 1).Value = PlannedProductArray((i + 1), 1)
   Else:
   MsgBox "This is the second FG that was added to previous FG and it did not meet the requirements to be added to load 2"
   Exit Do
   End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) + PlannedProductArray(i + 2, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) + PlannedProductArray(i + 2, 8) < 44500 Then
   Debug.Print "Next Key needs to be transferred to Load 2"
   Sheet6.Range("J" & destcount + 2).Value = PlannedProductArray((i + 2), 1)
   Else:
   MsgBox "This is the third FG that was added to previous FG and it did not meet the requirements to be added to load 1"

   Exit Do
   End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) + PlannedProductArray(i + 2, 6) + PlannedProductArray(i + 3, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) + PlannedProductArray(i + 2, 8) + PlannedProductArray(i + 3, 8) < 44500 Then
   Debug.Print "Next Key needs to be transferred to Load 2"
   Sheet6.Range("J" & destcount + 3).Value = PlannedProductArray((i + 3), 1)
   Else:
   MsgBox "This is the fourth FG that was added to previous FG and it did not meet the requirements to be added to load 2"
   Exit Do
   End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) + PlannedProductArray(i + 2, 6) + PlannedProductArray(i + 3, 6) + PlannedProductArray(i + 4, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) + PlannedProductArray(i + 2, 8) + PlannedProductArray(i + 3, 8) + PlannedProductArray(i + 4, 8) < 44500 Then
   Debug.Print "Next Key needs to be transferred to Load 2"
   Sheet6.Range("J" & destcount + 4).Value = PlannedProductArray((i + 4), 1)
   Else:
   MsgBox "This is the fifth FG that was added to previous FG and it did not meet the requirements to be added to load 2"
   Exit Do
   End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) + PlannedProductArray(i + 2, 6) + PlannedProductArray(i + 3, 6) + PlannedProductArray(i + 4, 6) + PlannedProductArray(i + 5, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) + PlannedProductArray(i + 2, 8) + PlannedProductArray(i + 3, 8) + PlannedProductArray(i + 4, 8) + PlannedProductArray(i + 5, 8) < 44500 Then
   Debug.Print "Next Key needs to be transferred to Load 2"
   Sheet6.Range("J" & destcount + 5).Value = PlannedProductArray((i + 5), 1)
   Else:
   MsgBox "This is the sixth FG that was added to previous FG and it did not meet the requirements to be added to load 2"
   Exit Do
   End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) + PlannedProductArray(i + 2, 6) + PlannedProductArray(i + 3, 6) + PlannedProductArray(i + 4, 6) + PlannedProductArray(i + 5, 6) + PlannedProductArray(i + 6, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) + PlannedProductArray(i + 2, 8) + PlannedProductArray(i + 3, 8) + PlannedProductArray(i + 4, 8) + PlannedProductArray(i + 5, 8) + PlannedProductArray(i + 6, 8) < 44500 Then
   Debug.Print "Next Key needs to be transferred to Load 2"
   Sheet6.Range("J" & destcount + 6).Value = PlannedProductArray((i + 6), 1)
   Else:
   MsgBox "This is the seventh FG that was added to previous FG and it did not meet the requirements to be added to load 2"
   Exit Do
   End If
     
   Exit Do
   Loop
   
   i = i + 1
Call Math5(PlannedProductArray)
End Sub

Sub Math5(ByRef PlannedProductArray() As Variant)

                destcount = 7
                AlreadyDone = 0
            
              'For each Load or Math Routine, another DO WHILE loop is need to add all AlreadyDone keys
              
               Do While Sheet6.Range("J" & destcount).Value <> 0
               destcount = destcount + 1
               AlreadyDone = (AlreadyDone) + 1
               Loop
               
               destcount = 27
               Do While Sheet6.Range("J" & destcount).Value <> 0
               destcount = destcount + 1
               AlreadyDone = (AlreadyDone) + 1
               Loop
               
               destcount = 49
               Do While Sheet6.Range("J" & destcount).Value <> 0
               destcount = destcount + 1
               AlreadyDone = (AlreadyDone) + 1
               Loop
               
               destcount = 70
               Do While Sheet6.Range("J" & destcount).Value <> 0
               destcount = destcount + 1
               AlreadyDone = (AlreadyDone) + 1
               Loop
               
               destcount = 92
               Do While Sheet6.Range("J" & destcount).Value <> 0
               destcount = destcount + 1
               AlreadyDone = (AlreadyDone) + 1
               Loop


Do While True
i = 7 + (AlreadyDone)
destcount = 113
    If PlannedProductArray(i, 6) < 26 And PlannedProductArray(i, 8) < 44500 Then
    Debug.Print "Key" & destcount; " needs to be transferred to Load 1"
    Sheet6.Range("J" & destcount).Value = PlannedProductArray(i, 1)
    Else:
    MsgBox "This should not be considered a partial load"
    Exit Do
    End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) < 44500 Then
   Debug.Print "Key 2 needs to be transferred to Load 2"
   Sheet6.Range("J" & destcount + 1).Value = PlannedProductArray((i + 1), 1)
   Else:
   MsgBox "This is the second FG that was added to previous FG and it did not meet the requirements to be added to load 2"
   Exit Do
   End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) + PlannedProductArray(i + 2, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) + PlannedProductArray(i + 2, 8) < 44500 Then
   Debug.Print "Next Key needs to be transferred to Load 2"
   Sheet6.Range("J" & destcount + 2).Value = PlannedProductArray((i + 2), 1)
   Else:
   MsgBox "This is the third FG that was added to previous FG and it did not meet the requirements to be added to load 1"

   Exit Do
   End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) + PlannedProductArray(i + 2, 6) + PlannedProductArray(i + 3, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) + PlannedProductArray(i + 2, 8) + PlannedProductArray(i + 3, 8) < 44500 Then
   Debug.Print "Next Key needs to be transferred to Load 2"
   Sheet6.Range("J" & destcount + 3).Value = PlannedProductArray((i + 3), 1)
   Else:
   MsgBox "This is the fourth FG that was added to previous FG and it did not meet the requirements to be added to load 2"
   Exit Do
   End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) + PlannedProductArray(i + 2, 6) + PlannedProductArray(i + 3, 6) + PlannedProductArray(i + 4, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) + PlannedProductArray(i + 2, 8) + PlannedProductArray(i + 3, 8) + PlannedProductArray(i + 4, 8) < 44500 Then
   Debug.Print "Next Key needs to be transferred to Load 2"
   Sheet6.Range("J" & destcount + 4).Value = PlannedProductArray((i + 4), 1)
   Else:
   MsgBox "This is the fifth FG that was added to previous FG and it did not meet the requirements to be added to load 2"
   Exit Do
   End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) + PlannedProductArray(i + 2, 6) + PlannedProductArray(i + 3, 6) + PlannedProductArray(i + 4, 6) + PlannedProductArray(i + 5, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) + PlannedProductArray(i + 2, 8) + PlannedProductArray(i + 3, 8) + PlannedProductArray(i + 4, 8) + PlannedProductArray(i + 5, 8) < 44500 Then
   Debug.Print "Next Key needs to be transferred to Load 2"
   Sheet6.Range("J" & destcount + 5).Value = PlannedProductArray((i + 5), 1)
   Else:
   MsgBox "This is the sixth FG that was added to previous FG and it did not meet the requirements to be added to load 2"
   Exit Do
   End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) + PlannedProductArray(i + 2, 6) + PlannedProductArray(i + 3, 6) + PlannedProductArray(i + 4, 6) + PlannedProductArray(i + 5, 6) + PlannedProductArray(i + 6, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) + PlannedProductArray(i + 2, 8) + PlannedProductArray(i + 3, 8) + PlannedProductArray(i + 4, 8) + PlannedProductArray(i + 5, 8) + PlannedProductArray(i + 6, 8) < 44500 Then
   Debug.Print "Next Key needs to be transferred to Load 2"
   Sheet6.Range("J" & destcount + 6).Value = PlannedProductArray((i + 6), 1)
   Else:
   MsgBox "This is the seventh FG that was added to previous FG and it did not meet the requirements to be added to load 2"
   Exit Do
   End If
     
   Exit Do
   Loop
   
   i = i + 1
Call Math6(PlannedProductArray)
End Sub
Sub Math6(ByRef PlannedProductArray() As Variant)

                destcount = 7
                AlreadyDone = 0
            
              'For each Load or Math Routine, another DO WHILE loop is need to add all AlreadyDone keys
              
               Do While Sheet6.Range("J" & destcount).Value <> 0
               destcount = destcount + 1
               AlreadyDone = (AlreadyDone) + 1
               Loop
               
               destcount = 27
               Do While Sheet6.Range("J" & destcount).Value <> 0
               destcount = destcount + 1
               AlreadyDone = (AlreadyDone) + 1
               Loop
               
               destcount = 49
               Do While Sheet6.Range("J" & destcount).Value <> 0
               destcount = destcount + 1
               AlreadyDone = (AlreadyDone) + 1
               Loop
               
               destcount = 70
               Do While Sheet6.Range("J" & destcount).Value <> 0
               destcount = destcount + 1
               AlreadyDone = (AlreadyDone) + 1
               Loop
               
               destcount = 92
               Do While Sheet6.Range("J" & destcount).Value <> 0
               destcount = destcount + 1
               AlreadyDone = (AlreadyDone) + 1
               Loop
               
               destcount = 113
               Do While Sheet6.Range("J" & destcount).Value <> 0
               destcount = destcount + 1
               AlreadyDone = (AlreadyDone) + 1
               Loop


Do While True
i = 7 + (AlreadyDone)
destcount = 135
    If PlannedProductArray(i, 6) < 26 And PlannedProductArray(i, 8) < 44500 Then
    Debug.Print "Key" & destcount; " needs to be transferred to Load 1"
    Sheet6.Range("J" & destcount).Value = PlannedProductArray(i, 1)
    Else:
    MsgBox "This should not be considered a partial load"
    Exit Do
    End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) < 44500 Then
   Debug.Print "Key 2 needs to be transferred to Load 2"
   Sheet6.Range("J" & destcount + 1).Value = PlannedProductArray((i + 1), 1)
   Else:
   MsgBox "This is the second FG that was added to previous FG and it did not meet the requirements to be added to load 2"
   Exit Do
   End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) + PlannedProductArray(i + 2, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) + PlannedProductArray(i + 2, 8) < 44500 Then
   Debug.Print "Next Key needs to be transferred to Load 2"
   Sheet6.Range("J" & destcount + 2).Value = PlannedProductArray((i + 2), 1)
   Else:
   MsgBox "This is the third FG that was added to previous FG and it did not meet the requirements to be added to load 1"

   Exit Do
   End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) + PlannedProductArray(i + 2, 6) + PlannedProductArray(i + 3, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) + PlannedProductArray(i + 2, 8) + PlannedProductArray(i + 3, 8) < 44500 Then
   Debug.Print "Next Key needs to be transferred to Load 2"
   Sheet6.Range("J" & destcount + 3).Value = PlannedProductArray((i + 3), 1)
   Else:
   MsgBox "This is the fourth FG that was added to previous FG and it did not meet the requirements to be added to load 2"
   Exit Do
   End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) + PlannedProductArray(i + 2, 6) + PlannedProductArray(i + 3, 6) + PlannedProductArray(i + 4, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) + PlannedProductArray(i + 2, 8) + PlannedProductArray(i + 3, 8) + PlannedProductArray(i + 4, 8) < 44500 Then
   Debug.Print "Next Key needs to be transferred to Load 2"
   Sheet6.Range("J" & destcount + 4).Value = PlannedProductArray((i + 4), 1)
   Else:
   MsgBox "This is the fifth FG that was added to previous FG and it did not meet the requirements to be added to load 2"
   Exit Do
   End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) + PlannedProductArray(i + 2, 6) + PlannedProductArray(i + 3, 6) + PlannedProductArray(i + 4, 6) + PlannedProductArray(i + 5, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) + PlannedProductArray(i + 2, 8) + PlannedProductArray(i + 3, 8) + PlannedProductArray(i + 4, 8) + PlannedProductArray(i + 5, 8) < 44500 Then
   Debug.Print "Next Key needs to be transferred to Load 2"
   Sheet6.Range("J" & destcount + 5).Value = PlannedProductArray((i + 5), 1)
   Else:
   MsgBox "This is the sixth FG that was added to previous FG and it did not meet the requirements to be added to load 2"
   Exit Do
   End If
   
   If PlannedProductArray(i, 6) + PlannedProductArray(i + 1, 6) + PlannedProductArray(i + 2, 6) + PlannedProductArray(i + 3, 6) + PlannedProductArray(i + 4, 6) + PlannedProductArray(i + 5, 6) + PlannedProductArray(i + 6, 6) < 26 And PlannedProductArray(i, 8) + PlannedProductArray(i + 1, 8) + PlannedProductArray(i + 2, 8) + PlannedProductArray(i + 3, 8) + PlannedProductArray(i + 4, 8) + PlannedProductArray(i + 5, 8) + PlannedProductArray(i + 6, 8) < 44500 Then
   Debug.Print "Next Key needs to be transferred to Load 2"
   Sheet6.Range("J" & destcount + 6).Value = PlannedProductArray((i + 6), 1)
   Else:
   MsgBox "This is the seventh FG that was added to previous FG and it did not meet the requirements to be added to load 2"
   Exit Do
   End If
     
   Exit Do
   Loop
   
   i = i + 1
End Sub


Sub Tabulation_Main()
'________________________________________________________________________________________________________________________________
'Author:  Tracy R Kinder
'Date 12/14/19
'This routine is to save a little time by using user input boxes to enter the pertinent information for case loading
'Next routine to program is to take everything entered for a specific date and to automatically load it into the scratchpad
'by key for each FG and then create an array snapshot of that information to manipulate the case count and weight to
'maximize each load to plan.  However keep in mind Mircro items from line three need to be kept together as much as
'possible.  Those loads with Micro items need to be identified in routine HI0207 as GEM
'________________________________________________________________________________________________________________________________

' This how I unprotect the master Tabulation sheet in order to process new load planning
ActiveSheet.Unprotect "Odundbemru@123456"

'This is just to activate the cell I wish to enter new information
Sheet5.Range("B7").End(xlDown).Offset(1, 0) = Activate

'Declaration of Local Variables

Dim DateOfToday As Date
Dim CasesOnHold As Double
Dim CaseCount As Double
Dim FG_Item As Variant
Dim myrange As Range
Dim DateRange As Range
Dim Affirmation As Boolean

'Setting range for vlookup variable for another sheet in the workbook for VLookup
Set myrange = Sheet5.Range("A:A")
Set DateRange = Sheet2.Range("B7:B533")
Affirmation = True
Application.ScreenUpdating = False
DateOfToday = Sheet2.Range("G3").Value



'This is the main portion of the program to input the load information to the Tabulation worksheet


Do While (Affirmation) = True

        Do While IsDate(DateOfToday) = True
        DateOfToday = Application.InputBox(Prompt:="Please enter a valid Date (MM/DD/YYYY)", Title:="Enter Date", Type:=2)
        DateOfToday = CDate(DateOfToday)
                    If IsDate(DateOfToday) = True Then
                    MsgBox "This is a valid date format"
                    Exit Do
                    Else
                    MsgBox "This is not a valid date format"
                    End If
        Loop
        
        CasesOnHold = Application.InputBox(Prompt:="Please enter Number of Today's Planned Cases that are on Hold  ", Title:="Enter Cases on Hold for FG", Type:=1)
        
        CaseCount = Application.InputBox(Prompt:="Please enter Number of Today's Planned Cases  ", Title:="Enter Number of Cases for This FG", Type:=1)
        
        FG_Item = Application.InputBox(Prompt:="Please enter Finished Good Number of Today's Planned Cases", Title:="Enter F.G. # or Click From Database", Type:=1)
        FG_Item = Application.VLookup(FG_Item, myrange, 1, False)
        
        'Application.VLookup used here instead of Application.Worksheet.VLookup to elimate use of GOTO and Resume Statements
        
        'This Do While Loop used to ensure user puts in a Finished Good Item that is actually in the database
        
                Do While IsError(FG_Item) = True
                MsgBox "This number is not in the database"
                FG_Item = Application.InputBox(Prompt:="Please Re-enter Finished Good Number", Title:="Enter F.G. # or Click From Database", Type:=1)
                FG_Item = Application.VLookup(FG_Item, myrange, 1, False)
                
                    If IsError(FG_Item) = False Then
                    Exit Do
                    End If
                
                Loop
        
        
        
        'This portion takes the good information placed in the input boxes and updates the Tabulation Sheet
        
        Range("B7").End(xlDown).Offset(1, 0) = DateOfToday
        Range("C7").End(xlDown).Offset(1, 0) = CasesOnHold
        Range("D7").End(xlDown).Offset(1, 0) = CaseCount
        Range("F7").End(xlDown).Offset(1, 0) = FG_Item
        Range("B7").End(xlDown).Offset(1, 0) = Activate
        
        Affirmation = Application.InputBox(Prompt:="Add another? 1 = Yes, 0 = No", Title:=" 1 = Yes or 0 = No?", Type:=1)
           If Affirmation = False Then
           Exit Do
           End If
           
Loop

Range("B7").End(xlDown).Offset(1, 0) = Activate

'Always turn screen updating back on at the end of a routine


Application.ScreenUpdating = True

'It is a good practice to password protect the sheet in order to prevent tampering or inadvertent manipulation of information
'located in the cells of the spreadsheet

ActiveSheet.Protect "Odundbemru@123456"

End Sub



Sub TransferToScratch1()

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


'Setting variables to initial values

c = Sheet2.Range("G3").Value
counter = 7
destcount = 6
Set rng = Sheet2.Range("B7").CurrentRegion
lastRow = Sheet2.Range("B7").End(xlDown).Row

'Main program to transfer Key #'s to scratchpad for daily load planning of partial FG # loads
' counters will have to be reprogramed from using hard code to lower and upper bounds in the for next loop
'______________________________________________________________________________________________________________

For counter = 7 To lastRow
If c = Sheet2.Range("B" & counter) Then
destcount = destcount + 1
Sheet6.Range("A" & destcount) = Sheet2.Range("B" & counter).Offset(0, -1)
Else
End If
Next counter

'________________________________________________________
'For Full Loads only of a single FG #

counter = 7
destcount = 6

For counter = 7 To lastRow
If Sheet2.Range("B" & counter) = c And Sheet2.Range("B" & counter).Offset(0, 11) >= 1 Then
destcount = destcount + 1
Sheet6.Range("A" & destcount + 90) = Sheet2.Range("B" & counter).Offset(0, -1)
Else
End If
Next counter

'It is a good practice to password protect the sheet in order to prevent tampering or inadvertent manipulation of information
'located in the cells of the spreadsheet

Sheet6.Protect "Odundbemru@123456"
Application.ScreenUpdating = True

End Sub
" xmlns=""><RIS><RI N="Event" /></RIS><S><UCSS T="1" C="NexusTenantTokenLivePersonaCardUserActions" S="Medium" /><F T="2"><O T="GE"><L><S T="1" F="EventSamplingPolicy" /></L><R><V V="191" T="U8" /></R></O></F></S><C T="W" I="0" O="false" N="EventName"><S T="2" F="EventName" /></C><C T="W" I="1" O="true" N="EventContract"><S T="2" F="EventContract" M="Ignore" /></C><C T="U64" I="2" O="false" N="EventFlags"><S T="2" F="EventFlags" /></C><C T="D" I="3" 