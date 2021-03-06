VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Tabulation"
   ClientHeight    =   4836
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public IsCancelled As Boolean

Private Sub optButton_Click()

End Sub

Private Sub UserForm1_Initialize()

    IsCancelled = True
End Sub

Private Sub btnCancel_Click()
    Me.Hide
End Sub

Private Sub btnOk_Click()
    If IsInputOk Then
        IsCancelled = False
        Me.Hide
    End If
    
    
End Sub

Public Sub SetValues(UserForm1)

Dim UserInputCaseCount As Double
Dim UserInputFinishedGoodItem As Double

    With UserForm1
        SetValue Me.txtLabelCaseCount, .LabelCaseCount
        SetValue Me.txtLabelFinishedGoodItem, .LabelFinishedGoodItem
        SetValue Me.txtUserInputCaseCount, .UserInputCaseCount
        SetValue Me.txtUserInputFinishedGoodItem, .UserInputFinishedGoodItem
        SetValue Me.optButton, .Button
    End With
End Sub

Public Sub GetValues(UserForm1)
    With udtTabulation
        .LabelCaseCount = GetValue(Me.txtLabelCaseCount, TypeName(.LabelCaseCount))
        .LabelFinishedGoodItem = GetValue(Me.txtLabelFinishedGoodItem, TypeName(.LabelFinishedGoodItem))
        .UserInputCaseCount = GetValue(Me.txtUserInputCaseCount, TypeName(.UserInputCaseCount))
        .UserInputFinishedGoodItem = GetValue(Me.txtUserInputFinishedGoodItem, TypeName(.UserInputFinishedGoodItem))
        .Button = GetValue(Me.optButton, TypeName(.Button))
    End With
End Sub

Private Function IsInputOk() As Boolean
Dim ctl As MSForms.Control
Dim strMessage As String
    IsInputOk = False
    For Each ctl In Me.Controls
        If IsInputControl(ctl) Then
            If IsRequired(ctl) Then
                If Not HasValue(ctl) Then
                    strMessage = ControlName(ctl) & " must have value"
                End If
            End If
            If Not IsCorrectType(ctl) Then
                strMessage = ControlName(ctl) & " is not correct"
            End If
        End If
        If Len(strMessage) > 0 Then
            ctl.SetFocus
            GoTo HandleMessage
        End If
    Next
    IsInputOk = True
HandleExit:
    Exit Function
HandleMessage:
    MsgBox strMessage
    GoTo HandleExit
End Function

Public Sub FillList(ControlName As String, Values As Variant)
    With Me.Controls(ControlName)
        Dim iArrayForNext As Long
        .Clear
        For iArrayForNext = LBound(Values) To UBound(Values)
            .AddItem Values(iArrayForNext)
        Next
    End With
End Sub

Private Function IsCorrectType(ctl As MSForms.Control) As Boolean
Dim strControlDataType As String, strMessage As String
Dim dummy As Variant
    strControlDataType = ControlDataType(ctl)
On Error GoTo HandleError
    Select Case strControlDataType
    Case "Boolean"
        dummy = CBool(GetValue(ctl, strControlDataType))
    Case "Byte"
        dummy = CByte(GetValue(ctl, strControlDataType))
    Case "Currency"
        dummy = CCur(GetValue(ctl, strControlDataType))
    Case "Date"
        dummy = CDate(GetValue(ctl, strControlDataType))
    Case "Double"
        dummy = CDbl(GetValue(ctl, strControlDataType))
    Case "Decimal"
        dummy = CDec(GetValue(ctl, strControlDataType))
    Case "Integer"
        dummy = CInt(GetValue(ctl, strControlDataType))
    Case "Long"
        dummy = CLng(GetValue(ctl, strControlDataType))
    Case "Single"
        dummy = CSng(GetValue(ctl, strControlDataType))
    Case "String"
        dummy = CStr(GetValue(ctl, strControlDataType))
    Case "Variant"
        dummy = CVar(GetValue(ctl, strControlDataType))
    End Select
    IsCorrectType = True
HandleExit:
    Exit Function
HandleError:
    IsCorrectType = False
    Resume HandleExit
End Function

Private Function ControlDataType(ctl As MSForms.Control) As String
    Select Case ctl.Name
    Case "txtLabelCaseCount": ControlDataType = "String"
    Case "txtLabelFinishedGoodItem": ControlDataType = "String"
    Case "txtUserInputCaseCount": ControlDataType = "Double"
    Case "txtUserInputFinishedGoodItem": ControlDataType = "Double"
    Case "optButton": ControlDataType = "String"
    End Select
End Function

Private Function ControlName(ctl As MSForms.Control) As String
On Error GoTo HandleError
    If Not ctl Is Nothing Then
        ControlName = ctl.Name
        Select Case TypeName(ctl)
        Case "TextBox", "ListBox", "ComboBox"
            If ctl.TabIndex > 0 Then
                Dim c As MSForms.Control
                For Each c In Me.Controls
                    If c.TabIndex = ctl.TabIndex - 1 Then
                        If TypeOf c Is MSForms.Label Then
                            ControlName = c.Caption
                        End If
                    End If
                Next
            End If
        Case Else
            ControlName = ctl.Caption
        End Select
    End If
HandleExit:
    Exit Function
HandleError:
    Resume HandleExit
End Function

Private Function IsRequired(ctl As MSForms.Control) As Boolean
    Select Case ctl.Name
    Case "txtLabelCaseCount", "txtLabelFinishedGoodItem", "txtUserInputCaseCount", "txtUserInputFinishedGoodItem", "optButton"
        IsRequired = True
    Case Else
        IsRequired = False
    End Select
End Function

Private Function IsInputControl(ctl As MSForms.Control) As Boolean
    Select Case TypeName(ctl)
    Case "TextBox", "ComboBox", "ListBox", "CheckBox", "OptionButton", "ToggleButton"
        IsInputControl = True
    Case Else
        IsInputControl = False
    End Select
End Function

Private Function HasValue(ctl As MSForms.Control) As Boolean
    Dim var As Variant
    var = GetValue(ctl, "Variant")
    If IsNull(var) Then
        HasValue = False
    ElseIf Len(var) = 0 Then
        HasValue = False
    Else
        HasValue = True
    End If
End Function

Private Function GetValue(ctl As MSForms.Control, strTypeName As String) As Variant
On Error GoTo HandleError
    Dim Value As Variant
    Value = ctl.Value
    If IsNull(Value) And strTypeName <> "Variant" Then
        Select Case strTypeName
        Case "String"
            Value = ""
        Case Else
            Value = 0
        End Select
    End If
HandleExit:
    GetValue = Value
    Exit Function
HandleError:
    Resume HandleExit
End Function

Private Sub SetValue(ctl As MSForms.Control, Value As Variant)
On Error GoTo HandleError
    ctl.Value = Value
HandleExit:
    Exit Sub
HandleError:
    Resume HandleExit UserForm1.Show
End Sub


roductArray(i + 2, 8) + PlannedProductArray(i + 3, 8) + PlannedProductArray(i + 4, 8) < 44500 Then
   Debug.Print "Next Key needs to be transferred to Load 1"
   She                                                                                                                                                                                                                                                                                                                                                         