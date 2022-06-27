VERSION 1.4.1 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rcTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'******************************************************************************************************************************************************************************************************************** _
Desenvolvido em Excel VBA _
por: Ricardo Camisa _
Inscreva-se no meu canal e deia uma força aí. _
link https  //www.youtube.com/c/ricardocamisa _
*************************************************************************************************************************************************************************************************************************
Public WithEvents MyForm As MSForms.UserForm
Public WithEvents TextBox As MSForms.TextBox
Public WithEvents combobox As MSForms.combobox
Public WithEvents barBottom As MSForms.Label
Public WithEvents barTitle As MSForms.Label
Public WithEvents DropArrow As MSForms.Label
Private colBaseCtrl As Collection
Private ColTextBox As Collection
Private rcTextBox  As New rcTextBox


Public Sub Add(Form As MSForms.UserForm)
    Dim ctrl As Control
    For Each ctrl In Form.Controls
        If TypeName(ctrl) = "TextBox" Or TypeName(ctrl) = "ComboBox" Then
            colBaseCtrl.Add ctrl
        Else
        End If
    Next ctrl
    Call Me.Create(Form)
End Sub

Public Sub Create(ByVal NewForm As MSForms.UserForm)
    Dim h      As Integer
    Set MyForm = NewForm
    For h = 1 To colBaseCtrl.Count
        Call Create_Sub(h, MyForm)
    Next h
End Sub

Private Sub Create_Sub(ByVal Index As Integer, Form As MSForms.UserForm)
    Dim i As Integer
    Select Case TypeName(colBaseCtrl(Index))
        Case Is = "TextBox"
            With colBaseCtrl(Index)
                .BackStyle = fmBackStyleOpaque
                .BorderStyle = fmBorderStyleSingle
                .BorderColor = vbWhite
                .Font.Name = "Poppins"
                .Font.Size = 9
            End With
            Set rcTextBox = New rcTextBox
            Set rcTextBox.TextBox = colBaseCtrl(Index)
            Set rcTextBox.MyForm = Form
            ColTextBox.Add rcTextBox
            
            Set barTitle = MyForm.Controls.Add("Forms.Label.1", colBaseCtrl(Index).Name & "barTitle")
            With barTitle
                .Caption = colBaseCtrl(Index).Tag
                .Left = colBaseCtrl(Index).Left + 8
                .BackStyle = fmBackStyleTransparent
                .ForeColor = RGB(87, 87, 200)
                .BorderStyle = fmBorderStyleNone
                .Width = colBaseCtrl(Index).Width
                .Top = colBaseCtrl(Index).Top + 2
                .Font.Name = "Poppins"
                .Enabled = False
                .Font.Size = 9
            End With
            Set rcTextBox = New rcTextBox
            Set rcTextBox.barTitle = barTitle
            Set rcTextBox.MyForm = Form
            ColTextBox.Add rcTextBox
            
            Set barBottom = MyForm.Controls.Add("Forms.Label.1", colBaseCtrl(Index).Name & "barBottom")
            With barBottom
                .Caption = ""
                .Height = 1
                .Left = colBaseCtrl(Index).Left + 8
                .BackStyle = fmBackStyleOpaque
                .BackColor = RGB(87, 87, 200)
                .BorderStyle = fmBorderStyleNone
                .Width = colBaseCtrl(Index).Width
                .Top = colBaseCtrl(Index).Top + colBaseCtrl(Index).Height
            End With
            Set rcTextBox = New rcTextBox
            Set rcTextBox.barBottom = barBottom
            Set rcTextBox.MyForm = Form
            ColTextBox.Add rcTextBox
        Case Is = "ComboBox"
            With colBaseCtrl(Index)
                .BackStyle = fmBackStyleOpaque
                .BorderStyle = fmBorderStyleSingle
                .BorderColor = vbWhite
                .Font.Name = "Poppins"
                .Font.Size = 9
                .ShowDropButtonWhen = 0
                .Style = fmStyleDropDownList
                .DropButtonStyle = fmDropButtonStyleArrow
            End With
            Set rcTextBox = New rcTextBox
            Set rcTextBox.combobox = colBaseCtrl(Index)
            Set rcTextBox.MyForm = Form
            ColTextBox.Add rcTextBox
            
            Set barTitle = MyForm.Controls.Add("Forms.Label.1", colBaseCtrl(Index).Name & "barTitle")
            With barTitle
                .Caption = colBaseCtrl(Index).Tag
                .Left = colBaseCtrl(Index).Left + 8
                .BackStyle = fmBackStyleTransparent
                .ForeColor = RGB(87, 87, 200)
                .BorderStyle = fmBorderStyleNone
                .Width = colBaseCtrl(Index).Width
                .Top = colBaseCtrl(Index).Top + 2
                .Font.Name = "Poppins"
                .Enabled = False
                .Font.Size = 9
            End With
            Set rcTextBox = New rcTextBox
            Set rcTextBox.barTitle = barTitle
            Set rcTextBox.MyForm = Form
            ColTextBox.Add rcTextBox
            
            Set barBottom = MyForm.Controls.Add("Forms.Label.1", colBaseCtrl(Index).Name & "barBottom")
            With barBottom
                .Caption = ""
                .Height = 1
                .Left = colBaseCtrl(Index).Left + 8
                .BackStyle = fmBackStyleOpaque
                .BackColor = RGB(87, 87, 200)
                .BorderStyle = fmBorderStyleNone
                .Width = colBaseCtrl(Index).Width - 8
                .Top = colBaseCtrl(Index).Top + colBaseCtrl(Index).Height
            End With
            
            Set DropArrow = MyForm.Controls.Add("Forms.Label.1", colBaseCtrl(Index).Name & "DropArrow")
            With DropArrow
                .Font.Name = "Webdings"
                .Caption = "6"
                .Height = colBaseCtrl(Index).Height
                .Width = .Height
                .Left = (colBaseCtrl(Index).Width + colBaseCtrl(Index).Left) - .Width + 8
                .BackStyle = fmBackStyleTransparent
                .ForeColor = RGB(87, 87, 200)
                .BorderStyle = fmBorderStyleNone
                .Top = colBaseCtrl(Index).Top + 3
                .Enabled = False
                .Font.Size = 10
            End With
            Set rcTextBox = New rcTextBox
            Set rcTextBox.DropArrow = DropArrow
            Set rcTextBox.MyForm = Form
            ColTextBox.Add rcTextBox
    End Select
End Sub

Private Sub Class_Initialize()
    Set colBaseCtrl = New Collection
    Set ColTextBox = New Collection
End Sub

Private Sub Class_Terminate()
    Set colBaseCtrl = Nothing
    Set ColTextBox = Nothing
    Set MyForm = Nothing
End Sub

Private Sub Combobox_Change()
    Dim Control As Control
    For Each Control In MyForm.Controls
        If TypeName(Control) = "Label" And Control.Name = combobox.Name & "barBottom" Then
            MyForm.Controls(Control.Name).BackColor = RGB(87, 87, 200)
            With MyForm.Controls(combobox.Name & "barTitle")
                .ForeColor = RGB(87, 87, 200)
                .Top = combobox.Top - 15
                .Enabled = True
            End With
            Exit For
        End If
    Next
End Sub

Private Sub ComboBox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control As Control
    For Each Control In MyForm.Controls
        If TypeName(Control) = "Label" And Control.Name = combobox.Name & "barBottom" Then
            MyForm.Controls(Control.Name).BackColor = RGB(87, 87, 200)
            With MyForm.Controls(combobox.Name & "barTitle")
                .ForeColor = RGB(87, 87, 200)
                .Top = combobox.Top - 15
                .Enabled = True
            End With
            Exit For
        End If
    Next
End Sub

Private Sub MyForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control As Control
    For Each Control In MyForm.Controls
        If TypeName(Control) = "TextBox" Or TypeName(Control) = "ComboBox" Then
            If Control.Text = "" Then
                MyForm.Controls(Control.Name & "barBottom").BackColor = RGB(87, 87, 200)
                With MyForm.Controls(Control.Name & "barTitle")
                    .ForeColor = RGB(87, 87, 200)
                    .Top = Control.Top + 2
                    .Enabled = False
                End With
            End If
        End If
    Next
End Sub

Private Sub TextBox_Change()
    Dim Control As Control
    If TextBox.value = "" Then
        For Each Control In MyForm.Controls
            If TypeName(Control) = "Label" And Control.Name = TextBox.Name & "barBottom" Then
                MyForm.Controls(TextBox.Name & "barBottom").BackColor = RGB(87, 87, 200)
                With MyForm.Controls(TextBox.Name & "barTitle")
                    .ForeColor = RGB(87, 87, 200)
                    .Top = TextBox.Top + 2
                    .Enabled = True
                End With
                Exit For
            End If
        Next
    Else
        For Each Control In MyForm.Controls
            If TypeName(Control) = "Label" And Control.Name = TextBox.Name & "barBottom" Then
                MyForm.Controls(TextBox.Name & "barBottom").BackColor = RGB(87, 87, 200)
                With MyForm.Controls(TextBox.Name & "barTitle")
                    .ForeColor = RGB(87, 87, 200)
                    .Top = TextBox.Top - 15
                    .Enabled = True
                End With
                Exit For
            End If
        Next
    End If
End Sub

Private Sub TextBox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control As Control
    For Each Control In MyForm.Controls
        If TypeName(Control) = "Label" And Control.Name = TextBox.Name & "barBottom" Then
            MyForm.Controls(TextBox.Name & "barBottom").BackColor = RGB(87, 87, 200)
            With MyForm.Controls(TextBox.Name & "barTitle")
                .ForeColor = RGB(87, 87, 200)
                .Top = TextBox.Top - 15
                .Enabled = True
            End With
            Exit For
        End If
    Next
End Sub




