VERSION 1.3.1 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rcTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************************************************************************************************************************************************** _
Desenvolvido em Excel VBA _
por: Ricardo Camisa _
Email rich0.7.2014@gmail.com _
Telefone:  +244 925341780 _
Inscreva-se no meu canal e deia uma força aí. _
link https  //www.youtube.com/c/ricardocamisa _
*************************************************************************************************************************************************************************************************************************
Public WithEvents Form As MSForms.UserForm
Public WithEvents Frame As MSForms.Frame
Public WithEvents TextBox As MSForms.TextBox
Public WithEvents ComboBox As MSForms.ComboBox
Public WithEvents barBottom As MSForms.Label
Public WithEvents barTitle As MSForms.Label
Public WithEvents DropArrow As MSForms.Label
Public CnvNumber As Long
Public CnvString As String


Public Sub AddRcTextBox(MyForm As Object)
    Set Form = MyForm
    
    For Each ctrl In MyForm.Controls
        If ctrl.Tag = "Form" Then
            Set Frame = ctrl
            nameCtrl = ctrl.Name
            Dim Control As Control


            For Each Control In ctrl.Controls

                Select Case TypeName(Control)

                    Case Is = "TextBox"
                        Set rcTextBox = New rcTextBox
                        Set rcTextBox.TextBox = Control
                        Set rcTextBox.Form = Form
                        ColTextBox.Add rcTextBox

                        With Control
                            .BackStyle = fmBackStyleOpaque
                            .BorderStyle = fmBorderStyleSingle
                            .BorderColor = vbWhite
                            .Font.Name = "Poppins"
                            .Font.Size = 9
                        End With
                        
                        Set barTitle = Form.Controls(nameCtrl).Add("Forms.Label.1", Control.Name & "barTitle")
                        With barTitle
                            .Caption = Control.Tag
                            .Left = Control.Left + 8
                            .BackStyle = fmBackStyleTransparent
                            .ForeColor = RGB(102, 102, 102)
                            .BorderStyle = fmBorderStyleNone
                            .Width = Control.Width
                            .Top = Control.Top + 2
                            .Font.Name = "Poppins"
                            .Enabled = False
                            .Font.Size = 9
                        End With
                        Set rcTextBox = New rcTextBox
                        Set rcTextBox.barTitle = barTitle
                        ColTextBox.Add rcTextBox
                        
                        Set barBottom = Form.Controls(nameCtrl).Add("Forms.Label.1", Control.Name & "barBottom")
                        With barBottom
                            .Caption = ""
                            .Height = 1
                            .Left = Control.Left + 8
                            .BackStyle = fmBackStyleOpaque
                            .BackColor = RGB(102, 102, 102)
                            .BorderStyle = fmBorderStyleNone
                            .Width = Control.Width
                            .Top = Control.Top + Control.Height
                        End With

                    Case Is = "ComboBox"
                    
                        Set rcTextBox = New rcTextBox
                        Set rcTextBox.ComboBox = Control
                        Set rcTextBox.Form = Form
                        ColTextBox.Add rcTextBox

                        With Control
                            .BackStyle = fmBackStyleOpaque
                            .BorderStyle = fmBorderStyleSingle
                            .BorderColor = vbWhite
                            .Font.Name = "Poppins"
                            .Font.Size = 9
                            .ShowDropButtonWhen = 0
                            .Style = fmStyleDropDownList
                            .DropButtonStyle = fmDropButtonStyleArrow
                        End With
                        
                        Set barTitle = Form.Controls(nameCtrl).Add("Forms.Label.1", Control.Name & "barTitle")
                        With barTitle
                            .Caption = Control.Tag
                            .Left = Control.Left + 8
                            .BackStyle = fmBackStyleTransparent
                            .ForeColor = RGB(102, 102, 102)
                            .BorderStyle = fmBorderStyleNone
                            .Width = Control.Width
                            .Top = Control.Top + 2
                            .Font.Name = "Poppins"
                            .Enabled = False
                            .Font.Size = 9
                        End With
                        Set rcTextBox = New rcTextBox
                        Set rcTextBox.barTitle = barTitle
                        ColTextBox.Add rcTextBox
                        
                        Set barBottom = Form.Controls(nameCtrl).Add("Forms.Label.1", Control.Name & "barBottom")
                        With barBottom
                            .Caption = ""
                            .Height = 1
                            .Left = Control.Left + 8
                            .BackStyle = fmBackStyleOpaque
                            .BackColor = RGB(102, 102, 102)
                            .BorderStyle = fmBorderStyleNone
                            .Width = Control.Width - 8
                            .Top = Control.Top + Control.Height
                        End With
                        
                        Set DropArrow = Form.Controls(nameCtrl).Add("Forms.Label.1", Control.Name & "DropArrow")
                        With DropArrow
                            .Font.Name = "Webdings"
                            .Caption = "6"
                            .Height = Control.Height
                            .Width = .Height
                            .Left = (Control.Width + Control.Left) - .Width + 6
                            .BackStyle = fmBackStyleTransparent
                            .ForeColor = RGB(102, 102, 102)
                            .BorderStyle = fmBorderStyleNone
                            .Top = Control.Top + 1.5
                            .Enabled = False
                            .Font.Size = 10
                        End With
                        Set rcTextBox = New rcTextBox
                        Set rcTextBox.DropArrow = DropArrow
                        ColTextBox.Add rcTextBox


                End Select

            Next
        End If
    Next
End Sub

Private Sub ComboBox_Change()
    Dim lbl As Control

    If ComboBox.Value = "" Then
        For Each lbl In Form.Controls(nameCtrl).Controls
            If TypeName(lbl) = "Label" And lbl.Name = ComboBox.Name & "barBottom" Then
                namelbl = lbl.Name
            End If
        Next

        Form.Controls(ComboBox.Name & "barBottom").BackColor = RGB(102, 102, 102)
        With Form.Controls(ComboBox.Name & "barTitle")
            .ForeColor = RGB(102, 102, 102)
            .Top = ComboBox.Top + 2
            .Enabled = False
        End With
    Else
        For Each lbl In Form.Controls(nameCtrl).Controls
            If TypeName(lbl) = "Label" And lbl.Name = ComboBox.Name & "barBottom" Then
                namelbl = lbl.Name
            End If
        Next

        Form.Controls(ComboBox.Name & "barBottom").BackColor = RGB(4, 0, 247)
        With Form.Controls(ComboBox.Name & "barTitle")
            .ForeColor = RGB(4, 0, 247)
            .Top = ComboBox.Top - 15
            .Enabled = True
        End With
    End If
End Sub

Private Sub ComboBox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim lbl As Control

    For Each lbl In Form.Controls(nameCtrl).Controls
        If TypeName(lbl) = "Label" And lbl.Name = ComboBox.Name & "barBottom" Then
            namelbl = lbl.Name
        End If
    Next

    Form.Controls(ComboBox.Name & "barBottom").BackColor = RGB(4, 0, 247)
    With Form.Controls(ComboBox.Name & "barTitle")
        .ForeColor = RGB(4, 0, 247)
        .Top = ComboBox.Top - 15
        .Enabled = True
    End With
End Sub

Private Sub Frame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Control As Control

    For Each Control In Frame.Controls
        If TypeName(Control) = "TextBox" Or TypeName(Control) = "ComboBox" Then
            If Control.Text = "" Then
                Frame.Controls(Control.Name & "barBottom").BackColor = RGB(102, 102, 102)
                With Frame.Controls(Control.Name & "barTitle")
                    .ForeColor = RGB(102, 102, 102)
                    .Top = Control.Top + 2
                    .Enabled = False
                End With
            End If
        End If
    Next

End Sub

Private Sub TextBox_Change()
    Dim lbl As Control

    If TextBox.Value = "" Then
        For Each lbl In Form.Controls(nameCtrl).Controls
            If TypeName(lbl) = "Label" And lbl.Name = TextBox.Name & "barBottom" Then
                namelbl = lbl.Name
            End If
        Next

        Form.Controls(TextBox.Name & "barBottom").BackColor = RGB(102, 102, 102)
        With Form.Controls(TextBox.Name & "barTitle")
            .ForeColor = RGB(102, 102, 102)
            .Top = TextBox.Top + 2
            .Enabled = False
        End With
    Else
        For Each lbl In Form.Controls(nameCtrl).Controls
            If TypeName(lbl) = "Label" And lbl.Name = TextBox.Name & "barBottom" Then
                namelbl = lbl.Name
            End If
        Next

        Form.Controls(TextBox.Name & "barBottom").BackColor = RGB(4, 0, 247)
        With Form.Controls(TextBox.Name & "barTitle")
            .ForeColor = RGB(4, 0, 247)
            .Top = TextBox.Top - 15
            .Enabled = True
        End With
    End If

End Sub

Private Sub TextBox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim lbl As Control

    For Each lbl In Form.Controls(nameCtrl).Controls
        If TypeName(lbl) = "Label" And lbl.Name = TextBox.Name & "barBottom" Then
            namelbl = lbl.Name
        End If
    Next

    Form.Controls(TextBox.Name & "barBottom").BackColor = RGB(4, 0, 247)
    With Form.Controls(TextBox.Name & "barTitle")
        .ForeColor = RGB(4, 0, 247)
        .Top = TextBox.Top - 15
        .Enabled = True
    End With

End Sub


