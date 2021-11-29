VERSION 1.0 CLASS
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
Email: rich.7.2014@gmail.com _
Telefone: +244 925341780 _
Inscreva-se no meu canal e deia uma força aí. _
link: https://www.youtube.com/c/ricardocamisa _
*************************************************************************************************************************************************************************************************************************

Public Form                 As MSForms.UserForm
Public WithEvents TextBox   As MSForms.TextBox
Public WithEvents barBottom As MSForms.Label
Public WithEvents barTitle  As MSForms.Label
Public CnvNumber            As Long
Public CnvString            As String

Public Sub AddRcTextBox(MyForm As Object)
    Set Form = MyForm
    
    For Each ctrl In MyForm.Controls
        If ctrl.Tag = "Form" Then
            nameCtrl = ctrl.Name
            Dim Ctrltxt As Control
            
            For Each Ctrltxt In ctrl.Controls
                If TypeName(Ctrltxt) = "TextBox" Then
                    Set rcTextBox = New rcTextBox
                    Set rcTextBox.TextBox = Ctrltxt
                    Set rcTextBox.Form = Form
                    ColTextBox.Add rcTextBox
                    
                    With Ctrltxt
                        .BackStyle = fmBackStyleOpaque
                        .BorderStyle = fmBorderStyleSingle
                        .BorderColor = vbWhite
                        .Font.Name = "Poppins"
                        .Font.Size = 9
                    End With
                    
                    Set barTitle = Form.Controls(nameCtrl).Add("Forms.Label.1", Ctrltxt.Name & "barTitle")
                    With barTitle
                        .Caption = Ctrltxt.Tag
                        .Left = Ctrltxt.Left + 8
                        .BackStyle = fmBackStyleTransparent
                        .ForeColor = RGB(102, 102, 102)
                        .BorderStyle = fmBorderStyleNone
                        .Width = Ctrltxt.Width
                        .Top = Ctrltxt.Top + 2
                        .Font.Name = "Poppins"
                        .Font.Size = 9
                    End With
                    Set rcTextBox = New rcTextBox
                    Set rcTextBox.barTitle = barTitle
                    ColTextBox.Add rcTextBox
                    
                    Set barBottom = Form.Controls(nameCtrl).Add("Forms.Label.1", Ctrltxt.Name & "barBottom")
                    With barBottom
                        .Caption = ""
                        .Height = 1
                        .Left = Ctrltxt.Left + 8
                        .BackStyle = fmBackStyleOpaque
                        .BackColor = RGB(102, 102, 102)
                        .BorderStyle = fmBorderStyleNone
                        .Width = Ctrltxt.Width
                        .Top = Ctrltxt.Top + Ctrltxt.Height
                    End With

                End If
            Next
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
        End With
    End If

End Sub

