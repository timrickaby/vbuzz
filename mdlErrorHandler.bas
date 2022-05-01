Attribute VB_Name = "mdlMessageHandler"

' vbGo
' VBuzz - The VB Information Platform
' Copyright © 1999 - 2001 vbGo. All Rights Reserved
' Version 1.0 Build 2

Option Explicit

Public Enum enumMessageTypes
    MessageConnection = 0
    MessageInformation = 1
    MessageQuestion = 2
    MessageAlert = 3
    MessageError = 4
End Enum

Public Enum enumReturnTypes
    ButtonOK = 0
    ButtonCancel = 1
    ButtonClose = 2
End Enum

Public Enum enumButtons
    ButtonsOKCancel = 0
    ButtonsClose = 1
    ButtonsCancel = 2
    ButtonsNone = 3
End Enum


Public Function ShowMessage(ByVal strMessageTitle As String, _
    ByVal strMessageBody As String, _
    eMessageType As enumMessageTypes, _
    eButtons As enumButtons, _
    Optional eDisplayModally As FormShowConstants = vbModeless) As enumReturnTypes

    If frmMessage.Visible = True Then frmMessage.Hide

    frmMessage.imgInternet.Visible = False
    frmMessage.imgInformation.Visible = False
    frmMessage.imgQuestion.Visible = False
    frmMessage.imgAlert.Visible = False
    frmMessage.imgError.Visible = False
    
    Select Case eMessageType
        Case MessageConnection
            frmMessage.imgInternet.Visible = True
        Case MessageInformation
            frmMessage.imgInformation.Visible = True
        Case MessageQuestion
            frmMessage.imgQuestion.Visible = True
        Case MessageAlert
            frmMessage.imgAlert.Visible = True
        Case MessageError
            frmMessage.imgError.Visible = True
    End Select

    If eButtons = ButtonsClose Then
        frmMessage.cmdCancel.Visible = False
        frmMessage.cmdOK.Visible = False
        frmMessage.cmdClose.Visible = True
    
    ElseIf eButtons = ButtonsOKCancel Then
        frmMessage.cmdCancel.Visible = True
        frmMessage.cmdOK.Visible = True
        frmMessage.cmdClose.Visible = False
    
    ElseIf eButtons = ButtonsCancel Then
        frmMessage.cmdCancel.Visible = True
        frmMessage.cmdOK.Visible = False
        frmMessage.cmdClose.Visible = False
        
    ElseIf eButtons = ButtonsNone Then
        frmMessage.cmdCancel.Visible = False
        frmMessage.cmdOK.Visible = False
        frmMessage.cmdClose.Visible = False
        frmMessage.Height = frmMessage.Height - 300
    End If
    
    frmMessage.lblMessageTitle.Caption = strMessageTitle
    frmMessage.lblMessageBody.Caption = strMessageBody
    frmMessage.Show
    
    If frmMessage.strReturnType = "BUTTONOK" Then
        ShowMessage = ButtonOK
    ElseIf frmMessage.strReturnType = "BUTTONCANCEL" Then
        ShowMessage = ButtonCancel
    ElseIf frmMessage.strReturnType = "BUTTONCLOSE" Then
        ShowMessage = ButtonClose
    End If
End Function