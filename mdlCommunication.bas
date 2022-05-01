Attribute VB_Name = "mdlCommunication"

' vbGo
' VBuzz - The VB Information Platform
' Copyright © 1999 - 2001 vbGo. All Rights Reserved
' Version 1.0 Build 2

Option Explicit
    
' **************************************************************
'
' DOWNLOAD FILE FROM THE INTERNET
' This method allows us to retrieve the specified file from the vbGo web site so that it can be displayed in Info Desk
'
' **************************************************************
'
Public Sub GetFromWeb(ByVal strWebAddress As String, ByVal strSaveLocation As String)

    Dim lngFunctionValue As Long
    Dim intPauseAmmount As Variant
    Dim intStartTime As Variant

    typeEngine.blnCurrentlyBusy = True

        Call mdlMessageHandler.ShowMessage("Requesting Authorisation...", "Please wait while vbGo authorises the request...", MessageConnection, ButtonsCancel)
        Call Pause(1)

        Call mdlMessageHandler.ShowMessage("Reading Information...", "Please wait while VBuzz reads the required information...", MessageConnection, ButtonsCancel)
        Call Pause(1)

        frmMessage.Hide
        frmDownload.Show
        Call Pause(5)

        lngFunctionValue = URLDownloadToFile(0, strWebAddress, strSaveLocation, 0, 0)
        frmDownload.Hide

        If lngFunctionValue = 0 Then
            Call mdlMessageHandler.ShowMessage("Storing Information...", "Please wait while VBuzz stores the downloaded information...", MessageConnection, ButtonsCancel)
            Call Pause(3)
            Exit Sub

        Else
            Call mdlMessageHandler.ShowMessage("Connection Error", _
            "VBuzz could not establish an Internet connection. " & _
            "The required content could not be downloaded. Please " & _
            "try again at a later time.", MessageError, ButtonsClose)
            Exit Sub
        End If

    If mdlMessageHandler.ShowMessage("Retrieving Information...", _
    "Please wait while VBuzz connections to the vbGo web site and " & _
    "retrieves the required information for VBuzz...", _
    MessageConnection, ButtonsCancel) = 1 Then

        Call mdlMessageHandler.ShowMessage("Information Retrieval Stopped", _
        "You have chosen to cancel the information retrieval process. " & _
        "VBuzz requires this information before it can display it. " & _
        "To begin the process again, click on the Connect && Retrieve " & _
        "button on the VBuzz platform.", MessageError, ButtonsClose): Exit Sub
    End If
End Sub



