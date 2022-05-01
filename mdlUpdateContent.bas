Attribute VB_Name = "mdlUpdateContent"

' vbGo
' VBuzz - The VB Information Platform
' Copyright © 1999 - 2001 vbGo. All Rights Reserved
' Version 1.0 Build 2

Option Explicit

Private Declare Function GetPrivateProfileString Lib "Kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, _
    ByVal lpFileName As String) As Long
    
Public Declare Function URLDownloadToFile Lib "URLMon.dll" Alias "URLDownloadToFileA" (ByVal lngCaller As Long, _
    ByVal strURL As String, ByVal strFileName As String, _
    ByVal lngReserved As Long, ByVal lngfnCB As Long) As Long
    


' **************************************************************
'
' READ SETTING FROM FILE
'
' **************************************************************

Private Function ReadFromFile(ByVal strFileName As String, ByVal strFileHeaderName As String, _
    ByVal strFileKeyName As String, ByVal strDefaultSetting As String) As String

    Dim strReturnedString As String
    Dim lngNC As Long
    
    ' Create a string buffer 255 chars in size 
    strReturnedString = String(255, 0)
    
    ' Retrieve the setting from the contents file
    lngNC = GetPrivateProfileString(strFileHeaderName, strFileKeyName, _
    strDefaultSetting, strReturnedString, 255, strFileName)
    
    If lngNC <> 0 Then
    
        strReturnedString = Left(strReturnedString, lngNC)
    End If
    
    ReadFromFile = strReturnedString
End Function



' **************************************************************
'
' CHECK FOR CONTENT UPDATE
' Check the vbGo web site to see if there is any content availible for download.
'
' **************************************************************

Public Sub CheckForUpdate()

    Dim lngFunctionValue As Long
    Dim strOldDateStamp As Date
    Dim strNewDateStamp As Date
    Dim strContentsFile As String
    
    strContentsFile = App.Path & "\content\vbuzztoc.vbuzz"

    ' Tell the user that we are connecting to the vbGo site
    With frmDownload
        .lblMessageTitle.Caption = "Checking For Content Updates..."
        .lblMessageBody.Caption = "Please wait while VBuzz connects to " & "vbGo and checks for any availible content updates."
        .lblStatusLabel.Caption = "Connecting to vbGo..."
        .picStatusFore.Width = (.picStatusBack.ScaleWidth / 2)
        .Show
    End With
    
    Call Pause(3)
    frmDownload.Hide
    
    ' Store the current date settings for the currently availible (local) viewer contents
    strOldDateStamp = ReadFromFile(strContentsFile, "CONTENTSTATUS", "DATE", Date)
    
    ' Download the index file which will contain all of the
    ' required information and store the functions returned value
    lngFunctionValue = URLDownloadToFile(0, "http://www.vbgo.co.uk/code/vbuzz/content/toc.vbuzz", strContentsFile, 0, 0)
    
    If lngFunctionValue <> 0 Then
        
        ' The API function returned an error code to use this lets
        ' us know that the function did not complete.
        Call mdlMessageHandler.ShowMessage("Connection Error", _
        "VBuzz could not determine whether there is currently " & _
        "updated contents ready for download. Please try again " & _
        "at a later time.", MessageError, ButtonsClose, vbModal)
        Exit Sub
    End If
        
    ' Check for update
    With frmDownload
        .lblMessageTitle.Caption = "Processing Stored Information"
        .lblMessageBody.Caption = "Please wait while VBuzz processes " & _
        "the content stored at vbGo and determines whether there " & _
        "is an updated version availible."
        .lblStatusLabel.Caption = "Checking For Content Update..."
        .picStatusFore.Width = (.picStatusBack.ScaleWidth)
        .Show
    End With
    
    Call Pause(3)
    frmDownload.Hide
    
    ' Store the new date settings for the currently availible (local) viewer contents
    strNewDateStamp = ReadFromFile(strContentsFile, "CONTENTSTATUS", "DATE", Date)
    
    ' Compare the dates to see if an update is availible.
    ' Function will return the difference between the days section of the two dates, 
    ' if the return value is minus or 0 then there is no update availible, higher than 1 and there is an update.
    If DateDiff("d", strOldDateStamp, strNewDateStamp) > 0 Then
    
        Dim intNumberOfFiles As Integer
        Dim intIndex As Integer
        Dim strFileNames() As String
    
        ' Store the number of files required for this content
        intNumberOfFiles = ReadFromFile(strContentsFile, "CONTENT", "COUNT", "")
        
        ' Resize the file name array to the number of files which we have
        ReDim strFileNames(intNumberOfFiles) As String
        
        ' Loop through all sections of the contents file and get the names of the files required
        For intIndex = 1 To intNumberOfFiles
            
            ' Store names of the files required in the file names array
            strFileNames(intIndex) = ReadFromFile(strContentsFile, "CONTENT", "FILE" & intIndex, "")
        Next intIndex
    
    Else
        Call mdlMessageHandler.ShowMessage("No Content Update Availible", _
        "There is currently no content update availible for VBuzz. The content on your system is up to date.", _
        MessageInformation, ButtonsClose, vbModal)
    End If
End Sub
