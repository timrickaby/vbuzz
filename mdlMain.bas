Attribute VB_Name = "mdlMain"

' vbGo
' VBuzz - The VB Information Platform
' Copyright © 1999 - 2001 vbGo. All Rights Reserved
' Main Interface Module

Option Explicit


' **************************************************************
'
' APPLICATION ENRTY
' This is where the application begins to execute so we will handle all command line parameters from here.
'
' **************************************************************

Private Sub Main()

    ' If there is already a previous instance of the application running, then we will not load again.
    If App.PrevInstance = True Then Exit Sub

    ' Check the current command line to see if any commands were passed to the viewer. This allows
    ' the viewer to be linked to an used from a web page. When we recieve a command line it will contain the
    ' entier protocol address.
    Select Case UCase(Command$)
    
        Case Chr(34) & "VBGOVBUZZ://ABOUT/" & Chr(34)
            frmAbout.Show vbModal
            Exit Sub
        
        Case Chr(34) & "VBGOVBUZZ://REVISE/" & Chr(34)
            frmSplash.Show vbModal
            
            Call mdlMessageHandler.ShowMessage( _
            "VBuzz Viewer Update", "There is currently no newer version of the VBuzz Viewer availible for download.", MessageInformation, ButtonsClose)
            Exit Sub
            
        Case Chr(34) & "VBGOVBUZZ://UPDATE/" & Chr(34)
            frmSplash.Show vbModal
            Call frmMain.cmdConnect_Click
            
        Case ""
            frmSplash.Show vbModal
            
        Case Else
            frmSplash.Show vbModal
            
            ' Tell the user that the parameter(s) which they passed to the viewer were no supported
            Call mdlMessageHandler.ShowMessage("Passed Parameter Not Supported", _
            "The parameter which you passed to the VBuzz viewer is not supported. This is because " & _
            "the parameter is only availible in a viewer which is newer than your current viewer version. Consider " & _
            "upgrading your viwer to the latest version.", MessageError, ButtonsClose)
            Exit Sub
    End Select
    
    frmMain.Show
End Sub



' **************************************************************
'
' PAUSE APPLICATION
' Used to pause the application for the specified amount of time, without tying up system resources
'
' **************************************************************

Public Sub Pause(intPauseTime As Integer)
    Dim intStartTime As Double
    Dim intPauseAmmount As Double
    
    intStartTime = Timer
    intPauseAmmount = intPauseTime + (1 / Rnd(Time))
    
    Do While Timer < intStartTime + intPauseAmmount
        DoEvents
    Loop
End Sub
