Attribute VB_Name = "mdlParser"

' vbGo
' VBuzz - The VB Information Platform
' Copyright © 1999 - 2001 vbGo. All Rights Reserved
' Version 1.0 Build 2

Option Explicit


Public Enum enumArgumentType
    paramValid = 0 ' Valid
    paramString = 1 ' String
    paramBlank = 2 ' Blank
End Enum


Public intArgumentCount As Integer
Public typeBaseParams(0 To 100) As tBaseParams
Public typeTempBaseParams(0 To 100) As tBaseParams

Public Type tBaseParams
    strValue As String ' Argument value
    eType As enumArgumentType ' Argument type
End Type


Public typeEngine As tEngine
Private Type tEngine
    intLineNumber As Integer ' Current line number
    intMasterFileNumber As Integer ' The main file number
    strCommand As String ' The current command word
    strCodeLine As String ' The current line of code
    strActiveFile As String ' Name of the currently active file
    blnCurrentlyBusy As Boolean ' True if currently busy
    blnStopExecuting As Boolean ' True to stop processing
    blnEndOfFile As Boolean ' Are we at the end of the file?
End Type



' **************************************************************
'
' CODE PARSER
'
' **************************************************************

Public Sub ParseCodeLine(ByVal strQuery As String, ByVal strParseChar As String)

    Dim intParserIndex As Integer: intParserIndex = 1
    Dim intArgument As Integer: intArgument = LBound(typeBaseParams)
    Dim strCurrentChar As String: strCurrentChar = ""
    Dim blnCharInCommand As Boolean: blnCharInCommand = True
    Dim blnCharInQuote As Boolean: blnCharInQuote = False
    Dim strLastChar As String: strLastChar = ""
    
    ' Exit procedure if the query is empty
    If Len(strQuery) = 0 Then Exit Sub
    
    Erase typeBaseParams

    ' Parser loop
    Do While intParserIndex <= Len(strQuery)

        strCurrentChar = Mid(strQuery, intParserIndex, 1) ' Remove a character from the curent string
        
        ' REMOVE COMMENTS
        ' If we have a comment character then we can exit the parser loop
        If (strCurrentChar = "\") And (Mid(strQuery, (intParserIndex + 1), 1) = "\") And (blnCharInQuote = False) Then
            Exit Sub
            
        ' SEARCH & REMOVE COMMAS
        ' If we have found a comma then we will need to replace it the parse character so firstly because
        ' the commas are not required they are there for readability but also because a comma which is not
        ' surrounded by a parse char is actually added to the arguments to which it is connected.
        ElseIf (strCurrentChar = ",") And (blnCharInQuote = False) Then
    
            strLastChar = strLastChar ' Keep the last character at the same value
            strCurrentChar = strParseChar ' Replace the comma with the parser character
            
            If strLastChar = strParseChar Then
                strLastChar = strCurrentChar
            End If
        End If
        
        ' ADD CHARACTER TO ARGUMENT
        If strCurrentChar <> strParseChar Then
        
            If strCurrentChar = Chr(34) Then
                ' If we are already in a string then register this quotation mark as the end of the string, else make it as the start of the string.
                blnCharInQuote = Not blnCharInQuote
                
            Else
                ' VALIDATE CHARACTER AND ADD
                ' Make sure that the current character and last character is not a parse character.
                ' If the code line contains more than one parse character in a row, only the
                ' first one is processed and the rest are ignored.
                If (strCurrentChar <> strParseChar) And (strLastChar <> strParseChar) Then
                    
                    If blnCharInCommand = False Then 
                        blnCharInCommand = True

                    Else

                        ' Add the current character to the rest of thecurrent argument.
                        typeBaseParams(intArgument).strValue = typeBaseParams(intArgument).strValue & strCurrentChar
                        
                        ' Now that we have assigned the character to the argument we need to assing the character which we have just
                        ' used to the last character variable
                        strLastChar = strCurrentChar
                    End If
                End If
            End If
        
            ' Add character to existing string
            If blnCharInQuote = True Then
                typeBaseParams(intArgument).strValue = typeBaseParams(intArgument).strValue & strCurrentChar
                typeBaseParams(intArgument).eType = paramString ' Set the type of argument which we are in
                
            Else
                ' Move to next argument
                If blnCharInCommand Then
                    
                    intArgument = intArgument + 1
                    intArgumentCount = intArgument
                    blnCharInCommand = False
                    typeBaseParams(intArgument).strValue = ""
                    typeBaseParams(intArgument).eType = paramValid ' Set the type of argument which we are in
                End If
            End If

        ' Increase the current argument number and then continue
        intParserIndex = intParserIndex + 1
    Loop
    
    ' Store the current command word as the first argument
    typeEngine.strCommand = typeBaseParams(0).strValue
End Sub



' **************************************************************
'
' READ A LINE FROM THE SPECIFIED MACRO
' 
' Global method of reading a line of code from a macro. 
' The function also process "split lines" into one simple line
'
' **************************************************************

Public Sub ReadCodeLine(ByVal blnIncreaseLineNumber As Boolean, ByVal blnParseLine As Boolean)

    Dim strTempLine As String
    
    typeEngine.blnEndOfFile = False
    
    Do
        ' Check on evey pass to see if we at the end of the current macro file, then stop if we are otherwise we
        ' will overrun the end of the file and cause VB to raise errors
        If EOF(typeEngine.intMasterFileNumber) = True Then
            ' Store that we are now at the end of the file
            typeEngine.blnEndOfFile = False: Exit Sub
        End If
        
        ' First of all check to see if we need to increase the current line count / number, some calls to this function
        ' are made by background procedures and therefore do not require that a change is made to the line number
        If blnIncreaseLineNumber = True Then
            ' Increase the current line number by one
            typeEngine.intLineNumber = typeEngine.intLineNumber + 1
        End If
        
        ' Input by line to the global line storage variable
        Line Input #typeEngine.intMasterFileNumber, typeEngine.strCodeLine
        
    Loop Until (Left(Trim(typeEngine.strCodeLine), 2) <> "\\") And _
    (Trim(typeEngine.strCodeLine) <> "")
    
    ' Trim any other unwanted spaces from the string
    typeEngine.strCodeLine = Trim(typeEngine.strCodeLine)
    
    ' We allow the macro editor to write a single long line as a multi line macro and therefore we will now have to see
    ' if we are on a multiline string
    strTempLine = typeEngine.strCodeLine
    
    Do Until Right(UCase(Trim(strTempLine)), 2) <> " _"
        ' If the end two characters of the code line are " _" then
        ' we are beign asked to add the next line to this line.
        ' Remove the end two chars ( _) from this line
        typeEngine.strCodeLine = Left(typeEngine.strCodeLine, Len(typeEngine.strCodeLine) - 2)
        
        ' Add the next line and increase the current line number
        Line Input #typeEngine.intMasterFileNumber, strTempLine
                            
        ' Store the current line with the new line and loop back
        ' to the top to see if we need to add the next line. Also add
        ' a space to the end of the currently stored line just to be
        ' sure that there is a parse character in the line. The parser
        ' can always remove this if there are too many parse chars
        typeEngine.strCodeLine = typeEngine.strCodeLine & Chr(32) & Trim(strTempLine)
        
        If blnIncreaseLineNumber = True Then
            typeEngine.intLineNumber = typeEngine.intLineNumber + 1
        End If
    Loop
        
    If blnParseLine = True Then
        Call ParseCodeLine(Trim(typeEngine.strCodeLine), Chr(32))
    End If
End Sub
