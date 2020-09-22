Attribute VB_Name = "ModServer"
Dim SFileNum As Byte, CFileNum As Byte, BytesSent, BytesReceived

Public Function ReceiveData(WS As Winsock)
Dim Dat As String, Filename As String, FileSize As String

' Get the data from winsock
WS.GetData Dat$

' Check to see what the data is
If Left(Dat$, 4) = "FILE" Then ' A file is being sent
    Dat$ = Right(Dat$, Len(Dat$) - 5)
    
    ' Get the name and the size of the file
    a = InStr(1, Dat$, ":")
    Filename = Mid(Dat$, 1, a - 1)
    FileSize = Mid(Dat$, a + 1)
    
    ' Ask the user if he would like to save the file
    retval = MsgBox(WS.RemoteHostIP & " is attempting to send a file." & _
    vbNewLine & "File Name: " & Filename & vbNewLine & _
    "File Size: " & FileSize & vbNewLine & _
    "Would you like to receive this file?", vbYesNo, "File Transfer")
    
    ' See what the user clicked, and respond appropriately
    If retval = vbYes Then
        ' Show a save box...I will use the one of form1
        filesend.CD.Filename = Filename
        filesend.CD.ShowSave
        
        ' Open the file
        SFileNum = FreeFile
        Open filesend.CD.Filename For Binary As #SFileNum
        
        ' Tell the other computer to start sending it
        WS.SendData "ACCEPTED"
    Else
        ' tell the other computer that u didn't want it
        WS.SendData "DENIED"
    End If

ElseIf Dat$ = "DENIED" Then ' The user denied the file
    MsgBox "The other user denied the file!", vbInformation, "File Transfer"
    Exit Function
ElseIf Dat$ = "ACCEPTED" Then ' The user accepted the file
    ' Check to see if the file can be sent in one go
    If LOF(CFileNum) <= 4500 Then
        ' get that part of the file
        Dat$ = "LAST:" & Input(LOF(cfilnum), #CFileNum)
        ' Send the data
        WS.SendData Dat$
        BytesSent = BytesSent + LOF(CFileNum)
        DoEvents
        MsgBox "File Sent."
        Close #CFileNum
    Else
        ' Get a bit of data
        Dat$ = Input(4500, #CFileNum)
        ' Send the data
        WS.SendData Dat$
        BytesSent = BytesSent + 4500
        DoEvents
    End If

ElseIf Left(Dat$, 4) = "LAST" Then ' The last part of the file
    Dat$ = Mid$(Dat, 6)
    Put #SFileNum, , Dat
    DoEvents
    BytesReceived = BytesReceived + (Len(Dat$) - 5)
    Close #SFileNum
    MsgBox "File Received.", vbInformation, "Done"
    BytesReceived = 0
    
ElseIf Dat$ = "OK" Then ' Send next part
    ' check to see how much is left to send
    a = LOF(CFileNum)
    b = Loc(CFileNum)
    c = a - b
    If c <= 4500 Then
    ' Send the last part
        Dat$ = "LAST:" & Input(c, #CFileNum)
        WS.SendData Dat$
        BytesSent = BytesSent + c
        DoEvents
        Close #CFileNum
        MsgBox "File Sent."
        BytesSent = 0
    Else
        ' send some more
        Dat$ = Input(4500, #CFileNum)
        WS.SendData Dat$
        BytesSent = BytesSent + 4500
        DoEvents
    End If
Else
    ' add the data to the file
    Put #SFileNum, , Dat
    DoEvents
    BytesReceived = BytesReceived + Len(Dat$)
    ' tell the comp to send more
    WS.SendData "OK"
End If

Form1.Caption = BytesSent
Form1.Caption = BytesReceived

End Function

Public Function SendFile(WS As Winsock, Filename As String, FileNum)

CFileNum = FileNum

' Get Just the filename (eg "hello.txt" rather than
' "C:\hello.txt")
For i = 0 To LOF(CFileNum)
    a = Mid(Filename, Len(Filename) - i, 1)
    
    ' Check to see if its a "\"
    If a = "\" Then
        ' Get the filename
        Filename = Right(Filename, i)
        ' Get the filesize
        FileSize = LOF(CFileNum)
        ' Send the data
        WS.SendData "FILE:" & Filename & ":" & FileSize
        Exit Function
    End If
Next i
End Function
