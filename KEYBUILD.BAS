Attribute VB_Name = "Keybuild"

Option Explicit
Public Function Convert2Hex(ByVal sAsciiData As String) As String

  Dim lDataLen As Long, iCounter As Long
  Dim sHexData As String, sReturnData As String

    lDataLen = Len(sAsciiData)
    For iCounter = 1 To lDataLen
        sHexData = Hex(Asc(Mid$(sAsciiData, iCounter, 1)))
        If Len(sHexData) < 2 Then sHexData = "0" & sHexData
        sReturnData = sReturnData & sHexData
        sHexData = ""
    Next iCounter
    Convert2Hex = sReturnData

End Function

Public Function convert2Ascii(ByVal sHexData As String) As String

  Dim lDataLen As Long, iCounter As Long
  Dim sAsciiData As String, sReturnData As String

    lDataLen = Len(sHexData)
    For iCounter = 1 To lDataLen Step 2
        sAsciiData = Chr$(CLng("&H" & (Mid$(sHexData, iCounter, 2))))
        sReturnData = sReturnData & sAsciiData
        sAsciiData = ""
    Next iCounter
    convert2Ascii = sReturnData

End Function

Public Sub preparekey1()

  Dim chunk As String
  Dim coredat As String
  Dim chunklen As String

    builder.Crypto.Password = builder.appname.Text & Chr(46) & Chr(46) & Chr(46) & Chr(46) & Chr(46) & Chr(46) & Chr(46) & Chr(46) & _
                              Chr(46) & Chr(46) & Chr(46) & Chr(35) & Chr(51) & Chr(46) & Chr(46) & Chr(46) & _
                              Chr(35) & Chr(50) & Chr(49) & Chr(46) & Chr(46) & Chr(46) & Chr(35) & Chr(49) & _
                              Chr(50) & Chr(46) & Chr(46) & Chr(46) & Chr(51) & Chr(48)
    chunk = builder.Crypto.EncryptData(builder.Trialkey.Text & "#°#" & "0.3")
    chunk = Convert2Hex(chunk)
    chunklen = Len(chunk)
    If Len(chunklen) = 1 Then
        chunk = "1" & Len(chunk) & chunk
    End If
    If Len(chunklen) = 2 Then
        chunk = "2" & Len(chunk) & chunk
    End If
    If Len(chunklen) = 3 Then
        chunk = "3" & Len(chunk) & chunk
    End If
    If Len(chunklen) = 4 Then
        chunk = "4" & Len(chunk) & chunk
    End If
    If Len(chunklen) = 5 Then
        chunk = "5" & Len(chunk) & chunk
    End If
    If Len(chunklen) = 6 Then
        chunk = "6" & Len(chunk) & chunk
    End If
    If builder.Option1.Value = True Then
        coredat = "1" & Len(builder.Text2.Text) & Trim$(builder.Text2.Text)
    End If
    If builder.Option2.Value = True Then
        coredat = "2" & Len(builder.Text3.Text) & Trim$(builder.Text3.Text)
        
    End If
    If builder.Option3.Value = True Then
        coredat = "3" & Len(Format(builder.DTPicker1.Value, "MM-DD-YY")) & Format(builder.DTPicker1.Value, "MM-DD-YY")
    End If
    builder.Crypto.Password = builder.Trialkey.Text & "0.3"
    If Len(coredat) = 3 Then
        coredat = coredat & "T"
    End If
    coredat = builder.Crypto.EncryptData(coredat)
    coredat = Convert2Hex(coredat)
    Licinfo.TKey = chunk & coredat & "="
    Clipboard.SetText Licinfo.TKey
    chunk = Empty
    coredat = Empty
    chunklen = Empty

End Sub

Public Sub preparekey2()

  Dim chunk As String
  Dim coredat As String
  Dim chunklen As String

    builder.Crypto.Password = builder.appname.Text & Chr(46) & Chr(46) & Chr(46) & Chr(35) & Chr(46) & Chr(46) & Chr(46) & Chr(51) & _
                              Chr(46) & Chr(46) & Chr(46) & Chr(51) & Chr(46) & Chr(46) & Chr(46) & Chr(51) & _
                              Chr(46) & Chr(46) & Chr(46) & Chr(77) & Chr(83) & Chr(86) & Chr(66) & Chr(77) & _
                              Chr(54) & Chr(48) & Chr(46) & Chr(68) & Chr(76) & Chr(76)
    chunk = builder.Crypto.EncryptData(builder.UnlockKey.Text & "#°#" & "0.03")
    chunk = Convert2Hex(chunk)
    chunklen = Len(chunk)
    If Len(chunklen) = 1 Then
        chunk = "1" & Len(chunk) & chunk
    End If
    If Len(chunklen) = 2 Then
        chunk = "2" & Len(chunk) & chunk
    End If
    If Len(chunklen) = 3 Then
        chunk = "3" & Len(chunk) & chunk
    End If
    If Len(chunklen) = 4 Then
        chunk = "4" & Len(chunk) & chunk
    End If
    If Len(chunklen) = 5 Then
        chunk = "5" & Len(chunk) & chunk
    End If
    If Len(chunklen) = 6 Then
        chunk = "6" & Len(chunk) & chunk
    End If
    builder.Crypto.Password = builder.UnlockKey.Text & "0.03"
    coredat = builder.Crypto.EncryptData(builder.UnlockKey.Text & "0.03")
    coredat = Convert2Hex(coredat)
    Licinfo.UKey = chunk & coredat
    coredat = Empty
    chunk = Empty

End Sub

':) Ulli's VB Code Formatter V2.14.7 (6/24/2004 4:15:19 PM) 0 + 126 = 126 Lines
