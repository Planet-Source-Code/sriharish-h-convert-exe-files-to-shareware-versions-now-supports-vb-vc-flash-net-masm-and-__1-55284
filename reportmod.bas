Attribute VB_Name = "reportmod"

Option Explicit
Public Sub createreport()

  Dim filenum As Integer

    filenum = FreeFile
    With builder
        Open Left(.Filename.Text, Len(.Filename.Text) - Len(.savedlg.filetitle)) & "reportdata.xml" For Output As filenum
        Print #filenum, LoadResString(101)
        Print #filenum, "<dat>"
        Print #filenum, "<filename>" & .Filename.Text & "</filename>"
        Print #filenum, "<appname>" & .appname.Text & "</appname>"
        Print #filenum, "<trialkey>" & .Trialkey.Text & "</trialkey>"
        If .Option1.Value = True Then
            Print #filenum, "<trialtype>" & "Trial by days" & "</trialtype>"
            Print #filenum, "<trialval>" & .Text2.Text & "</trialval>"
        End If
        If .Option2.Value = True Then
            Print #filenum, "<trialtype>" & "Trial by count" & "</trialtype>"
            Print #filenum, "<trialval>" & .Text3.Text & "</trialval>"
        End If
        If .Option3.Value = True Then
            Print #filenum, "<trialtype>" & "Trial by date" & "</trialtype>"
            Print #filenum, "<trialval>" & Format(.DTPicker1.Value, "MM-DD-YYYY") & "</trialval>"
        End If
        If .CheckDLoader.Value = 1 Then
            Print #filenum, "<dloader>" & "Enabled" & "</dloader>"
          Else
            Print #filenum, "<dloader>" & "Disabled" & "</dloader>"
        End If
        If .Checkloader.Value = 1 Then
            Print #filenum, "<cloader>" & "Enabled" & "</cloader>"
          Else
            Print #filenum, "<cloader>" & "Disabled" & "</cloader>"
        End If
        If .Checkhardware.Value = 1 Then
            Print #filenum, "<hardware>" & "Enabled" & "</hardware>"
          Else
            Print #filenum, "<hardware>" & "Disabled" & "</hardware>"
        End If
        Print #filenum, "<unlock>" & .UnlockKey.Text & "</unlock>"
        Print #filenum, "<email>" & .email.Text & "</email>"
        Print #filenum, "<website>" & .website.Text & "</website>"
        If .Checkfilemod.Value = 1 Then
            Print #filenum, "<filemod>" & "Enabled" & "</filemod>"
          Else
            Print #filenum, "<filemod>" & "Disabled" & "</filemod>"
        End If
        If .CheckIncreaseTrial.Value = 1 Then
            Print #filenum, "<increasetrial>" & "Enabled" & "</increasetrial>"
          Else
            Print #filenum, "<increasetrial>" & "Disabled" & "</increasetrial>"
        End If
        If .CheckOnecopy.Value = 1 Then
            Print #filenum, "<onecopy>" & "Enabled" & "</onecopy>"
          Else
            Print #filenum, "<onecopy>" & "Disabled" & "</onecopy>"
        End If
        If .CheckReset.Value = 1 Then
            Print #filenum, "<reset>" & "Enabled" & "</reset>"
          Else
            Print #filenum, "<reset>" & "Disabled" & "</reset>"
        End If
        If .CheckSDK.Value = 1 Then
            Print #filenum, "<sdk>" & "Enabled" & "</sdk>"
          Else
            Print #filenum, "<sdk>" & "Disabled" & "</sdk>"
        End If
        Print #filenum, "<version>" & .Text1.Text & "</version>"
        Print #filenum, "<blacklist>" & .List1.ListCount & "</blacklist>"
        Print #filenum, "</dat>"
        Close #filenum
        On Error Resume Next
            FileCopy App.path & "\" & "report.html", Left(.Filename.Text, Len(.Filename.Text) - Len(.savedlg.filetitle)) & "report.html"
        End With

End Sub
