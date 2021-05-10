Imports System.Threading
Imports System.Text
Imports System.IO
Imports System.Net.Mail

'Imports Microsoft.Office.Interop.Excel


Public Class Form1


    
    Function getDFFromSSH(ByVal hostName As String, ByVal userName As String, ByVal password As String, ByRef errCodeInt As Integer, ByRef errMessageStr As String) As String


        Dim ssh As New Chilkat.Ssh()
        Dim port As Integer
        Dim channelNum, posInt As Integer
        Dim termType As String = "xterm"
        Dim widthInChars As Integer = 80
        Dim heightInChars As Integer = 24
        Dim pixWidth As Integer = 0
        Dim pixHeight As Integer = 0
        Dim cmdOutputStr As String = ""
        Dim msgError As String = ""
        Dim n As Integer
        Dim pollTimeoutMs As Integer = 2000

        'Dim success As Boolean = ssh.UnlockComponent("Anything for 30-day trial")
        Dim success As Boolean = ssh.UnlockComponent("IGT")


        If (success <> True) Then
            errCodeInt = -1

            errMessageStr = "Error related to component license: " & ssh.LastErrorText
            cmdOutputStr = ""
            Return cmdOutputStr
            'Exit Function
        End If

       
        port = 22

        success = ssh.Connect(hostName, port)
        If (success <> True) Then
            errCodeInt = -2

            posInt = InStr(ssh.LastErrorText, "connectInner:")
            errMessageStr = Mid(ssh.LastErrorText, posInt)
            errMessageStr = "Error related to the conection: " & errMessageStr
            cmdOutputStr = ""
            Return cmdOutputStr

        End If


        success = ssh.AuthenticatePw(userName, password)
        If (success <> True) Then
            errCodeInt = -3
            errMessageStr = "Error related to the authentication: " & ssh.LastErrorText
            cmdOutputStr = ""
            Return cmdOutputStr
            
        End If

        channelNum = ssh.OpenSessionChannel()
        If (channelNum < 0) Then
            errCodeInt = -4
            errMessageStr = "Error related to OpenSessionChannel: " & ssh.LastErrorText
            cmdOutputStr = ""
            Return cmdOutputStr
        End If

        success = ssh.SendReqPty(channelNum, termType, widthInChars, heightInChars, pixWidth, pixHeight)
        If (success <> True) Then
            errCodeInt = -5
            errMessageStr = "Error related to SendReqPty: " & ssh.LastErrorText
            cmdOutputStr = ""
            Return cmdOutputStr
        End If

        success = ssh.SendReqShell(channelNum)
        If (success <> True) Then
            errCodeInt = -6
            errMessageStr = "Error related to SendReqShell: " & ssh.LastErrorText
            cmdOutputStr = ""
            Return cmdOutputStr
        End If


        n = ssh.ChannelReadAndPoll(channelNum, pollTimeoutMs)
        If (n < 0) Then
            errCodeInt = -7
            errMessageStr = "Error related to ChannelReadAndPoll: " & ssh.LastErrorText
            cmdOutputStr = ""
            Return cmdOutputStr
        End If

        cmdOutputStr = ssh.GetReceivedText(channelNum, "utf-8")
        If (ssh.LastMethodSuccess <> True) Then
            errCodeInt = -8
            errMessageStr = "Error related to GetReceivedText: " & ssh.LastErrorText
            cmdOutputStr = ""
            Return cmdOutputStr
        End If


        errCodeInt = sentStringToSSH(ssh, "df", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            cmdOutputStr = ""
            Return cmdOutputStr
        End If

        ssh.Disconnect()

        Return cmdOutputStr


    End Function


    'Function getDFWEBFromSSH(ByVal hostName As String, ByVal userName As String, ByVal password As String, ByVal hostNameWeb As String, ByRef errCodeInt As Integer, ByRef errMessageStr As String) As String


    '    Dim ssh As New Chilkat.Ssh()
    '    Dim port As Integer
    '    Dim channelNum As Integer
    '    Dim termType As String = "xterm"
    '    Dim widthInChars As Integer = 80
    '    Dim heightInChars As Integer = 24
    '    Dim pixWidth As Integer = 0
    '    Dim pixHeight As Integer = 0
    '    Dim cmdOutputStr As String = ""
    '    Dim msgError As String = ""
    '    Dim n As Integer
    '    Dim pollTimeoutMs As Integer = 2000

    '    'Dim success As Boolean = ssh.UnlockComponent("Anything for 30-day trial")
    '    Dim success As Boolean = ssh.UnlockComponent("IGT")


    '    If (success <> True) Then
    '        errCodeInt = -1
    '        errMessageStr = "Error related to component license: " & ssh.LastErrorText
    '        cmdOutputStr = ""
    '        Return cmdOutputStr
    '        'Exit Function
    '    End If


    '    port = 22

    '    success = ssh.Connect(hostName, port)
    '    If (success <> True) Then
    '        errCodeInt = -2
    '        errMessageStr = "Error related to the conection: " & ssh.LastErrorText
    '        cmdOutputStr = ""
    '        Return cmdOutputStr

    '    End If


    '    success = ssh.AuthenticatePw(userName, password)
    '    If (success <> True) Then
    '        errCodeInt = -3
    '        errMessageStr = "Error related to the authentication: " & ssh.LastErrorText
    '        cmdOutputStr = ""
    '        Return cmdOutputStr

    '    End If

    '    channelNum = ssh.OpenSessionChannel()
    '    If (channelNum < 0) Then
    '        errCodeInt = -4
    '        errMessageStr = "Error related to OpenSessionChannel: " & ssh.LastErrorText
    '        cmdOutputStr = ""
    '        Return cmdOutputStr
    '    End If

    '    success = ssh.SendReqPty(channelNum, termType, widthInChars, heightInChars, pixWidth, pixHeight)
    '    If (success <> True) Then
    '        errCodeInt = -5
    '        errMessageStr = "Error related to SendReqPty: " & ssh.LastErrorText
    '        cmdOutputStr = ""
    '        Return cmdOutputStr
    '    End If

    '    success = ssh.SendReqShell(channelNum)
    '    If (success <> True) Then
    '        errCodeInt = -6
    '        errMessageStr = "Error related to SendReqShell: " & ssh.LastErrorText
    '        cmdOutputStr = ""
    '        Return cmdOutputStr
    '    End If


    '    n = ssh.ChannelReadAndPoll(channelNum, pollTimeoutMs)
    '    If (n < 0) Then
    '        errCodeInt = -7
    '        errMessageStr = "Error related to ChannelReadAndPoll: " & ssh.LastErrorText
    '        cmdOutputStr = ""
    '        Return cmdOutputStr
    '    End If

    '    cmdOutputStr = ssh.GetReceivedText(channelNum, "utf-8")
    '    If (ssh.LastMethodSuccess <> True) Then
    '        errCodeInt = -8
    '        errMessageStr = "Error related to GetReceivedText: " & ssh.LastErrorText
    '        cmdOutputStr = ""
    '        Return cmdOutputStr
    '    End If


    '    errCodeInt = sentStringToSSH(ssh, "df", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
    '    If errCodeInt <> 0 Then
    '        errCodeInt = -9
    '        errMessageStr = ssh.LastErrorText
    '        cmdOutputStr = ""
    '        Return cmdOutputStr
    '    End If

    '    ssh.Disconnect()

    '    Return cmdOutputStr


    'End Function






    Function sentStringToSSH(ByVal ssh As Chilkat.Ssh, ByVal strText As String, ByVal channelNum As Integer, ByVal pollTimeoutMs As Integer, ByRef cmdOutputStr As String, ByRef msgError As String) As Integer
        Dim success As Boolean
        Dim n As Integer

        success = ssh.ChannelSendString(channelNum, strText & vbCrLf, "utf-8")
        If (success <> True) Then
            msgError = "ChannelSendString Error: " + ssh.LastErrorText
            Return -1
        End If

        n = ssh.ChannelReadAndPoll(channelNum, pollTimeoutMs)
        If (n < 0) Then
            msgError = "ChannelReadAndPoll Error: " + ssh.LastErrorText
            Return -2
        End If

        cmdOutputStr = ssh.GetReceivedText(channelNum, "utf-8")
        If (ssh.LastMethodSuccess <> True) Then
            msgError = "GetReceivedText Error: " + ssh.LastErrorText
            Return -3
        Else
            Return 0
        End If

    End Function


    Private Sub readFileDF(ByVal fileName As String, ByRef listFolders As ArrayList, ByRef listPercentage As ArrayList, ByRef msgError As String)


        Dim reader1 As StreamReader
        Dim strLine, strValue, strMounted As String

        Try
            reader1 = New StreamReader(fileName, Encoding.UTF7)
            strLine = ""

            Do While Not (strLine Is Nothing)
                strLine = reader1.ReadLine()
                strValue = Trim((Mid(strLine, 52, 3)))
                strMounted = Trim(Mid(strLine, 58, 20))

                If ((InStr(1, strMounted, "/") = 1)) Then
                    strValue = Trim((Mid(strLine, 53, 3)))
                    strMounted = Trim(Mid(strLine, 59, 20))
                End If

                If IsNumeric(strValue) Then

                    listPercentage.Add(CInt(strValue))
                    If strMounted = "" Then
                        strMounted = "root"
                    End If
                    listFolders.Add(strMounted)
                End If

            Loop
            reader1.Close()
        Catch ex As Exception
            MsgBox("", MsgBoxStyle.Information, "Error")
        End Try
    End Sub

    
    
    Private Sub okBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles okBtn.Click


        Me.okBtn.Enabled = False
        Me.cancelBtn.Enabled = False
        Me.Cursor = Cursors.WaitCursor

        getDFFromAllServersSSH()
        fillExcelAllServers()

        Me.Cursor = Cursors.Default

    End Sub

    Private Sub cancelBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cancelBtn.Click
        Me.Close()
    End Sub


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ProgressBar1.Maximum = 100
        okBtn.Visible = False
        cancelBtn.Visible = False
        Timer1.Enabled = True
    End Sub

    Private Sub getDFFromAllServersSSH()

        Dim errCodeInt As Integer = 0
        Dim resCode As Integer = 0
        Dim errMessageStr As String = ""
        Dim file As System.IO.StreamWriter
        Dim infoStr As String
        Dim listPercentage As New ArrayList
        Dim listFolders As New ArrayList
        Dim fecha As Date
        Dim monthStr, dayStr, yearStr, hourStr, minStr, secStr, fileLogNameStr As String
        Dim flagError = False


        fecha = Format(Now, "MM/dd/yyyy HH:mm:ss")
        monthStr = CStr(fecha.Month).PadLeft(2, "0")
        dayStr = CStr(fecha.Day).PadLeft(2, "0")
        yearStr = CStr(fecha.Year).PadLeft(4, "0")
        hourStr = CStr(fecha.Hour).PadLeft(2, "0")
        minStr = CStr(fecha.Minute).PadLeft(2, "0")
        secStr = CStr(fecha.Second).PadLeft(2, "0")
        fileLogNameStr = "logError_" & monthStr & dayStr & yearStr & ".txt"


        Me.RichTextBox1.Text = "Getting DF ESTE1..." & vbCrLf
        Me.Refresh()
        infoStr = getDFFromSSH("10.1.5.10", "xfer", "Welcome1", errCodeInt, errMessageStr)

        If errCodeInt <> 0 Then
            Select Case errCodeInt
                Case -1 'Error related to component license
                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine(errMessageStr)
                    file.Close()

                    resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

                    Me.Dispose()


                Case Else '-2, -3, -4, -5, -6, -7, -8, -9

                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine("ESTE1: " & errMessageStr)
                    file.Close()
                    flagError = True
                    resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

            End Select
        End If


        file = My.Computer.FileSystem.OpenTextFileWriter("DF_ESTE1.txt", False)
        file.WriteLine(infoStr)
        file.Close()

        Me.Refresh()

        Me.RichTextBox1.Text = "Getting DF ESTE2..." & vbCrLf
        Me.Refresh()
        infoStr = getDFFromSSH("10.1.5.11", "xfer", "Welcome1", errCodeInt, errMessageStr)
        If errCodeInt <> 0 Then
            Select Case errCodeInt
                Case -1 'Error related to component license
                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine(errMessageStr)
                    file.Close()

                    resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

                    Me.Dispose()


                Case Else '-2, -3, -4, -5, -6, -7, -8, -9

                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine("ESTE2: " & errMessageStr)
                    file.Close()
                    flagError = True
                    resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

            End Select
        End If

        file = My.Computer.FileSystem.OpenTextFileWriter("DF_ESTE2.txt", False)
        file.WriteLine(infoStr)
        file.Close()

        Me.Refresh()

        Me.RichTextBox1.Text = "Getting DF ESTE3..." & vbCrLf
        Me.Refresh()
        infoStr = getDFFromSSH("10.1.5.12", "xfer", "Welcome1", errCodeInt, errMessageStr)
        If errCodeInt <> 0 Then
            Select Case errCodeInt
                Case -1 'Error related to component license
                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine(errMessageStr)
                    file.Close()

                    resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

                    Me.Dispose()


                Case Else '-2, -3, -4, -5, -6, -7, -8, -9

                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine("ESTE3: " & errMessageStr)
                    file.Close()
                    flagError = True
                    resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

            End Select
        End If

        file = My.Computer.FileSystem.OpenTextFileWriter("DF_ESTE3.txt", False)
        file.WriteLine(infoStr)
        file.Close()

        Me.Refresh()

        Me.RichTextBox1.Text = "Getting DF ESTE4..." & vbCrLf
        Me.Refresh()
        infoStr = getDFFromSSH("10.2.5.13", "xfer", "Welcome1", errCodeInt, errMessageStr)
        If errCodeInt <> 0 Then
            Select Case errCodeInt
                Case -1 'Error related to component license
                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine(errMessageStr)
                    file.Close()

                    resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

                    Me.Dispose()


                Case Else '-2, -3, -4, -5, -6, -7, -8, -9

                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine("ESTE4: " & errMessageStr)
                    file.Close()
                    flagError = True

                    resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

            End Select
        End If

        file = My.Computer.FileSystem.OpenTextFileWriter("DF_ESTE4.txt", False)
        file.WriteLine(infoStr)
        file.Close()

        Me.Refresh()

        Me.RichTextBox1.Text = "Getting DF ESTE5..." & vbCrLf
        Me.Refresh()
        infoStr = getDFFromSSH("10.2.5.14", "xfer", "Welcome1", errCodeInt, errMessageStr)
        If errCodeInt <> 0 Then
            Select Case errCodeInt
                Case -1 'Error related to component license
                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine(errMessageStr)
                    file.Close()

                    resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

                    Me.Dispose()


                Case Else '-2, -3, -4, -5, -6, -7, -8, -9

                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine("ESTE5: " & errMessageStr)
                    file.Close()
                    flagError = True

                    resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

            End Select
        End If

        file = My.Computer.FileSystem.OpenTextFileWriter("DF_ESTE5.txt", False)
        file.WriteLine(infoStr)
        file.Close()

        Me.Refresh()

        Me.RichTextBox1.Text = "Getting DF ESC1..." & vbCrLf
        Me.Refresh()
        infoStr = getDFFromSSH("10.1.5.61", "ricie", "@m3t5", errCodeInt, errMessageStr)
        If errCodeInt <> 0 Then
            Select Case errCodeInt
                Case -1 'Error related to component license
                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine(errMessageStr)
                    file.Close()

                    resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

                    Me.Dispose()


                Case Else '-2, -3, -4, -5, -6, -7, -8, -9

                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine("ESC1: " & errMessageStr)
                    file.Close()
                    flagError = True

            End Select
        End If

        file = My.Computer.FileSystem.OpenTextFileWriter("DF_ESC1.txt", False)
        file.WriteLine(infoStr)
        file.Close()

        'Me.RichTextBox1.Text = "Getting DF ESC2..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.1.5.62", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    Select Case errCodeInt
        '        Case -1 'Error related to component license
        '            file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
        '            file.WriteLine(errMessageStr)
        '            file.Close()

        '            resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

        '            Me.Dispose()


        '        Case Else '-2, -3, -4, -5, -6, -7, -8, -9

        '            file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
        '            file.WriteLine("ESC2: " & errMessageStr)
        '            file.Close()
        '            flagError = True

        '    End Select
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_ESC2.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()

        Me.RichTextBox1.Text = "Getting DF ESC3..." & vbCrLf
        Me.Refresh()
        infoStr = getDFFromSSH("10.1.5.63", "ricie", "@m3t5", errCodeInt, errMessageStr)
        If errCodeInt <> 0 Then
            Select Case errCodeInt
                Case -1 'Error related to component license
                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine(errMessageStr)
                    file.Close()

                    resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

                    Me.Dispose()


                Case Else '-2, -3, -4, -5, -6, -7, -8, -9

                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine("ESC3: " & errMessageStr)
                    file.Close()
                    flagError = True

            End Select
        End If

        file = My.Computer.FileSystem.OpenTextFileWriter("DF_ESC3.txt", False)
        file.WriteLine(infoStr)
        file.Close()

        Me.RichTextBox1.Text = "Getting DF ESC4..." & vbCrLf
        Me.Refresh()
        infoStr = getDFFromSSH("10.2.5.64", "ricie", "@m3t5", errCodeInt, errMessageStr)
        If errCodeInt <> 0 Then
            Select Case errCodeInt
                Case -1 'Error related to component license
                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine(errMessageStr)
                    file.Close()

                    resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

                    Me.Dispose()


                Case Else '-2, -3, -4, -5, -6, -7, -8, -9

                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine("ESC4: " & errMessageStr)
                    file.Close()
                    flagError = True

            End Select
        End If

        file = My.Computer.FileSystem.OpenTextFileWriter("DF_ESC4.txt", False)
        file.WriteLine(infoStr)
        file.Close()


        'Me.RichTextBox1.Text = "Getting DF ESC5..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.2.5.65", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    Select Case errCodeInt
        '        Case -1 'Error related to component license
        '            file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
        '            file.WriteLine(errMessageStr)
        '            file.Close()

        '            resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

        '            Me.Dispose()


        '        Case Else '-2, -3, -4, -5, -6, -7, -8, -9

        '            file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
        '            file.WriteLine("ESC5: " & errMessageStr)
        '            file.Close()
        '            flagError = True

        '    End Select
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_ESC5.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()


        'Me.RichTextBox1.Text = "Getting DF ESC6..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.2.5.65", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    Select Case errCodeInt
        '        Case -1 'Error related to component license
        '            file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
        '            file.WriteLine(errMessageStr)
        '            file.Close()

        '            resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

        '            Me.Dispose()


        '        Case Else '-2, -3, -4, -5, -6, -7, -8, -9

        '            file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
        '            file.WriteLine("ESC6: " & errMessageStr)
        '            file.Close()
        '            flagError = True

        '    End Select
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_ESC6.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()



        Me.RichTextBox1.Text = "Getting DF B2BAPP1..." & vbCrLf
        Me.Refresh()
        infoStr = getDFFromSSH("10.1.5.103", "ricie", "@m3t5", errCodeInt, errMessageStr)
        If errCodeInt <> 0 Then
            Select Case errCodeInt
                Case -1 'Error related to component license
                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine(errMessageStr)
                    file.Close()

                    resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

                    Me.Dispose()


                Case Else '-2, -3, -4, -5, -6, -7, -8, -9

                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine("B2BAPP1: " & errMessageStr)
                    file.Close()
                    flagError = True

            End Select
        End If

        file = My.Computer.FileSystem.OpenTextFileWriter("DF_B2BAPP1.txt", False)
        file.WriteLine(infoStr)
        file.Close()


        Me.RichTextBox1.Text = "Getting DF B2BAPP2..." & vbCrLf
        Me.Refresh()
        infoStr = getDFFromSSH("10.1.5.104", "ricie", "@m3t5", errCodeInt, errMessageStr)
        If errCodeInt <> 0 Then
            Select Case errCodeInt
                Case -1 'Error related to component license
                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine(errMessageStr)
                    file.Close()

                    resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

                    Me.Dispose()


                Case Else '-2, -3, -4, -5, -6, -7, -8, -9

                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine("B2BAPP2: " & errMessageStr)
                    file.Close()
                    flagError = True

            End Select
        End If

        file = My.Computer.FileSystem.OpenTextFileWriter("DF_B2BAPP2.txt", False)
        file.WriteLine(infoStr)
        file.Close()


        Me.RichTextBox1.Text = "Getting DF B2BWEB1..." & vbCrLf
        Me.Refresh()
        infoStr = getDFFromSSH("10.1.4.121", "ricie", "@m3t5", errCodeInt, errMessageStr)
        If errCodeInt <> 0 Then
            Select Case errCodeInt
                Case -1 'Error related to component license
                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine(errMessageStr)
                    file.Close()

                    resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

                    Me.Dispose()


                Case Else '-2, -3, -4, -5, -6, -7, -8, -9

                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine("B2BWEB1: " & errMessageStr)
                    file.Close()
                    flagError = True

            End Select
        End If

        file = My.Computer.FileSystem.OpenTextFileWriter("DF_B2BWEB1.txt", False)
        file.WriteLine(infoStr)
        file.Close()


        Me.RichTextBox1.Text = "Getting DF B2BWEB2..." & vbCrLf
        Me.Refresh()
        infoStr = getDFFromSSH("10.1.4.122", "ricie", "@m3t5", errCodeInt, errMessageStr)
        If errCodeInt <> 0 Then
            Select Case errCodeInt
                Case -1 'Error related to component license
                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine(errMessageStr)
                    file.Close()

                    resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

                    Me.Dispose()


                Case Else '-2, -3, -4, -5, -6, -7, -8, -9

                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine("B2BWEB2: " & errMessageStr)
                    file.Close()
                    flagError = True

            End Select
        End If

        file = My.Computer.FileSystem.OpenTextFileWriter("DF_B2BWEB2.txt", False)
        file.WriteLine(infoStr)
        file.Close()



        Me.RichTextBox1.Text = "Getting DF PWA1..." & vbCrLf
        Me.Refresh()
        infoStr = getDFFromSSH("10.1.5.77", "ricie", "@m3t5", errCodeInt, errMessageStr)
        If errCodeInt <> 0 Then
            Select Case errCodeInt
                Case -1 'Error related to component license
                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine(errMessageStr)
                    file.Close()

                    resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

                    Me.Dispose()


                Case Else '-2, -3, -4, -5, -6, -7, -8, -9

                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine("PWA1: " & errMessageStr)
                    file.Close()
                    flagError = True

            End Select
        End If

        file = My.Computer.FileSystem.OpenTextFileWriter("DF_PWA1.txt", False)
        file.WriteLine(infoStr)
        file.Close()


        Me.RichTextBox1.Text = "Getting DF PWA2..." & vbCrLf
        Me.Refresh()
        infoStr = getDFFromSSH("10.1.5.78", "ricie", "@m3t5", errCodeInt, errMessageStr)
        If errCodeInt <> 0 Then
            Select Case errCodeInt
                Case -1 'Error related to component license
                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine(errMessageStr)
                    file.Close()

                    resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

                    Me.Dispose()


                Case Else '-2, -3, -4, -5, -6, -7, -8, -9

                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine("PWA2: " & errMessageStr)
                    file.Close()
                    flagError = True

            End Select
        End If

        file = My.Computer.FileSystem.OpenTextFileWriter("DF_PWA2.txt", False)
        file.WriteLine(infoStr)
        file.Close()



        Me.RichTextBox1.Text = "Getting DF PPAPP1..." & vbCrLf
        Me.Refresh()
        infoStr = getDFFromSSH("10.1.5.166", "ricie", "@m3t5", errCodeInt, errMessageStr)
        If errCodeInt <> 0 Then
            Select Case errCodeInt
                Case -1 'Error related to component license
                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine(errMessageStr)
                    file.Close()

                    resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

                    Me.Dispose()


                Case Else '-2, -3, -4, -5, -6, -7, -8, -9

                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine("PPAPP1: " & errMessageStr)
                    file.Close()
                    flagError = True

            End Select
        End If

        file = My.Computer.FileSystem.OpenTextFileWriter("DF_PPAPP1.txt", False)
        file.WriteLine(infoStr)
        file.Close()


        Me.RichTextBox1.Text = "Getting DF PPAPP2..." & vbCrLf
        Me.Refresh()
        infoStr = getDFFromSSH("10.1.5.167", "ricie", "@m3t5", errCodeInt, errMessageStr)
        If errCodeInt <> 0 Then
            Select Case errCodeInt
                Case -1 'Error related to component license
                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine(errMessageStr)
                    file.Close()

                    resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

                    Me.Dispose()


                Case Else '-2, -3, -4, -5, -6, -7, -8, -9

                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine("PPAPP2: " & errMessageStr)
                    file.Close()
                    flagError = True

            End Select
        End If

        file = My.Computer.FileSystem.OpenTextFileWriter("DF_PPAPP2.txt", False)
        file.WriteLine(infoStr)
        file.Close()

        Me.Refresh()

        Me.RichTextBox1.Text = "Getting DF PPWEB1..." & vbCrLf
        Me.Refresh()
        infoStr = getDFFromSSH("10.1.218.170", "ricie", "@m3t5", errCodeInt, errMessageStr)
        If errCodeInt <> 0 Then
            Select Case errCodeInt
                Case -1 'Error related to component license
                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine(errMessageStr)
                    file.Close()

                    resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

                    Me.Dispose()


                Case Else '-2, -3, -4, -5, -6, -7, -8, -9

                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine("PPWEB1: " & errMessageStr)
                    file.Close()
                    flagError = True

            End Select
        End If

        file = My.Computer.FileSystem.OpenTextFileWriter("DF_PPWEB1.txt", False)
        file.WriteLine(infoStr)
        file.Close()


        Me.RichTextBox1.Text = "Getting DF PPWEB2..." & vbCrLf
        Me.Refresh()
        infoStr = getDFFromSSH("10.1.218.171", "ricie", "@m3t5", errCodeInt, errMessageStr)
        If errCodeInt <> 0 Then
            Select Case errCodeInt
                Case -1 'Error related to component license
                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine(errMessageStr)
                    file.Close()

                    resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

                    Me.Dispose()


                Case Else '-2, -3, -4, -5, -6, -7, -8, -9

                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine("PPWEB2: " & errMessageStr)
                    file.Close()
                    flagError = True

            End Select
        End If

        file = My.Computer.FileSystem.OpenTextFileWriter("DF_PPWEB2.txt", False)
        file.WriteLine(infoStr)
        file.Close()

        '===================================================================================

        Me.RichTextBox1.Text = "Getting DF NYDB1..." & vbCrLf
        Me.Refresh()
        infoStr = getDFFromSSH("10.1.5.91", "ricie", "@m3t5", errCodeInt, errMessageStr)
        If errCodeInt <> 0 Then
            Select Case errCodeInt
                Case -1 'Error related to component license
                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine(errMessageStr)
                    file.Close()

                    resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

                    Me.Dispose()


                Case Else '-2, -3, -4, -5, -6, -7, -8, -9

                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine("NYDB1: " & errMessageStr)
                    file.Close()
                    flagError = True

            End Select
        End If

        file = My.Computer.FileSystem.OpenTextFileWriter("DF_DB1.txt", False)
        file.WriteLine(infoStr)
        file.Close()

        Me.RichTextBox1.Text = "Getting DF NYDB2..." & vbCrLf
        Me.Refresh()
        infoStr = getDFFromSSH("10.1.5.92", "ricie", "@m3t5", errCodeInt, errMessageStr)
        If errCodeInt <> 0 Then
            Select Case errCodeInt
                Case -1 'Error related to component license
                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine(errMessageStr)
                    file.Close()

                    resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

                    Me.Dispose()


                Case Else '-2, -3, -4, -5, -6, -7, -8, -9

                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine("NYDB2: " & errMessageStr)
                    file.Close()
                    flagError = True

            End Select
        End If

        file = My.Computer.FileSystem.OpenTextFileWriter("DF_DB2.txt", False)
        file.WriteLine(infoStr)
        file.Close()

        Me.RichTextBox1.Text = "Getting DF NYDB3..." & vbCrLf
        Me.Refresh()
        infoStr = getDFFromSSH("10.2.5.93", "ricie", "@m3t5", errCodeInt, errMessageStr)
        If errCodeInt <> 0 Then
            Select Case errCodeInt
                Case -1 'Error related to component license
                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine(errMessageStr)
                    file.Close()

                    resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

                    Me.Dispose()


                Case Else '-2, -3, -4, -5, -6, -7, -8, -9

                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine("NYDB3: " & errMessageStr)
                    file.Close()
                    flagError = True

            End Select
        End If

        file = My.Computer.FileSystem.OpenTextFileWriter("DF_DB3.txt", False)
        file.WriteLine(infoStr)
        file.Close()

        Me.Refresh()

        Me.RichTextBox1.Text = "Getting DF NYDB4..." & vbCrLf
        Me.Refresh()
        infoStr = getDFFromSSH("10.2.5.94", "ricie", "@m3t5", errCodeInt, errMessageStr)
        If errCodeInt <> 0 Then
            Select Case errCodeInt
                Case -1 'Error related to component license
                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine(errMessageStr)
                    file.Close()

                    resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

                    Me.Dispose()


                Case Else '-2, -3, -4, -5, -6, -7, -8, -9

                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine("NYDB4: " & errMessageStr)
                    file.Close()
                    flagError = True

            End Select
        End If

        file = My.Computer.FileSystem.OpenTextFileWriter("DF_DB4.txt", False)
        file.WriteLine(infoStr)
        file.Close()


        'Me.RichTextBox1.Text = "Getting DF NYPPDB1..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.1.5.59", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    Select Case errCodeInt
        '        Case -1 'Error related to component license
        '            file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
        '            file.WriteLine(errMessageStr)
        '            file.Close()

        '            resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

        '            Me.Dispose()


        '        Case Else '-2, -3, -4, -5, -6, -7, -8, -9

        '            file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
        '            file.WriteLine("NYPPDB1: " & errMessageStr)
        '            file.Close()
        '            flagError = True

        '    End Select
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_PPDB1.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()


        'Me.RichTextBox1.Text = "Getting DF NYPPDB2..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.1.5.60", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    Select Case errCodeInt
        '        Case -1 'Error related to component license
        '            file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
        '            file.WriteLine(errMessageStr)
        '            file.Close()

        '            resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

        '            Me.Dispose()


        '        Case Else '-2, -3, -4, -5, -6, -7, -8, -9

        '            file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
        '            file.WriteLine("NYPPDB2: " & errMessageStr)
        '            file.Close()
        '            flagError = True

        '    End Select
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_PPDB2.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()


        'Me.Refresh()

        'Me.RichTextBox1.Text = "Getting DF NYPPDB3..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.2.5.89", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    Select Case errCodeInt
        '        Case -1 'Error related to component license
        '            file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
        '            file.WriteLine(errMessageStr)
        '            file.Close()

        '            resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

        '            Me.Dispose()


        '        Case Else '-2, -3, -4, -5, -6, -7, -8, -9

        '            file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
        '            file.WriteLine("NYPPDB3: " & errMessageStr)
        '            file.Close()
        '            flagError = True

        '    End Select
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_PPDB3.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()


        'Me.RichTextBox1.Text = "Getting DF NYPPDB4..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.2.5.90", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    Select Case errCodeInt
        '        Case -1 'Error related to component license
        '            file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
        '            file.WriteLine(errMessageStr)
        '            file.Close()

        '            resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

        '            Me.Dispose()


        '        Case Else '-2, -3, -4, -5, -6, -7, -8, -9

        '            file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
        '            file.WriteLine("NYPPDB4: " & errMessageStr)
        '            file.Close()
        '            flagError = True

        '    End Select
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_PPDB4.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()

        '======================================================================================

        'Me.RichTextBox1.Text = "Getting DF NYB2BAPP3..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.2.5.106", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    Select Case errCodeInt
        '        Case -1 'Error related to component license
        '            file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
        '            file.WriteLine(errMessageStr)
        '            file.Close()

        '            resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

        '            Me.Dispose()


        '        Case Else '-2, -3, -4, -5, -6, -7, -8, -9

        '            file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
        '            file.WriteLine("NYB2BAPP3: " & errMessageStr)
        '            file.Close()
        '            flagError = True

        '    End Select
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_B2BAPP3.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()

        '=================================================

        'NEW
        'Me.RichTextBox1.Text = "Getting DF NYB2BWEB3..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFWEBFromSSH("10.2.5.106", "ricie", "@m3t5", "nyb2bweb3", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    Select Case errCodeInt
        '        Case -1 'Error related to component license
        '            file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
        '            file.WriteLine(errMessageStr)
        '            file.Close()

        '            resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

        '            Me.Dispose()


        '        Case Else '-2, -3, -4, -5, -6, -7, -8, -9

        '            file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
        '            file.WriteLine("NYB2BWEB3: " & errMessageStr)
        '            file.Close()
        '            resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

        '    End Select
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_B2BAPP3.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()




        '======================================================

        'Me.RichTextBox1.Text = "Getting DF NYB2BAPP4..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.2.5.107", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    Select Case errCodeInt
        '        Case -1 'Error related to component license
        '            file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
        '            file.WriteLine(errMessageStr)
        '            file.Close()

        '            resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

        '            Me.Dispose()


        '        Case Else '-2, -3, -4, -5, -6, -7, -8, -9

        '            file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
        '            file.WriteLine("NYB2BAPP4: " & errMessageStr)
        '            file.Close()
        '            flagError = True

        '    End Select
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_B2BAPP4.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()


        Me.RichTextBox1.Text = "Getting DF NYPWA3..." & vbCrLf
        Me.Refresh()
        infoStr = getDFFromSSH("10.2.5.79", "ricie", "@m3t5", errCodeInt, errMessageStr)
        If errCodeInt <> 0 Then
            Select Case errCodeInt
                Case -1 'Error related to component license
                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine(errMessageStr)
                    file.Close()

                    resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

                    Me.Dispose()


                Case Else '-2, -3, -4, -5, -6, -7, -8, -9

                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine("NYPWA3: " & errMessageStr)
                    file.Close()
                    flagError = True

            End Select
        End If

        file = My.Computer.FileSystem.OpenTextFileWriter("DF_PWA3.txt", False)
        file.WriteLine(infoStr)
        file.Close()

        Me.RichTextBox1.Text = "Getting DF NYPWA4..." & vbCrLf
        Me.Refresh()
        infoStr = getDFFromSSH("10.2.5.80", "ricie", "@m3t5", errCodeInt, errMessageStr)
        If errCodeInt <> 0 Then
            Select Case errCodeInt
                Case -1 'Error related to component license
                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine(errMessageStr)
                    file.Close()

                    resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

                    Me.Dispose()


                Case Else '-2, -3, -4, -5, -6, -7, -8, -9

                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine("NYPWA4: " & errMessageStr)
                    file.Close()
                    flagError = True

            End Select
        End If

        file = My.Computer.FileSystem.OpenTextFileWriter("DF_PWA4.txt", False)
        file.WriteLine(infoStr)
        file.Close()



        Me.RichTextBox1.Text = "Getting DF NYPPAPP3..." & vbCrLf
        Me.Refresh()
        infoStr = getDFFromSSH("10.2.5.168", "ricie", "@m3t5", errCodeInt, errMessageStr)
        If errCodeInt <> 0 Then
            Select Case errCodeInt
                Case -1 'Error related to component license
                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine(errMessageStr)
                    file.Close()

                    resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

                    Me.Dispose()


                Case Else '-2, -3, -4, -5, -6, -7, -8, -9

                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine("NYPPAPP3: " & errMessageStr)
                    file.Close()
                    flagError = True

            End Select
        End If

        file = My.Computer.FileSystem.OpenTextFileWriter("DF_PPAPP3.txt", False)
        file.WriteLine(infoStr)
        file.Close()

        Me.RichTextBox1.Text = "Getting DF NYPPAPP4..." & vbCrLf
        Me.Refresh()
        infoStr = getDFFromSSH("10.2.5.169", "ricie", "@m3t5", errCodeInt, errMessageStr)
        If errCodeInt <> 0 Then
            Select Case errCodeInt
                Case -1 'Error related to component license
                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine(errMessageStr)
                    file.Close()

                    resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")

                    Me.Dispose()


                Case Else '-2, -3, -4, -5, -6, -7, -8, -9

                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine("NYPPAPP4: " & errMessageStr)
                    file.Close()
                    flagError = True
            End Select
        End If

        file = My.Computer.FileSystem.OpenTextFileWriter("DF_PPAPP4.txt", False)
        file.WriteLine(infoStr)
        file.Close()

        '=====================================================================================================

        Me.RichTextBox1.Text = "Getting DF NYSFTP1..." & vbCrLf
        Me.Refresh()
        infoStr = getDFFromSSH("10.1.5.182", "ricie", "@m3t5", errCodeInt, errMessageStr)
        If errCodeInt <> 0 Then
            Select Case errCodeInt
                Case -1 'Error related to component license
                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine(errMessageStr)
                    file.Close()
                    resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")
                    Me.Dispose()

                Case Else '-2, -3, -4, -5, -6, -7, -8, -9
                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine("NYSFTP1: " & errMessageStr)
                    file.Close()
                    flagError = True
            End Select
        End If
        file = My.Computer.FileSystem.OpenTextFileWriter("NYSFTP1.txt", False)
        file.WriteLine(infoStr)
        file.Close()



        Me.RichTextBox1.Text = "Getting DF NYSFTP2..." & vbCrLf
        Me.Refresh()
        infoStr = getDFFromSSH("10.2.5.183", "ricie", "@m3t5", errCodeInt, errMessageStr)
        If errCodeInt <> 0 Then
            Select Case errCodeInt
                Case -1 'Error related to component license
                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine(errMessageStr)
                    file.Close()
                    resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")
                    Me.Dispose()

                Case Else '-2, -3, -4, -5, -6, -7, -8, -9
                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine("NYSFTP2: " & errMessageStr)
                    file.Close()
                    flagError = True
            End Select
        End If
        file = My.Computer.FileSystem.OpenTextFileWriter("NYSFTP2.txt", False)
        file.WriteLine(infoStr)
        file.Close()



        Me.RichTextBox1.Text = "Getting DF NYDMZSFTP1..." & vbCrLf
        Me.Refresh()
        infoStr = getDFFromSSH("10.1.6.184", "ricie", "@m3t5", errCodeInt, errMessageStr)
        If errCodeInt <> 0 Then
            Select Case errCodeInt
                Case -1 'Error related to component license
                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine(errMessageStr)
                    file.Close()
                    resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log.")
                    Me.Dispose()

                Case Else '-2, -3, -4, -5, -6, -7, -8, -9
                    file = My.Computer.FileSystem.OpenTextFileWriter(fileLogNameStr, True)
                    file.WriteLine("NYDMZSFTP1: " & errMessageStr)
                    file.Close()
                    flagError = True
            End Select
        End If
        file = My.Computer.FileSystem.OpenTextFileWriter("NYDMZSFTP1.txt", False)
        file.WriteLine(infoStr)
        file.Close()



        Me.Refresh()

        If (flagError) Then
            resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com", fileLogNameStr, "Error DF app", "Attached is the log with the error description.")
        End If



    End Sub

    Private Sub fillExcelAllServers()

        Dim errCodeInt As Integer = 0
        Dim errMessageStr As String = ""
        Dim listPercentage As New ArrayList
        Dim listFolders As New ArrayList
        Dim xlibro As Microsoft.Office.Interop.Excel.Application
        Dim indexReturn As Integer
        Dim valueStr As String
        Dim pathToFindStr, timeStr As String
        Dim fecha As Date
        Dim monthStr, dayStr, yearStr, hourStr, minStr, secStr, fileNameStr, strPathFile, strPathRemoteFile As String

        fecha = Format(Now, "MM/dd/yyyy")
        timeStr = Format(Now, "HH:mm:ss")


        Me.RichTextBox1.Text = "Filling excel form..." & vbCrLf
        Me.Refresh()


        strPathFile = System.Windows.Forms.Application.StartupPath
        fileNameStr = strPathFile & "\Disk Space Check.xlsx"

        xlibro = CreateObject("Excel.Application")
        xlibro.Workbooks.Open(fileNameStr)
        xlibro.Visible = False
        xlibro.Sheets("ESTE").Select()
        xlibro.Cells(3, 1) = fecha
        xlibro.Cells(3, 2) = timeStr

        readFileDF("DF_ESTE1.txt", listFolders, listPercentage, errMessageStr)
        For i = 6 To 35
            pathToFindStr = xlibro.Cells(i, 1).Value
            indexReturn = listFolders.IndexOf(pathToFindStr)
            valueStr = listPercentage.Item(indexReturn)
            xlibro.Cells(i, 2) = valueStr

        Next
        listFolders.Clear()
        listPercentage.Clear()

        readFileDF("DF_ESTE2.txt", listFolders, listPercentage, errMessageStr)

        For i = 6 To 35
            pathToFindStr = xlibro.Cells(i, 1).Value
            indexReturn = listFolders.IndexOf(pathToFindStr)
            valueStr = listPercentage.Item(indexReturn)
            xlibro.Cells(i, 3) = valueStr

        Next

        listFolders.Clear()
        listPercentage.Clear()

        readFileDF("DF_ESTE3.txt", listFolders, listPercentage, errMessageStr)

        For i = 6 To 35
            pathToFindStr = xlibro.Cells(i, 1).Value
            indexReturn = listFolders.IndexOf(pathToFindStr)
            valueStr = listPercentage.Item(indexReturn)
            xlibro.Cells(i, 4) = valueStr

        Next

        listFolders.Clear()
        listPercentage.Clear()

        readFileDF("DF_ESTE4.txt", listFolders, listPercentage, errMessageStr)
        For i = 6 To 35
            pathToFindStr = xlibro.Cells(i, 1).Value
            indexReturn = listFolders.IndexOf(pathToFindStr)
            valueStr = listPercentage.Item(indexReturn)
            xlibro.Cells(i, 5) = valueStr

        Next

        listFolders.Clear()
        listPercentage.Clear()

        readFileDF("DF_ESTE5.txt", listFolders, listPercentage, errMessageStr)

        For i = 6 To 35
            pathToFindStr = xlibro.Cells(i, 1).Value
            indexReturn = listFolders.IndexOf(pathToFindStr)
            valueStr = listPercentage.Item(indexReturn)
            xlibro.Cells(i, 6) = valueStr
        Next

        listFolders.Clear()
        listPercentage.Clear()

        '============================================================================================

        xlibro.Sheets("ESC").Select()
        xlibro.Cells(3, 1) = fecha
        xlibro.Cells(3, 2) = timeStr

        readFileDF("DF_ESC1.txt", listFolders, listPercentage, errMessageStr)
        For i = 6 To 13
            pathToFindStr = xlibro.Cells(i, 1).Value
            indexReturn = listFolders.IndexOf(pathToFindStr)
            valueStr = listPercentage.Item(indexReturn)
            xlibro.Cells(i, 2) = valueStr

        Next
        listFolders.Clear()
        listPercentage.Clear()


        'readFileDF("DF_ESC2.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 13
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    valueStr = listPercentage.Item(indexReturn)
        '    xlibro.Cells(i, 3) = valueStr

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()

        readFileDF("DF_ESC3.txt", listFolders, listPercentage, errMessageStr)
        For i = 6 To 13
            pathToFindStr = xlibro.Cells(i, 1).Value
            indexReturn = listFolders.IndexOf(pathToFindStr)
            valueStr = listPercentage.Item(indexReturn)
            xlibro.Cells(i, 4) = valueStr

        Next
        listFolders.Clear()
        listPercentage.Clear()

        readFileDF("DF_ESC4.txt", listFolders, listPercentage, errMessageStr)
        For i = 6 To 13
            pathToFindStr = xlibro.Cells(i, 1).Value
            indexReturn = listFolders.IndexOf(pathToFindStr)
            valueStr = listPercentage.Item(indexReturn)
            xlibro.Cells(i, 5) = valueStr

        Next
        listFolders.Clear()
        listPercentage.Clear()

        'readFileDF("DF_ESC5.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 13
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    valueStr = listPercentage.Item(indexReturn)
        '    xlibro.Cells(i, 6) = valueStr

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()


        'readFileDF("DF_ESC6.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 13
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    valueStr = listPercentage.Item(indexReturn)
        '    xlibro.Cells(i, 7) = valueStr

        'Next
        listFolders.Clear()
        listPercentage.Clear()

        '-=======================================================================================

        xlibro.Sheets("PortalsPDC").Select()
        xlibro.Cells(3, 1) = fecha
        xlibro.Cells(3, 2) = timeStr

        readFileDF("DF_B2BAPP1.txt", listFolders, listPercentage, errMessageStr)
        For i = 6 To 13
            pathToFindStr = xlibro.Cells(i, 1).Value
            indexReturn = listFolders.IndexOf(pathToFindStr)
            valueStr = listPercentage.Item(indexReturn)
            xlibro.Cells(i, 2) = valueStr

        Next
        listFolders.Clear()
        listPercentage.Clear()


        readFileDF("DF_B2BAPP2.txt", listFolders, listPercentage, errMessageStr)
        For i = 6 To 13
            pathToFindStr = xlibro.Cells(i, 1).Value
            indexReturn = listFolders.IndexOf(pathToFindStr)
            valueStr = listPercentage.Item(indexReturn)
            xlibro.Cells(i, 3) = valueStr

        Next
        listFolders.Clear()
        listPercentage.Clear()

        readFileDF("DF_B2BWEB1.txt", listFolders, listPercentage, errMessageStr)
        For i = 6 To 13
            pathToFindStr = xlibro.Cells(i, 1).Value
            indexReturn = listFolders.IndexOf(pathToFindStr)
            valueStr = listPercentage.Item(indexReturn)
            xlibro.Cells(i, 4) = valueStr

        Next
        listFolders.Clear()
        listPercentage.Clear()

        readFileDF("DF_B2BWEB2.txt", listFolders, listPercentage, errMessageStr)
        For i = 6 To 13
            pathToFindStr = xlibro.Cells(i, 1).Value
            indexReturn = listFolders.IndexOf(pathToFindStr)
            valueStr = listPercentage.Item(indexReturn)
            xlibro.Cells(i, 5) = valueStr

        Next
        listFolders.Clear()
        listPercentage.Clear()

        readFileDF("DF_PWA1.txt", listFolders, listPercentage, errMessageStr)
        For i = 6 To 15
            pathToFindStr = xlibro.Cells(i, 1).Value
            indexReturn = listFolders.IndexOf(pathToFindStr)
            valueStr = listPercentage.Item(indexReturn)
            xlibro.Cells(i, 6) = valueStr

        Next
        listFolders.Clear()
        listPercentage.Clear()


        readFileDF("DF_PWA2.txt", listFolders, listPercentage, errMessageStr)
        For i = 6 To 15
            pathToFindStr = xlibro.Cells(i, 1).Value
            indexReturn = listFolders.IndexOf(pathToFindStr)
            valueStr = listPercentage.Item(indexReturn)
            xlibro.Cells(i, 7) = valueStr

        Next
        listFolders.Clear()
        listPercentage.Clear()


        readFileDF("DF_PPAPP1.txt", listFolders, listPercentage, errMessageStr)
        For i = 6 To 13
            pathToFindStr = xlibro.Cells(i, 1).Value
            indexReturn = listFolders.IndexOf(pathToFindStr)
            valueStr = listPercentage.Item(indexReturn)
            xlibro.Cells(i, 8) = valueStr

        Next
        listFolders.Clear()
        listPercentage.Clear()

        readFileDF("DF_PPAPP2.txt", listFolders, listPercentage, errMessageStr)
        For i = 6 To 13
            pathToFindStr = xlibro.Cells(i, 1).Value
            indexReturn = listFolders.IndexOf(pathToFindStr)
            valueStr = listPercentage.Item(indexReturn)
            xlibro.Cells(i, 9) = valueStr

        Next
        listFolders.Clear()
        listPercentage.Clear()

        readFileDF("DF_PPWEB1.txt", listFolders, listPercentage, errMessageStr)
        For i = 6 To 13
            pathToFindStr = xlibro.Cells(i, 1).Value
            indexReturn = listFolders.IndexOf(pathToFindStr)
            valueStr = listPercentage.Item(indexReturn)
            xlibro.Cells(i, 10) = valueStr

        Next
        listFolders.Clear()
        listPercentage.Clear()

        readFileDF("DF_PPWEB2.txt", listFolders, listPercentage, errMessageStr)
        For i = 6 To 13
            pathToFindStr = xlibro.Cells(i, 1).Value
            indexReturn = listFolders.IndexOf(pathToFindStr)
            valueStr = listPercentage.Item(indexReturn)
            xlibro.Cells(i, 11) = valueStr

        Next
        listFolders.Clear()
        listPercentage.Clear()

        '====================================================================================================================

        xlibro.Sheets("Database").Select()
        xlibro.Cells(3, 1) = fecha
        xlibro.Cells(3, 2) = timeStr

        readFileDF("DF_DB1.txt", listFolders, listPercentage, errMessageStr)
        For i = 6 To 23
            pathToFindStr = xlibro.Cells(i, 1).Value
            indexReturn = listFolders.IndexOf(pathToFindStr)
            If indexReturn >= 0 Then
                valueStr = listPercentage.Item(indexReturn)
                xlibro.Cells(i, 2) = valueStr
            End If
        Next
        listFolders.Clear()
        listPercentage.Clear()


        readFileDF("DF_DB2.txt", listFolders, listPercentage, errMessageStr)
        For i = 6 To 23
            pathToFindStr = xlibro.Cells(i, 1).Value
            indexReturn = listFolders.IndexOf(pathToFindStr)

            If indexReturn >= 0 Then
                valueStr = listPercentage.Item(indexReturn)
                xlibro.Cells(i, 3) = valueStr
            End If
        Next
        listFolders.Clear()
        listPercentage.Clear()


        readFileDF("DF_DB3.txt", listFolders, listPercentage, errMessageStr)
        For i = 6 To 23
            pathToFindStr = xlibro.Cells(i, 1).Value
            indexReturn = listFolders.IndexOf(pathToFindStr)
            If indexReturn >= 0 Then
                valueStr = listPercentage.Item(indexReturn)
                xlibro.Cells(i, 4) = valueStr
            End If

        Next
        listFolders.Clear()
        listPercentage.Clear()


        readFileDF("DF_DB4.txt", listFolders, listPercentage, errMessageStr)
        For i = 6 To 23
            pathToFindStr = xlibro.Cells(i, 1).Value
            indexReturn = listFolders.IndexOf(pathToFindStr)
            If indexReturn >= 0 Then
                valueStr = listPercentage.Item(indexReturn)
                xlibro.Cells(i, 5) = valueStr
            End If

        Next
        listFolders.Clear()
        listPercentage.Clear()

        'readFileDF("DF_PPDB1.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 23
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    If indexReturn >= 0 Then
        '        valueStr = listPercentage.Item(indexReturn)
        '        xlibro.Cells(i, 6) = valueStr
        '    End If

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()


        'readFileDF("DF_PPDB2.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 23
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    If indexReturn >= 0 Then
        '        valueStr = listPercentage.Item(indexReturn)
        '        xlibro.Cells(i, 7) = valueStr
        '    End If

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()


        'readFileDF("DF_PPDB3.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 23
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    If indexReturn >= 0 Then
        '        valueStr = listPercentage.Item(indexReturn)
        '        xlibro.Cells(i, 8) = valueStr
        '    End If

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()

        'readFileDF("DF_PPDB4.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 23
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    If indexReturn >= 0 Then
        '        valueStr = listPercentage.Item(indexReturn)
        '        xlibro.Cells(i, 9) = valueStr
        '    End If
        'Next
        'listFolders.Clear()
        'listPercentage.Clear()

        '=================================================================================================

        xlibro.Sheets("PortalsBDC").Select()
        xlibro.Cells(3, 1) = fecha
        xlibro.Cells(3, 2) = timeStr

        'readFileDF("DF_B2BAPP3.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 13
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    valueStr = listPercentage.Item(indexReturn)
        '    xlibro.Cells(i, 2) = valueStr

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()


        'readFileDF("DF_B2BAPP4.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 13
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    valueStr = listPercentage.Item(indexReturn)
        '    xlibro.Cells(i, 3) = valueStr

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()

        'readFileDF("DF_B2BWEB3.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 13
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    valueStr = listPercentage.Item(indexReturn)
        '    xlibro.Cells(i, 4) = valueStr

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()

        'readFileDF("DF_B2BWEB4.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 13
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    valueStr = listPercentage.Item(indexReturn)
        '    xlibro.Cells(i, 5) = valueStr

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()

        readFileDF("DF_PWA3.txt", listFolders, listPercentage, errMessageStr)
        For i = 6 To 15
            pathToFindStr = xlibro.Cells(i, 1).Value
            indexReturn = listFolders.IndexOf(pathToFindStr)
            valueStr = listPercentage.Item(indexReturn)
            xlibro.Cells(i, 6) = valueStr

        Next
        listFolders.Clear()
        listPercentage.Clear()


        readFileDF("DF_PWA4.txt", listFolders, listPercentage, errMessageStr)
        For i = 6 To 15
            pathToFindStr = xlibro.Cells(i, 1).Value
            indexReturn = listFolders.IndexOf(pathToFindStr)
            valueStr = listPercentage.Item(indexReturn)
            xlibro.Cells(i, 7) = valueStr

        Next
        listFolders.Clear()
        listPercentage.Clear()


        readFileDF("DF_PPAPP3.txt", listFolders, listPercentage, errMessageStr)
        For i = 6 To 13
            pathToFindStr = xlibro.Cells(i, 1).Value
            indexReturn = listFolders.IndexOf(pathToFindStr)
            valueStr = listPercentage.Item(indexReturn)
            xlibro.Cells(i, 8) = valueStr

        Next
        listFolders.Clear()
        listPercentage.Clear()

        'readFileDF("DF_PPAPP4.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 13
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    valueStr = listPercentage.Item(indexReturn)
        '    xlibro.Cells(i, 9) = valueStr

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()

        'readFileDF("DF_PPWEB3.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 13
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    valueStr = listPercentage.Item(indexReturn)
        '    xlibro.Cells(i, 10) = valueStr

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()

        'readFileDF("DF_PPWEB4.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 13
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    valueStr = listPercentage.Item(indexReturn)
        '    xlibro.Cells(i, 11) = valueStr

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()


        xlibro.Sheets("SFTP - DMZSFTP").Select()
        xlibro.Cells(3, 1) = fecha
        xlibro.Cells(3, 2) = timeStr

        readFileDF("NYSFTP1.txt", listFolders, listPercentage, errMessageStr)
        For i = 6 To 13
            pathToFindStr = xlibro.Cells(i, 1).Value
            indexReturn = listFolders.IndexOf(pathToFindStr)
            valueStr = listPercentage.Item(indexReturn)
            xlibro.Cells(i, 2) = valueStr

        Next
        listFolders.Clear()
        listPercentage.Clear()


        readFileDF("NYSFTP2.txt", listFolders, listPercentage, errMessageStr)
        For i = 6 To 13
            pathToFindStr = xlibro.Cells(i, 1).Value
            indexReturn = listFolders.IndexOf(pathToFindStr)
            valueStr = listPercentage.Item(indexReturn)
            xlibro.Cells(i, 3) = valueStr

        Next
        listFolders.Clear()
        listPercentage.Clear()


        readFileDF("NYDMZSFTP1.txt", listFolders, listPercentage, errMessageStr)
        For i = 6 To 13
            pathToFindStr = xlibro.Cells(i, 1).Value
            indexReturn = listFolders.IndexOf(pathToFindStr)
            valueStr = listPercentage.Item(indexReturn)
            xlibro.Cells(i, 4) = valueStr

        Next
        listFolders.Clear()
        listPercentage.Clear()



        xlibro.Sheets("PortalsPDC").Select()


        fecha = Format(Now, "MM/dd/yyyy HH:mm:ss")
        monthStr = CStr(fecha.Month).PadLeft(2, "0")
        dayStr = CStr(fecha.Day).PadLeft(2, "0")
        yearStr = CStr(fecha.Year).PadLeft(4, "0")
        hourStr = CStr(fecha.Hour).PadLeft(2, "0")
        minStr = CStr(fecha.Minute).PadLeft(2, "0")
        secStr = CStr(fecha.Second).PadLeft(2, "0")
        fileNameStr = "DiskSpaceCheck_" & monthStr & dayStr & yearStr & "_" & hourStr & minStr & secStr & ".xlsx"
        strPathRemoteFile = "files/DF_Files"


        xlibro.ActiveWorkbook.SaveAs(strPathFile & "\" & fileNameStr)

        xlibro.Workbooks.Close()

        xlibro.Quit()

        Me.Cursor = Cursors.Default
        Me.RichTextBox1.Text = "All Done."
        Me.Refresh()


        '===========================================================================================================================================================



        Me.RichTextBox1.Text = "Uploading to SFTP..." & vbCrLf
        Me.Refresh()

        errCodeInt = upLoadFileSFTP("10.1.5.182", "22", "xfer", "Welcome1", strPathRemoteFile & "/" & fileNameStr, strPathFile & "\" & fileNameStr, errMessageStr)


        Select Case errCodeInt

            Case 1
                MsgBox("SFTP Component error. " + errMessageStr, MsgBoxStyle.OkOnly, "Error.")
                Exit Sub
            Case 2

                MsgBox("Conection error. " + errMessageStr, MsgBoxStyle.OkOnly, "Error.")
                Exit Sub

            Case 3
                MsgBox("Authenticate error. " + errMessageStr, MsgBoxStyle.OkOnly, "Error.")
                Exit Sub

            Case 4
                MsgBox("SFTP Initialize error. " + errMessageStr, MsgBoxStyle.OkOnly, "Error.")
                Exit Sub


            Case 5
                MsgBox("Upload file error. " + errMessageStr, MsgBoxStyle.OkOnly, "Error.")
                Exit Sub

            Case 0

                'file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "LogTransferWinnnerAndJackpotFiles_" & monthStr & dayStr & yearStr & ".txt", True)
                'file.WriteLine(Now & "  winner_summary_report_p010_d" & drawNumberStr & "_wincnt_english.rep" & " was successfully transferred")
                'file.Close()

                'Return 0
                Me.Cursor = Cursors.Default
                Me.RichTextBox1.Text = "All Done."
                Me.Refresh()


        End Select


    End Sub


    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        Dim errCodeInt As Integer = 0
        Dim errMessageStr As String = ""
        Static count As Integer = 0
        count = count + 10
        If count <= 100 Then
            ProgressBar1.Value = count
        Else
            Timer1.Enabled = False

            stopBtn.Visible = False
            ProgressBar1.Visible = False

            okBtn.Visible = False
            cancelBtn.Visible = False


            getDFFromAllServersSSH()
            fillExcelAllServers()

            Me.Dispose()


        End If


        'Me.RichTextBox1.Text = "Getting DF ESTE1..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.1.5.10", "prosys", "Numb3r1j0b", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        '    Me.Close()
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_ESTE1.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()

        'Me.Refresh()

        'Me.RichTextBox1.Text = "Getting DF ESTE2..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.1.5.11", "prosys", "Numb3r1j0b", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        '    Me.Close()
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_ESTE2.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()

        'Me.Refresh()

        'Me.RichTextBox1.Text = "Getting DF ESTE3..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.1.5.12", "prosys", "Numb3r1j0b", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        '    Me.Close()
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_ESTE3.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()

        'Me.Refresh()

        'Me.RichTextBox1.Text = "Getting DF ESTE4..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.2.5.13", "prosys", "Numb3r1j0b", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        '    Me.Close()
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_ESTE4.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()

        'Me.Refresh()

        'Me.RichTextBox1.Text = "Getting DF ESTE5..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.2.5.14", "prosys", "Numb3r1j0b", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        '    Me.Close()
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_ESTE5.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()

        'Me.Refresh()

        'Me.RichTextBox1.Text = "Getting DF ESC1..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.1.5.61", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        '    Me.Close()
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_ESC1.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()

        'Me.RichTextBox1.Text = "Getting DF ESC2..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.1.5.62", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        '    Me.Close()
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_ESC2.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()

        'Me.RichTextBox1.Text = "Getting DF ESC3..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.1.5.63", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        '    Me.Close()
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_ESC3.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()

        'Me.RichTextBox1.Text = "Getting DF ESC4..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.2.5.64", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        '    Me.Close()
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_ESC4.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()


        'Me.RichTextBox1.Text = "Getting DF ESC5..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.2.5.65", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        '    Me.Close()
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_ESC5.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()


        'Me.RichTextBox1.Text = "Getting DF ESC6..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.2.5.65", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        '    Me.Close()
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_ESC6.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()



        'Me.RichTextBox1.Text = "Getting DF B2BAPP1..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.1.5.103", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        '    Me.Close()
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_B2BAPP1.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()


        'Me.RichTextBox1.Text = "Getting DF B2BAPP2..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.1.5.104", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        '    Me.Close()
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_B2BAPP2.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()


        'Me.RichTextBox1.Text = "Getting DF B2BWEB1..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.1.4.121", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        '    Me.Close()
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_B2BWEB1.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()


        'Me.RichTextBox1.Text = "Getting DF B2BWEB2..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.1.4.122", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        '    Me.Close()
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_B2BWEB2.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()



        'Me.RichTextBox1.Text = "Getting DF PWA1..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.1.5.77", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        '    Me.Close()
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_PWA1.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()


        'Me.RichTextBox1.Text = "Getting DF PWA2..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.1.5.78", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        '    Me.Close()
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_PWA2.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()



        'Me.RichTextBox1.Text = "Getting DF PPAPP1..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.1.5.166", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        '    Me.Close()
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_PPAPP1.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()


        'Me.RichTextBox1.Text = "Getting DF PPAPP2..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.1.5.167", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        '    Me.Close()
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_PPAPP2.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()

        'Me.Refresh()

        'Me.RichTextBox1.Text = "Getting DF PPWEB1..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.1.218.170", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        '    Me.Close()
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_PPWEB1.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()


        'Me.RichTextBox1.Text = "Getting DF PPWEB2..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.1.218.171", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        '    Me.Close()
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_PPWEB2.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()

        ''===================================================================================

        'Me.RichTextBox1.Text = "Getting DF NYDB1..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.1.5.91", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_DB1.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()

        'Me.RichTextBox1.Text = "Getting DF NYDB2..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.1.5.92", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_DB2.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()

        'Me.RichTextBox1.Text = "Getting DF NYDB3..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.2.5.93", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_DB3.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()

        'Me.Refresh()

        'Me.RichTextBox1.Text = "Getting DF NYDB4..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.2.5.94", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_DB4.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()


        'Me.RichTextBox1.Text = "Getting DF NYPPDB1..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.1.5.59", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_PPDB1.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()


        'Me.RichTextBox1.Text = "Getting DF NYPPDB2..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.1.5.60", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_PPDB2.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()


        'Me.Refresh()

        'Me.RichTextBox1.Text = "Getting DF NYPPDB3..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.2.5.89", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_PPDB3.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()


        'Me.RichTextBox1.Text = "Getting DF NYPPDB4..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.2.5.90", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_PPDB4.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()

        ''======================================================================================

        'Me.RichTextBox1.Text = "Getting DF NYB2BAPP3..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.2.5.106", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_B2BAPP3.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()

        'Me.RichTextBox1.Text = "Getting DF NYB2BAPP4..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.2.5.107", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_B2BAPP4.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()


        'Me.RichTextBox1.Text = "Getting DF NYPWA3..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.2.5.79", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_PWA3.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()

        'Me.RichTextBox1.Text = "Getting DF NYPWA4..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.2.5.80", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_PWA4.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()



        'Me.RichTextBox1.Text = "Getting DF NYPPAPP3..." & vbCrLf
        'Me.Refresh()
        'infoStr = getDFFromSSH("10.2.5.168", "ricie", "@m3t5", errCodeInt, errMessageStr)
        'If errCodeInt <> 0 Then
        '    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        'End If

        'file = My.Computer.FileSystem.OpenTextFileWriter("DF_PPAPP3.txt", False)
        'file.WriteLine(infoStr)
        'file.Close()

        ''Me.RichTextBox1.Text = "Getting DF NYPPAPP4..." & vbCrLf
        ''Me.Refresh()
        ''infoStr = getDFFromSSH("10.2.5.169", "ricie", "@m3t5", errCodeInt, errMessageStr)
        ''If errCodeInt <> 0 Then
        ''    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        ''End If

        ''file = My.Computer.FileSystem.OpenTextFileWriter("DF_PPAPP4.txt", False)
        ''file.WriteLine(infoStr)
        ''file.Close()

        'Me.Refresh()


        '=========================================================================================================================================


        'fecha = Format(Now, "MM/dd/yyyy")
        'timeStr = Format(Now, "HH:mm:ss")


        'Me.RichTextBox1.Text = "Filling excel form..." & vbCrLf
        'Me.Refresh()


        'strPathFile = System.Windows.Forms.Application.StartupPath
        'fileNameStr = strPathFile & "\Disk Space Check.xlsx"

        'xlibro = CreateObject("Excel.Application")
        'xlibro.Workbooks.Open(fileNameStr)
        'xlibro.Visible = False
        'xlibro.Sheets("ESTE").Select()
        'xlibro.Cells(3, 1) = fecha
        'xlibro.Cells(3, 2) = timeStr

        'readFileDF("DF_ESTE1.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 35
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    valueStr = listPercentage.Item(indexReturn)
        '    xlibro.Cells(i, 2) = valueStr

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()

        'readFileDF("DF_ESTE2.txt", listFolders, listPercentage, errMessageStr)

        'For i = 6 To 35
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    valueStr = listPercentage.Item(indexReturn)
        '    xlibro.Cells(i, 3) = valueStr

        'Next

        'listFolders.Clear()
        'listPercentage.Clear()

        'readFileDF("DF_ESTE3.txt", listFolders, listPercentage, errMessageStr)

        'For i = 6 To 35
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    valueStr = listPercentage.Item(indexReturn)
        '    xlibro.Cells(i, 4) = valueStr

        'Next

        'listFolders.Clear()
        'listPercentage.Clear()

        'readFileDF("DF_ESTE4.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 35
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    valueStr = listPercentage.Item(indexReturn)
        '    xlibro.Cells(i, 5) = valueStr

        'Next

        'listFolders.Clear()
        'listPercentage.Clear()

        'readFileDF("DF_ESTE5.txt", listFolders, listPercentage, errMessageStr)

        'For i = 6 To 35
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    valueStr = listPercentage.Item(indexReturn)
        '    xlibro.Cells(i, 6) = valueStr
        'Next

        'listFolders.Clear()
        'listPercentage.Clear()

        ''============================================================================================

        'xlibro.Sheets("ESC").Select()
        'xlibro.Cells(3, 1) = fecha
        'xlibro.Cells(3, 2) = timeStr

        'readFileDF("DF_ESC1.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 13
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    valueStr = listPercentage.Item(indexReturn)
        '    xlibro.Cells(i, 2) = valueStr

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()


        'readFileDF("DF_ESC2.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 13
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    valueStr = listPercentage.Item(indexReturn)
        '    xlibro.Cells(i, 3) = valueStr

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()

        'readFileDF("DF_ESC3.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 13
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    valueStr = listPercentage.Item(indexReturn)
        '    xlibro.Cells(i, 4) = valueStr

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()

        'readFileDF("DF_ESC4.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 13
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    valueStr = listPercentage.Item(indexReturn)
        '    xlibro.Cells(i, 5) = valueStr

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()

        'readFileDF("DF_ESC5.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 13
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    valueStr = listPercentage.Item(indexReturn)
        '    xlibro.Cells(i, 6) = valueStr

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()


        'readFileDF("DF_ESC6.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 13
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    valueStr = listPercentage.Item(indexReturn)
        '    xlibro.Cells(i, 7) = valueStr

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()

        ''-=======================================================================================

        'xlibro.Sheets("PortalsPDC").Select()
        'xlibro.Cells(3, 1) = fecha
        'xlibro.Cells(3, 2) = timeStr

        'readFileDF("DF_B2BAPP1.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 13
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    valueStr = listPercentage.Item(indexReturn)
        '    xlibro.Cells(i, 2) = valueStr

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()


        'readFileDF("DF_B2BAPP2.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 13
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    valueStr = listPercentage.Item(indexReturn)
        '    xlibro.Cells(i, 3) = valueStr

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()

        'readFileDF("DF_B2BWEB1.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 13
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    valueStr = listPercentage.Item(indexReturn)
        '    xlibro.Cells(i, 4) = valueStr

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()

        'readFileDF("DF_B2BWEB2.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 13
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    valueStr = listPercentage.Item(indexReturn)
        '    xlibro.Cells(i, 5) = valueStr

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()

        'readFileDF("DF_PWA1.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 15
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    valueStr = listPercentage.Item(indexReturn)
        '    xlibro.Cells(i, 6) = valueStr

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()


        'readFileDF("DF_PWA2.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 15
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    valueStr = listPercentage.Item(indexReturn)
        '    xlibro.Cells(i, 7) = valueStr

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()


        'readFileDF("DF_PPAPP1.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 13
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    valueStr = listPercentage.Item(indexReturn)
        '    xlibro.Cells(i, 8) = valueStr

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()

        'readFileDF("DF_PPAPP2.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 13
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    valueStr = listPercentage.Item(indexReturn)
        '    xlibro.Cells(i, 9) = valueStr

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()

        'readFileDF("DF_PPWEB1.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 13
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    valueStr = listPercentage.Item(indexReturn)
        '    xlibro.Cells(i, 10) = valueStr

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()

        'readFileDF("DF_PPWEB2.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 13
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    valueStr = listPercentage.Item(indexReturn)
        '    xlibro.Cells(i, 11) = valueStr

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()

        ''====================================================================================================================

        'xlibro.Sheets("Database").Select()
        'xlibro.Cells(3, 1) = fecha
        'xlibro.Cells(3, 2) = timeStr

        'readFileDF("DF_DB1.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 23
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    If indexReturn >= 0 Then
        '        valueStr = listPercentage.Item(indexReturn)
        '        xlibro.Cells(i, 2) = valueStr
        '    End If
        'Next
        'listFolders.Clear()
        'listPercentage.Clear()


        'readFileDF("DF_DB2.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 23
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)

        '    If indexReturn >= 0 Then
        '        valueStr = listPercentage.Item(indexReturn)
        '        xlibro.Cells(i, 3) = valueStr
        '    End If
        'Next
        'listFolders.Clear()
        'listPercentage.Clear()


        'readFileDF("DF_DB3.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 23
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    If indexReturn >= 0 Then
        '        valueStr = listPercentage.Item(indexReturn)
        '        xlibro.Cells(i, 4) = valueStr
        '    End If

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()


        'readFileDF("DF_DB4.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 23
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    If indexReturn >= 0 Then
        '        valueStr = listPercentage.Item(indexReturn)
        '        xlibro.Cells(i, 5) = valueStr
        '    End If

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()

        'readFileDF("DF_PPDB1.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 23
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    If indexReturn >= 0 Then
        '        valueStr = listPercentage.Item(indexReturn)
        '        xlibro.Cells(i, 6) = valueStr
        '    End If

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()


        'readFileDF("DF_PPDB2.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 23
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    If indexReturn >= 0 Then
        '        valueStr = listPercentage.Item(indexReturn)
        '        xlibro.Cells(i, 7) = valueStr
        '    End If

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()


        'readFileDF("DF_PPDB3.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 23
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    If indexReturn >= 0 Then
        '        valueStr = listPercentage.Item(indexReturn)
        '        xlibro.Cells(i, 8) = valueStr
        '    End If

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()

        'readFileDF("DF_PPDB4.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 23
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    If indexReturn >= 0 Then
        '        valueStr = listPercentage.Item(indexReturn)
        '        xlibro.Cells(i, 9) = valueStr
        '    End If
        'Next
        'listFolders.Clear()
        'listPercentage.Clear()

        ''=================================================================================================

        'xlibro.Sheets("PortalsBDC").Select()
        'xlibro.Cells(3, 1) = fecha
        'xlibro.Cells(3, 2) = timeStr

        'readFileDF("DF_B2BAPP3.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 13
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    valueStr = listPercentage.Item(indexReturn)
        '    xlibro.Cells(i, 2) = valueStr

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()


        'readFileDF("DF_B2BAPP4.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 13
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    valueStr = listPercentage.Item(indexReturn)
        '    xlibro.Cells(i, 3) = valueStr

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()

        ''readFileDF("DF_B2BWEB3.txt", listFolders, listPercentage, errMessageStr)
        ''For i = 6 To 13
        ''    pathToFindStr = xlibro.Cells(i, 1).Value
        ''    indexReturn = listFolders.IndexOf(pathToFindStr)
        ''    valueStr = listPercentage.Item(indexReturn)
        ''    xlibro.Cells(i, 4) = valueStr

        ''Next
        ''listFolders.Clear()
        ''listPercentage.Clear()

        ''readFileDF("DF_B2BWEB4.txt", listFolders, listPercentage, errMessageStr)
        ''For i = 6 To 13
        ''    pathToFindStr = xlibro.Cells(i, 1).Value
        ''    indexReturn = listFolders.IndexOf(pathToFindStr)
        ''    valueStr = listPercentage.Item(indexReturn)
        ''    xlibro.Cells(i, 5) = valueStr

        ''Next
        ''listFolders.Clear()
        ''listPercentage.Clear()

        'readFileDF("DF_PWA3.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 15
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    valueStr = listPercentage.Item(indexReturn)
        '    xlibro.Cells(i, 6) = valueStr

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()


        'readFileDF("DF_PWA4.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 15
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    valueStr = listPercentage.Item(indexReturn)
        '    xlibro.Cells(i, 7) = valueStr

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()


        'readFileDF("DF_PPAPP3.txt", listFolders, listPercentage, errMessageStr)
        'For i = 6 To 13
        '    pathToFindStr = xlibro.Cells(i, 1).Value
        '    indexReturn = listFolders.IndexOf(pathToFindStr)
        '    valueStr = listPercentage.Item(indexReturn)
        '    xlibro.Cells(i, 8) = valueStr

        'Next
        'listFolders.Clear()
        'listPercentage.Clear()

        ''readFileDF("DF_PPAPP4.txt", listFolders, listPercentage, errMessageStr)
        ''For i = 6 To 13
        ''    pathToFindStr = xlibro.Cells(i, 1).Value
        ''    indexReturn = listFolders.IndexOf(pathToFindStr)
        ''    valueStr = listPercentage.Item(indexReturn)
        ''    xlibro.Cells(i, 9) = valueStr

        ''Next
        ''listFolders.Clear()
        ''listPercentage.Clear()

        ''readFileDF("DF_PPWEB3.txt", listFolders, listPercentage, errMessageStr)
        ''For i = 6 To 13
        ''    pathToFindStr = xlibro.Cells(i, 1).Value
        ''    indexReturn = listFolders.IndexOf(pathToFindStr)
        ''    valueStr = listPercentage.Item(indexReturn)
        ''    xlibro.Cells(i, 10) = valueStr

        ''Next
        ''listFolders.Clear()
        ''listPercentage.Clear()

        ''readFileDF("DF_PPWEB4.txt", listFolders, listPercentage, errMessageStr)
        ''For i = 6 To 13
        ''    pathToFindStr = xlibro.Cells(i, 1).Value
        ''    indexReturn = listFolders.IndexOf(pathToFindStr)
        ''    valueStr = listPercentage.Item(indexReturn)
        ''    xlibro.Cells(i, 11) = valueStr

        ''Next
        ''listFolders.Clear()
        ''listPercentage.Clear()
        'xlibro.Sheets("PortalsPDC").Select()


        'fecha = Format(Now, "MM/dd/yyyy HH:mm:ss")
        'monthStr = CStr(fecha.Month).PadLeft(2, "0")
        'dayStr = CStr(fecha.Day).PadLeft(2, "0")
        'yearStr = CStr(fecha.Year).PadLeft(4, "0")
        'hourStr = CStr(fecha.Hour).PadLeft(2, "0")
        'minStr = CStr(fecha.Minute).PadLeft(2, "0")
        'secStr = CStr(fecha.Second).PadLeft(2, "0")
        'fileNameStr = "DiskSpaceCheck_" & monthStr & dayStr & yearStr & "_" & hourStr & minStr & secStr & ".xlsx"
        'strPathRemoteFile = "files/DF_Files"


        'xlibro.ActiveWorkbook.SaveAs(strPathFile & "\" & fileNameStr)

        'xlibro.Workbooks.Close()

        'xlibro.Quit()

        'Me.Cursor = Cursors.Default
        'Me.RichTextBox1.Text = "All Done."
        'Me.Refresh()




        '===========================================================================================================================================================



        'Me.RichTextBox1.Text = "Uploading to SFTP..." & vbCrLf
        'Me.Refresh()

        'errCodeInt = upLoadFileSFTP("10.1.5.182", "22", "xfer", "Welcome1", strPathRemoteFile & "/" & fileNameStr, strPathFile & "\" & fileNameStr, errMessageStr)


        'Select Case errCodeInt

        '    Case 1
        '        MsgBox("SFTP Component error. " + errMessageStr, MsgBoxStyle.OkOnly, "Error.")
        '        Exit Sub
        '    Case 2

        '        MsgBox("Conection error. " + errMessageStr, MsgBoxStyle.OkOnly, "Error.")
        '        Exit Sub

        '    Case 3
        '        MsgBox("Authenticate error. " + errMessageStr, MsgBoxStyle.OkOnly, "Error.")
        '        Exit Sub

        '    Case 4
        '        MsgBox("SFTP Initialize error. " + errMessageStr, MsgBoxStyle.OkOnly, "Error.")
        '        Exit Sub


        '    Case 5
        '        MsgBox("Upload file error. " + errMessageStr, MsgBoxStyle.OkOnly, "Error.")
        '        Exit Sub

        '    Case 0

        '        'file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "LogTransferWinnnerAndJackpotFiles_" & monthStr & dayStr & yearStr & ".txt", True)
        '        'file.WriteLine(Now & "  winner_summary_report_p010_d" & drawNumberStr & "_wincnt_english.rep" & " was successfully transferred")
        '        'file.Close()

        '        'Return 0
        '        Me.Cursor = Cursors.Default
        '        Me.RichTextBox1.Text = "All Done."
        '        Me.Refresh()


        'End Select


           



    End Sub

    Private Sub stopBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles stopBtn.Click
        Timer1.Enabled = False
        stopBtn.Visible = False
        ProgressBar1.Visible = False
        
        okBtn.Visible = True
        cancelBtn.Visible = True
    End Sub

    Function sent_Email(ByVal smtpHostStr As String, ByVal toStr As String, ByVal attachStr As String, ByVal subjectStr As String, ByVal bodyStr As String) As Integer
        Dim SMTPserver As New SmtpClient
        Dim mail As New MailMessage
        Dim oAttch As Attachment = New Attachment(attachStr)


        Try
            SMTPserver.Host = smtpHostStr
            mail = New MailMessage
            mail.From = New MailAddress("do.not.reply@gtech-noreply.com")
            mail.To.Add(toStr) 'The Man you want to send the message to him
            mail.Subject = subjectStr
            mail.Body = bodyStr
            mail.Attachments.Add(oAttch)
            SMTPserver.Send(mail)
            Return 0
            'MessageBox.Show("Done!", "Message Sent", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)


        Catch ex As Exception
            Return -1

        End Try




    End Function
End Class
