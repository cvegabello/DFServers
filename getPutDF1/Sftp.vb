
Module Sftp

    Function downloadFileSFTP(ByVal hostName As String, ByVal port As Integer, ByVal username As String, ByVal password As String, ByVal remoteFile As String, ByVal localFile As String, ByRef strError As String) As Integer

        Dim sftp As New Chilkat.SFtp()
        Dim success As Boolean = sftp.UnlockComponent("Anything for 30-day trial")

        Dim handle As String

        If (success <> True) Then
            strError = sftp.LastErrorText
            Return 1
        End If


        '  Set some timeouts, in milliseconds:
        sftp.ConnectTimeoutMs = 5000
        sftp.IdleTimeoutMs = 10000

        '  Connect to the SSH server.
        '  The standard SSH port = 22
        '  The hostname may be a hostname or IP address.

        success = sftp.Connect(hostName, port)
        If (success <> True) Then
            strError = sftp.LastErrorText
            Return 2
        Else
            'MsgBox("EXITO Coneccion", MsgBoxStyle.OkOnly, "Funciono Conneccion")
            'Console.WriteLine("Exito")
        End If

        success = sftp.AuthenticatePw(username, password)
        If (success <> True) Then

            strError = sftp.LastErrorText
            Return 3
        Else
            'MsgBox("Funciono Autenticacion", MsgBoxStyle.OkOnly, "Funciono Autenticacion. ")
        End If


        '  After authenticating, the SFTP subsystem must be initialized:
        success = sftp.InitializeSftp()
        If (success <> True) Then
            strError = sftp.LastErrorText
            Return 4
        Else

            'MsgBox("Funciono Inicializacion", MsgBoxStyle.OkOnly, "Funciono Inicializacion. ")

        End If


        handle = sftp.OpenFile(remoteFile, "readOnly", "openExisting")
        If (handle = vbNullString) Then
            strError = sftp.LastErrorText
            Return 5
        Else
            'MsgBox("Funciono abrir el archivo", MsgBoxStyle.OkOnly, "Funciono abrir el archivo. ")

        End If


        '  Download the file:
        success = sftp.DownloadFile(handle, localFile)
        If (success <> True) Then
            strError = sftp.LastErrorText
            Return 6
        Else
            'MsgBox("Funciono bajar el archivo", MsgBoxStyle.OkOnly, "Funciono bajar el archivo. ")

        End If

        '  Close the file.
        success = sftp.CloseHandle(handle)
        If (success <> True) Then
            strError = sftp.LastErrorText
            Return 7
        Else
            'MsgBox("Funciono cerrar el archivo", MsgBoxStyle.OkOnly, "Funciono cerrar el archivo. ")
        End If
        Return 0

        'MsgBox("Funciono todo", MsgBoxStyle.OkOnly, "EXITO TOTAL")


    End Function

    Function upLoadFileSFTP(ByVal hostName As String, ByVal port As Integer, ByVal username As String, ByVal password As String, ByVal remoteFile As String, ByVal localFile As String, ByRef strError As String) As Integer

        Dim sftp As New Chilkat.SFtp()
        Dim success As Boolean = sftp.UnlockComponent("Anything for 30-day trial")

        If (success <> True) Then
            strError = sftp.LastErrorText
            Return 1
        End If

        '  Set some timeouts, in milliseconds:
        sftp.ConnectTimeoutMs = 5000
        sftp.IdleTimeoutMs = 10000

        '  Connect to the SSH server.
        '  The standard SSH port = 22
        '  The hostname may be a hostname or IP address.
        
        success = sftp.Connect(hostName, port)
        If (success <> True) Then
            strError = sftp.LastErrorText
            Return 2
        Else

        End If

        
        success = sftp.AuthenticatePw(username, password)
        If (success <> True) Then
            strError = sftp.LastErrorText
            Return 3
        Else
            'MsgBox("Funciono Autenticacion", MsgBoxStyle.OkOnly, "Funciono Autenticacion. ")
        End If


        '  After authenticating, the SFTP subsystem must be initialized:
        success = sftp.InitializeSftp()
        If (success <> True) Then
            strError = sftp.LastErrorText
            Return 4
        Else
            'MsgBox("Funciono Inicializacion", MsgBoxStyle.OkOnly, "Funciono Inicializacion. ")
        End If

        success = sftp.UploadFileByName(remoteFile, localFile)
        If (success <> True) Then
            strError = sftp.LastErrorText
            Return 5
        End If

        Return 0

        'MsgBox("Funciono todo", MsgBoxStyle.OkOnly, "EXITO TOTAL")


    End Function


End Module
