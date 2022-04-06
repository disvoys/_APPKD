Imports System.Net
Imports System.Net.Sockets
Imports System.Text
Imports MySqlConnector

'Version
'ShutDown

'UsersAppKD
'NamePc IPLocale IPPublic Localisation Count Type Ban


Public Class ClassDB

    Const id_ As String = ""
    Const pw_ As String = ""

    Sub ConnectionToDB()

        Dim err As Boolean = False

        With cn
            .ConnectionString = "server=db4free.net;port=3306;user id=" & id_ & ";password=" & pw_ & ";database=testexkd"
            .Open()

        End With

        CheckVersion()
        CheckSiFullban()

    End Sub

    Sub CheckSiFullban()

        cn.Open()

        Dim cmd As New MySqlCommand("Select ShutDown FROM SettingsAppKD", cn)
        Dim reader As MySqlDataReader = cmd.ExecuteReader
        reader.Read()

        Dim BanFull As Integer = 0
        BanFull = reader(0).ToString

        If BanFull = 1 Then
            MonMainV3.ButtonCatiaTest.Visibility = Visibility.Hidden
            Dim m As New MessageErreur("T'es banni mon grand. Rend toi vite sur http://www.catiavb.net et écris moi", Notifications.Wpf.NotificationType.Error)
        Else
            '      Dim merr As New MessageErreur("L'application est à jour", Notifications.Wpf.NotificationType.Information)
        End If

        cn.Close()
    End Sub


    Sub CheckVersion()

        Dim cmd As New MySqlCommand("SELECT Version FROM SettingsAppKD", cn)
        Dim reader As MySqlDataReader = cmd.ExecuteReader
        reader.Read()

        Dim VersionAJour As String = ""
        VersionAJour = reader(0).ToString

        If VersionAJour <> MonMainV3.labelVersio.Text Then
            Dim merr As New MessageErreur("La Mise à jour " & VersionAJour & " est disponible sur www.catiavb.net", Notifications.Wpf.NotificationType.Error)
        Else
            '      Dim merr As New MessageErreur("L'application est à jour", Notifications.Wpf.NotificationType.Information)
        End If

        cn.Close()
    End Sub

    Sub CLoseConnexion()
        cn.Close()
    End Sub

    Function CheckSiClientExists(iplocale, ippublic) As Boolean

        CheckSiClientExists = False
        Try
            cn.Open()
        Catch ex As Exception
            Exit Function
        End Try


        Dim STRcmd As String = "SELECT EXISTS(SELECT * FROM UsersAppKD WHERE IPLocale = '" & iplocale & "' AND IPPublic = '" & ippublic & "')"
        Dim cmd As New MySqlCommand(STRcmd, cn)
        Dim reader As MySqlDataReader = cmd.ExecuteReader
        reader.Read()
        Dim i As Integer = reader(0)
        cn.Close()

        If i = 0 Then
            Return False
        Else
            Return True
        End If



    End Function

    Function CheckSiBan(MonIPv4Locale, MonIPPublic) As Boolean
        CheckSiBan = False
        Try
            cn.Open()
        Catch ex As Exception
            Exit Function

        End Try

        Dim a As String = "SELECT Ban FROM UsersAppKD WHERE IPLocale = '" & MonIPv4Locale & "'" & " AND IPPublic = '" & MonIPPublic & "'"
        Dim cmd As New MySqlCommand(a, cn)
        Dim reader As MySqlDataReader = cmd.ExecuteReader
        reader.Read()
        Dim _c As Integer = reader(0)

        If _c = 0 Then
            cn.Close()
            Return False
        Else
            cn.Close()
            Return True
        End If


    End Function
    Sub CheckUser()

        Dim NamePC As String = System.Net.Dns.GetHostName()
        Dim MonIPv4Locale As String = ""

        Dim iphe As System.Net.IPHostEntry = System.Net.Dns.GetHostEntry(NamePC)
        For Each ipheal As System.Net.IPAddress In iphe.AddressList
            If ipheal.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then
                MonIPv4Locale = ipheal.ToString()
                Exit For
            End If
        Next

        Dim MonIPPublic As String = GetExternalIp()
        cn.Close()

        If CheckSiClientExists(MonIPv4Locale, MonIPPublic) = True Then
            If CheckSiBan(MonIPv4Locale, MonIPPublic) = True Then
                MonMainV3.ButtonCatiaTest.Visibility = Visibility.Hidden
                Dim m As New MessageErreur("T'es banni mon grand. Rend toi vite sur http://www.catiavb.net et écris moi", Notifications.Wpf.NotificationType.Error)
            Else
                cn.Open()
                Dim a As String = "SELECT Count FROM UsersAppKD WHERE IPLocale = '" & MonIPv4Locale & "'" & " AND IPPublic = '" & MonIPPublic & "'"
                Dim cmd As New MySqlCommand(a, cn)
                Dim reader As MySqlDataReader = cmd.ExecuteReader
                reader.Read()
                Dim _c As Integer = reader(0) + 1
                cn.Close()
                cn.Open()
                Dim f As String = "UPDATE UsersAppKD SET count = " & _c & " WHERE IPLocale = '" & MonIPv4Locale & "'" & " AND IPPublic = '" & MonIPPublic & "'"
                Dim cmd_ As New MySqlCommand(f, cn)
                cmd_.ExecuteNonQuery()
                cn.Close()
                '   MsgBox(_c)
            End If

        Else
            cn.Open()
            Dim STRcmd As String = "INSERT INTO UsersAppKD (NamePC,IPLocale,IPPublic,Count,Ban) VALUES (@NamePC, @IPLocale, @IPPublic, @Count, @Ban)"
            Dim cmd As New MySqlCommand(STRcmd, cn)
            cmd.Parameters.AddWithValue("@NamePC", NamePC)
            cmd.Parameters.AddWithValue("@IPLocale", MonIPv4Locale)
            cmd.Parameters.AddWithValue("@IPPublic", MonIPPublic)
            cmd.Parameters.AddWithValue("@Count", 1)
            cmd.Parameters.AddWithValue("@Ban", 0)

            cmd.ExecuteNonQuery()
            cn.Close()
        End If


        Dim CountUsed As Integer = 1

        Dim Ban As Integer = 0

    End Sub

    Function GetExternalIp() As String
        Dim ExternalIP As String = "0.0.0.0"
        '  Return ExternalIP 'test KEVIN DESVOIS
        ExternalIP = (New WebClient()).DownloadString("http://checkip.dyndns.org/")
        ExternalIP = (New RegularExpressions.Regex("\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}")).Matches(ExternalIP)(0).ToString()
        Return ExternalIP
    End Function


    Function IsPortOpen(ByVal Host As String, ByVal PortNumber As Integer) As Boolean
        Dim Client As TcpClient = Nothing
        Try
            Client = New TcpClient(Host, PortNumber)
            Return True
        Catch ex As SocketException
            Return False
        Finally
            If Not Client Is Nothing Then
                Client.Close()
            End If
        End Try
    End Function
End Class

