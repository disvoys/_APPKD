Imports System.Net
Imports System.Net.Sockets
Imports System.Text
Imports MySqlConnector

'Version
'ShutDown

'UsersAppKD
'NamePc IPLocale IPPublic Localisation Count Type Ban


Public Class ClassDB

    Public cmd As New MySqlCommand("", cn)
    Sub ConnectionToDB()

        With cn
            .ConnectionString = "server=*****;user id=*****;password=*****;database=*****"
            .Open()
        End With

        CheckVersion()
            CheckSiFullban()
        CheckUser()

        'cn.Close()


    End Sub

    Sub CheckSiFullban()

        cmd.CommandText = "Select shutdown FROM settingsAPPKD"
        Dim reader As MySqlDataReader = cmd.ExecuteReader
        reader.Read()

        Dim BanFull As Integer = 0
        BanFull = reader(0).ToString

        If BanFull = 1 Then
            MonMainV3.ButtonCatiaTest.Visibility = Visibility.Hidden
            MsgBox("Tu n'as plus l'autorisation d'utiliser l'application. Contacte moi sur http://www.catiavb.net pour faire évoluer ton status", MsgBoxStyle.Critical)
            End
        Else
            Dim merr As New MessageErreur("L'application est à jour", Notifications.Wpf.NotificationType.Information)
        End If

        reader.Close()
    End Sub
    Sub CheckVersion()

        cmd.CommandText = "SELECT version FROM settingsAPPKD"
        Dim reader As MySqlDataReader = cmd.ExecuteReader
        reader.Read()

        Dim VersionAJour As String = reader(0).ToString

        If VersionAJour <> MonMainV3.labelVersio.Text Then
            Dim merr As New MessageErreur("La Mise à jour " & VersionAJour & " est disponible sur www.catiavb.net", Notifications.Wpf.NotificationType.Error)
        End If

        reader.Close()

    End Sub

    Sub CLoseConnexion()
        cn.Close()
    End Sub

    Function CheckSiClientExists(iplocale, ippublic) As Boolean

        Dim STRcmd As String = "SELECT EXISTS(SELECT * FROM bddUsersAPPKD WHERE CarteMere = '" & iplocale & "' AND IPPublic = '" & ippublic & "')"
        cmd.CommandText = STRcmd
        Dim reader As MySqlDataReader = cmd.ExecuteReader
        reader.Read()
        Dim i As Integer = reader(0)
        reader.Close()

        If i = 0 Then
            Return False
        Else
            Return True
        End If


    End Function

    Function CheckSiBan(MonIPv4Locale, MonIPPublic) As Boolean
        CheckSiBan = False
        Dim a As String = "SELECT Ban FROM bddUsersAPPKD WHERE CarteMere = '" & MonIPv4Locale & "'" & " AND IPPublic = '" & MonIPPublic & "'"
        cmd.CommandText = a
        Dim reader As MySqlDataReader = cmd.ExecuteReader
        reader.Read()
        Dim _c As Integer = reader(0)
        reader.Close()
        If _c = 0 Then
            Return False
        Else
            Return True
        End If


    End Function
    Sub CheckUser()

        Dim NamePC As String = System.Net.Dns.GetHostName()
        Dim CM As String = FctionGetInfo.GetCarteMere()
        Dim MonIPPublic As String = FctionGetInfo.GetExternalIp()

        If CheckSiClientExists(CM, MonIPPublic) = True Then
            If CheckSiBan(CM, MonIPPublic) = True Then
                MonMainV3.ButtonCatiaTest.Visibility = Visibility.Hidden
                MsgBox("Tu n'as plus l'autorisation d'utiliser l'application. Contacte moi sur http://www.catiavb.net pour faire évoluer ton status", MsgBoxStyle.Critical)
                End
            Else
                Dim a As String = "SELECT Count FROM bddUsersAPPKD WHERE CarteMere = '" & CM & "'" & " AND IPPublic = '" & MonIPPublic & "'"
                cmd.CommandText = a
                Dim reader As MySqlDataReader = cmd.ExecuteReader
                reader.Read()
                Dim _c As Integer = reader(0) + 1
                reader.Close()
                Dim f As String = "UPDATE bddUsersAPPKD SET count = " & _c & " WHERE CarteMere = '" & CM & "'" & " AND IPPublic = '" & MonIPPublic & "'"
                cmd.CommandText = f
                cmd.ExecuteNonQuery()
            End If

        Else
            Dim STRcmd As String = "INSERT INTO bddUsersAPPKD (NamePC,CarteMere,IPPublic,Count,Ban) VALUES (@NamePC, @IPLocale, @IPPublic, @Count, @Ban)"
            cmd.CommandText = STRcmd
            cmd.Parameters.AddWithValue("@NamePC", NamePC)
            cmd.Parameters.AddWithValue("@IPLocale", CM)
            cmd.Parameters.AddWithValue("@IPPublic", MonIPPublic)
            cmd.Parameters.AddWithValue("@Count", 1)
            cmd.Parameters.AddWithValue("@Ban", 0)
            cmd.ExecuteNonQuery()
        End If

    End Sub




End Class

