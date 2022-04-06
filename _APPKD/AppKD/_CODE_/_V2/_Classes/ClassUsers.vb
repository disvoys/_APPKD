Imports System.Data
Imports MySqlConnector

Public Class ClassUsers



    Public Property Nom As String
    Public Property IPLocale As String
    Public Property IPPublic As String
    Public Property Connexions As Integer
    Public Property Ban As Integer


    Sub New(Nom_, IPlocale_, IPPublic_, Connextions_, ban_)

        Nom = Nom_
        IPLocale = IPlocale_
        IPPublic = IPPublic_
        Connexions = Connextions_
        Ban = ban_

    End Sub



End Class


Public Class LoadDTUsers
    Const id_ As String = ""
    Const pw_ As String = ""
    Sub Load()

        ConnexionToDB()
        cn_.Close()


    End Sub

    Sub ConnexionToDB()
        Dim err As Boolean = False
        Try
            With cn_
                .ConnectionString = "server=db4free.net;port=3306;user id=" & id_ & ";password=" & pw_ & ";database=testexkd;Connect Timeout=1000;pooling=true"
                .Open()
            End With
        Catch ex As Exception
            err = True
        End Try

        If err = True Then

        Else
            LoadData()
        End If

    End Sub

    Function LoadData()

        dt_.Clear()
        Dim cmd As New MySqlCommand("SELECT * FROM UsersAppKD", cn_)
        sda_ = New MySqlDataAdapter(cmd)
        sda_.Fill(dt_)
        ListLesUsers()
        Return dt_

    End Function

    Sub ListLesUsers()

        ListUsers.Clear()
        QteUsers = 0
        ListUsersstr.Clear()

        For Each item In dt_.Rows
            Dim _do As New ClassUsers(item(0).ToString, item(1).ToString, item(2).ToString, item(3).ToString, item(4).ToString)
            ListUsers.Add(_do)
        Next

    End Sub

    Sub BanUser(_iplocale As String, _ippublic As String, _ban As Integer)



        cn_.Open()

        If cn_.State = ConnectionState.Open Then
            Dim STRcmd As String = "UPDATE UsersAppKD SET ban = " & _ban & " WHERE `IPLocale`= '" & _iplocale & "' AND `IPPublic`= '" & _ippublic & "'"
            Dim cmd As New MySqlCommand(STRcmd, cn_)
            cmd.ExecuteNonQuery()

        End If

        cn_.Close()

    End Sub

    Sub RemoveUser(_iplocale As String, _ippublic As String)


        If Not cn.State = ConnectionState.Open Then cn_.Open()

        If cn_.State = ConnectionState.Open Then
            Dim STRcmd As String = "DELETE FROM UsersAppKD WHERE `IPLocale`= '" & _iplocale & "' AND `IPPublic`= '" & _ippublic & "'"
            Dim cmd As New MySqlCommand(STRcmd, cn_)
            cmd.ExecuteNonQuery()
        End If

        cn_.Close()

    End Sub

    Sub UpdateTableSettingsG(a As String, b As Integer)

        cn_.Open()

        If cn_.State = ConnectionState.Open Then
            Dim STRcmd As String = "UPDATE SettingsAppKD SET Version = '" & a & "' , ShutDown = " & b
            Dim cmd As New MySqlCommand(STRcmd, cn_)
            cmd.ExecuteNonQuery()

        End If

        cn_.Close()

    End Sub

End Class
