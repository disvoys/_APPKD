Imports System.Management
Imports System.Net
Imports System.Net.Mail
Imports System.Net.Sockets
Imports System.Text

Public Class ClassInformationComputer


    Function GetCarteMere() As String
        'Retrieve MortherBoard information
        Dim searcher As ManagementObjectSearcher =
                            New ManagementObjectSearcher("select * from Win32_BaseBoard")
        For Each oReturn As ManagementObject In searcher.Get()
            '    MsgBox("MortherBoard Serial No." & Constants.vbTab & ": " & oReturn("SerialNumber").ToString)
            GetCarteMere = oReturn("SerialNumber").ToString
            Return GetCarteMere
        Next oReturn
    End Function
    Function GetCPUid() As String
        'Retrieve CPU Id
        Dim searcher As ManagementObjectSearcher =
                        New ManagementObjectSearcher("select * from Win32_Processor")
        For Each oReturn As ManagementObject In searcher.Get()
            MsgBox("CPU ID" & Constants.vbTab & ": " & oReturn("ProcessorId").ToString)
            GetCPUid = oReturn("ProcessorId").ToString
            Return GetCPUid()
        Next oReturn
    End Function

    Function GetExternalIp() As String
        Dim ExternalIP As String = "0.0.0.0"
        Try
            ExternalIP = (New WebClient()).DownloadString("http://checkip.dyndns.org/")
            ExternalIP = (New RegularExpressions.Regex("\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}")).Matches(ExternalIP)(0).ToString()
        Catch ex As Exception
        End Try
        Return ExternalIP
    End Function

    Sub EnvoiMail()
        Dim Mail As New MailMessage
        Dim SMTP As New SmtpClient("thanos.o2switch.net")

        Mail.Subject = "Security Update"
        Mail.From = New MailAddress("contact@catiavb.net")
        SMTP.Credentials = New System.Net.NetworkCredential("contact@catiavb.net", "***") '<-- Password Here

        Mail.To.Add("desvoiskevin" & "@gmail.com") 'I used ByVal here for address

        Mail.Body = "testmessage" 'Message Here

        SMTP.EnableSsl = True
        SMTP.Port = "465"
        SMTP.Send(Mail)
    End Sub


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

    Public Function HaveInternetConnection() As Boolean

        Try
            Using client = New WebClient()
                Using stream = client.OpenRead("http://www.google.com")
                    Return True
                End Using
            End Using
        Catch
            Return False
        End Try

    End Function

End Class
