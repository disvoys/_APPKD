Imports Notifications.Wpf

Public Class MessageErreur

    Public Property Text As String
    Public Property Visible As Visibility


    Sub New(message As String, notificationType As NotificationType)
        Dim N As New NotificationManager
        Dim C As New NotificationContent
        C.Title = "Application KD | V3.0.19"
        C.Message = message
        C.Type = notificationType
        N.Show(C)
    End Sub

End Class
