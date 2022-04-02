Imports System.Runtime.InteropServices

Public Class GestionINIFiles
    ' API functions
    <DllImport("Kernel32.dll")>
    Private Shared Function GetPrivateProfileString(ByVal lpApplicationName As String,
      ByVal lpKeyName As String, ByVal lpDefault As String,
      ByVal lpReturnedString As System.Text.StringBuilder,
      ByVal nSize As Integer, ByVal lpFileName As String) _
      As Integer
    End Function

    <DllImport("Kernel32.dll")>
    Private Shared Function WritePrivateProfileString(ByVal lpApplicationName As String,
      ByVal lpKeyName As String, ByVal lpString As String,
      ByVal lpFileName As String) As Integer
    End Function

    <DllImport("Kernel32.dll")>
    Private Shared Function GetPrivateProfileInt(ByVal lpApplicationName As String,
      ByVal lpKeyName As String, ByVal nDefault As Integer,
      ByVal lpFileName As String) As Integer
    End Function

    <DllImport("Kernel32.dll")>
    Private Shared Function FlushPrivateProfileString(ByVal lpApplicationName As Integer,
      ByVal lpKeyName As Integer, ByVal lpString As Integer,
      ByVal lpFileName As String) As Integer
    End Function

    Private _filename As String

    ' Constructor, accepting a filename
    Public Sub New(ByVal Filename As String)
        _filename = Filename
    End Sub

    ' Read-only filename property
    ReadOnly Property FileName() As String
        Get
            Return _filename
        End Get
    End Property

    Public Function GetString(ByVal Section As String,
      ByVal Key As String, ByVal [Default] As String) As String
        ' Returns a string from your INI file
        Dim intCharCount As Integer
        Dim objResult As New System.Text.StringBuilder(256)
        intCharCount = GetPrivateProfileString(Section, Key, [Default], objResult, objResult.Capacity, _filename)
        If intCharCount > 0 Then
            GetString = Left(objResult.ToString, intCharCount)
        Else
            GetString = ""
        End If
    End Function

    Public Function GetInteger(ByVal Section As String,
      ByVal Key As String, ByVal [Default] As Integer) As Integer
        ' Returns an integer from your INI file
        Return GetPrivateProfileInt(Section, Key,
           [Default], _filename)
    End Function

    Public Function GetBoolean(ByVal Section As String,
      ByVal Key As String, ByVal [Default] As Boolean) As Boolean
        ' Returns a boolean from your INI file
        Return (GetPrivateProfileInt(Section, Key,
           CInt([Default]), _filename) = 1)
    End Function

    Public Sub WriteString(ByVal Section As String,
      ByVal Key As String, ByVal Value As String)
        ' Writes a string to your INI file
        WritePrivateProfileString(Section, Key, Value, _filename)
        Flush()
    End Sub

    Public Sub WriteInteger(ByVal Section As String,
      ByVal Key As String, ByVal Value As Integer)
        ' Writes an integer to your INI file
        WriteString(Section, Key, CStr(Value))
        Flush()
    End Sub

    Public Sub WriteBoolean(ByVal Section As String,
      ByVal Key As String, ByVal Value As Boolean)
        ' Writes a boolean to your INI file
        WriteString(Section, Key, CStr(CInt(Value)))
        Flush()
    End Sub

    Private Sub Flush()
        ' Stores all the cached changes to your INI file
        '  FlushPrivateProfileString(0, 0, 0, _filename)
    End Sub

End Class
