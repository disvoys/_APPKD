
Public Class GetMultipleCatia

    Sub MainGetMultipleCatia()

        <DllImport("user32.dll", CharSet:=CharSet.Auto)> Private Shared Sub GetClassName(ByVal hWnd As System.IntPtr, ByVal lpClassName As System.Text.StringBuilder, ByVal nMaxCount As Integer) End Sub
        <DllImport("ole32.dll", ExactSpelling:=True, PreserveSig:=False)> Private Shared Function GetRunningObjectTable(ByVal reserved As Int32) As IRunningObjectTable End Function
        <DllImport("ole32.dll", CharSet:=CharSet.Unicode, ExactSpelling:=True, PreserveSig:=False)> Private Shared Function CreateItemMoniker(ByVal lpszDelim As String, ByVal lpszItem As String) As IMoniker End Function
        <DllImport("ole32.dll", ExactSpelling:=True, PreserveSig:=False)> Private Shared Function CreateBindCtx(ByVal reserved As Integer) As IBindCtx End Function

    End Sub
End Class
