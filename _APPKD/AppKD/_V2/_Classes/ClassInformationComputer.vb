Imports System.Management

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


End Class
