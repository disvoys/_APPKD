Imports System.ComponentModel
Imports System.IO
Imports System.Text
Imports CATMat
Imports DRAFTINGITF
Imports HybridShapeTypeLib
Imports INFITF
Imports KnowledgewareTypeLib
Imports MECMOD
Imports PARTITF
Imports ProductStructureTypeLib
Public Class ClassRenameDassault

    Dim ListofRenamed As New List(Of ItemCatia)
    Dim MonNumOutillage As String = ""

    Public Sub New()

    End Sub

    Sub Rename(NumOutillage As String)

        ListofRenamed.Clear()
        NumOutillage = "MA99901Z00-SMSC0020"
        MonNumOutillage = NumOutillage

        RenameItem(ICRacine, NumOutillage, "000000")
        ListofRenamed.Add(ICRacine)
        Dim k As Integer = 1

        Dim i As Integer = 1
        For Each item As TreeViewItem In MonMainV3.MonTV.Items
            For Each tv As TreeViewItem In item.Items
                Dim ic As ItemCatia = GetICfromTV(tv)
                If Not ic Is Nothing And Not ListofRenamed.Contains(ic) Then
                    ListofRenamed.Add(ic)
                    Dim s As String = Format(i, "00") & "0000"
                    RenameItem(ic, NumOutillage, s)
                    i += 1
                End If
                If ic Is Nothing Then
                    RecursiveRename(tv, "01", k)
                    k += 1
                End If
            Next
        Next

        ColDoc.Refresh()

    End Sub

    Sub RecursiveRename(tv As TreeViewItem, i As String, k As Integer)

        For Each iTv As TreeViewItem In tv.Items
            Dim ic As ItemCatia = GetICfromTV(iTv)
            If ic Is Nothing Then
            Else
                If ListofRenamed.Contains(ic) Then
                Else
                    ListofRenamed.Add(ic)
                    Dim s As String = i & Format(k, "0000")
                    RenameItem(ic, MonNumOutillage, s)
                    k += 1
                End If
            End If
            RecursiveRename(iTv, i, k)
        Next

    End Sub

    Sub RenameItem(ic As ItemCatia, NumOutillage As String, indexIC As String)


        ic.l(getItemListProperties("NomPuls_Planche")).Value = indexIC
        ic.PartNumber = NumOutillage & "_" & indexIC
        ic.ProductCATIA.PartNumber = NumOutillage & "_" & indexIC
        Dim p As Parameters = ic.ProductCATIA.UserRefProperties
        FctionCATIA.AddParamatres(ic.Owner, ic)
        p.Item("NomPuls_Planche").value = ic.l(getItemListProperties("NomPuls_Planche")).Value



    End Sub
    Function GetICfromTV(tv As TreeViewItem)


        Dim sp As StackPanel = tv.Header
            Dim im As Image = sp.Children.Item(0)
            Dim lb As Label = sp.Children.Item(1)


            If Right(lb.Content, 3) = " | " Then Return Nothing
            Dim s() As String = Strings.Split(lb.Content, " | ")
            Dim MonPn As String = s(0)

            For Each ic As ItemCatia In ListDocuments
                If ic.PartNumber = MonPn Then

                    Return ic
                End If
            Next

            Return Nothing


    End Function
End Class
