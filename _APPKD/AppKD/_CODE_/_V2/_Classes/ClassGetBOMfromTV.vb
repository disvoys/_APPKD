Imports System.Text

Public Class ClassGetBOMfromTV

    Sub BOMAllElements()

        Using sr As IO.StreamReader = New IO.StreamReader(My.Computer.FileSystem.SpecialDirectories.Temp & "\BOM_.txt", Encoding.GetEncoding("iso-8859-1"))
            Dim BoolStart_ As Boolean = False
            While Not sr.EndOfStream
                Dim line As String = sr.ReadLine
                If line = "[ALL-BOM-APPKD]" Then
                    BoolStart_ = True
                End If
                If BoolStart_ = True Then
                    If Left(line, 1) = "[" And Right(line, 1) = "]" Then
                    Else
                        Dim l_() As String = Strings.Split(line, vbTab)
                        Dim MonPN As String = l_(0)
                        Dim MaQte As String = l_(1)

                        For Each d As ItemCatia In ListDocuments
                            If d.PartNumber = MonPN Then
                                d.Visible = True
                                d.Qte = MaQte
                                Exit For
                            End If
                        Next
                    End If
                End If
            End While
        End Using

    End Sub

    Sub BOMoneLEVEL(pnRACINE As String, qte As Integer)

        Dim VerifDoublonCoponents As Boolean = False
        Dim nameComponentsError As String = ""
        Dim ListCoponentsDoublon As New List(Of String)

        Using sr As IO.StreamReader = New IO.StreamReader(My.Computer.FileSystem.SpecialDirectories.Temp & "\BOM_.txt", Encoding.GetEncoding("iso-8859-1"))
            Dim BoolStart_ As Boolean = False
            While Not sr.EndOfStream
                Dim line As String = sr.ReadLine
                If line = "[" & pnRACINE & "]" Then
                    BoolStart_ = True
                    line = sr.ReadLine
                End If
                If BoolStart_ = True Then
                    If Left(line, 1) = "[" And Right(line, 1) = "]" Then
                        BoolStart_ = False
                    Else
                        Dim l_() As String = Strings.Split(line, vbTab)
                        Dim MonPN As String = l_(0)
                        Dim MaQte As String = l_(1)
                        Dim VerifComponents As Boolean = True
                        For Each d As ItemCatia In ListDocuments
                            If d.PartNumber = MonPN Then
                                If d.Type = "PART" Then
                                    d.Visible = True
                                    d.Qte = MaQte * qte + Convert.ToInt32(d.Qte)
                                    VerifComponents = False
                                    Exit For
                                End If
                                If d.Type = "PRODUCT" Then
                                    d.Visible = True
                                    d.Qte = MaQte * qte + Convert.ToInt32(d.Qte)
                                    VerifComponents = False
                                    Exit For
                                End If
                                If d.Type = "RACINE" Then
                                    d.Visible = True
                                    d.Qte = MaQte * qte + Convert.ToInt32(d.Qte)
                                    VerifComponents = False
                                    Exit For
                                End If
                            End If
                        Next
                        If VerifComponents = True Then
                            BOMoneLEVEL(MonPN, MaQte * qte)
                        End If
                    End If
                End If
            End While
        End Using

    End Sub

    Sub VerifAllComponentsDoublon()

        Dim nameComponentsError As String = ""
        Dim ListCoponentsDoublon As New List(Of String)

        Using sr As IO.StreamReader = New IO.StreamReader(My.Computer.FileSystem.SpecialDirectories.Temp & "\BOM_.txt", Encoding.GetEncoding("iso-8859-1"))
            While Not sr.EndOfStream
                Dim line As String = sr.ReadLine
                If Left(line, 1) = "[" And Right(line, 1) = "]" Then
                    If ListCoponentsDoublon.Contains(line) Then
                        nameComponentsError = line
                    Else
                        ListCoponentsDoublon.Add(line)
                    End If
                End If
            End While
        End Using

        If nameComponentsError <> "" Then
            Dim m As New MessageErreur("ERREUR sur la nomenclature. Le composant " & nameComponentsError & " existe en plusieurs entités : rennomer en un.", Notifications.Wpf.NotificationType.Error)
        End If

    End Sub
    Sub ResetBOM()

        For Each d As ItemCatia In ListDocuments
            d.Visible = False
            d.Qte = 0
        Next

    End Sub


    Sub GoBOM(pn As String)

        For Each d As ItemCatia In ListDocuments
            d.Visible = False
            d.Qte = 0
        Next

        Dim MaRacine As ItemCatia = GetRacine(pn)
        If MaRacine Is Nothing Then
            Dim m As New MessageErreur("Une erreur s'est produite. Impossible de récupérer la nomenclature de [" & pn & "]", Notifications.Wpf.NotificationType.Error)
            Exit Sub
        Else
            If pn = "Ensemble des éléments" Then
                BOMAllElements()
            Else
                ResetBOM()
                BOMoneLEVEL(pn, 1)
                VerifAllComponentsDoublon()
            End If
        End If

        For Each d As ItemCatia In ListDocuments
            If d.Qte = 0 Then d.Qte = ""
        Next

    End Sub

    Function GetRacine(PN As String) As ItemCatia

        If PN = "Ensemble des éléments" Then Return ICRacine
        For Each item As ItemCatia In ListDocuments
            If item.PartNumber = PN Then
                Return item
            End If
        Next

        Return Nothing

    End Function 'RECUPERE LE TV RACINE [OK]

End Class
