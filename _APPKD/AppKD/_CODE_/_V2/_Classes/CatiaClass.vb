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

Public Class CatiaClass



    Sub test()

        Dim d As DrawingDocument = CATIA.ActiveDocument 'need to be a CATDrawing opened
        Dim t As DrawingTable = d.Sheets.ActiveSheet.Views.ActiveView.Tables.Item(1)

        Dim i As Integer = 1
        Dim p As Product = CATIA.Documents.Item("Part1.CATPart").product 'i let you find your Part

        For Each MyParam As Parameter In p.Parameters
            If MyParam.Name Like "*Angle.*" Then 'i let you find a way to get your names
                t.SetCellString(i, 1, MyParam.ValueAsString)
                i += 1
            End If
        Next



    End Sub



    Function GetPathCATIA() As String

        Return CATIA.SystemService.Environ("CATDLLPath")

    End Function

#Region "Renommage"
    Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hwndParent As Long, ByVal hwndChildAfter As Long, ByVal lpszClass As Long, ByVal lpszWindow As String) As Long
    Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As String) As Long
    Private Const BM_CLICK = &HF5


    Private Sub ClickOk()
        Try
            Dim lSaveAll As Long
            Dim lOK As Long

            lSaveAll = FindWindow("#32770", "Save All")
            Do While lSaveAll = 0
                lSaveAll = FindWindow("#32770", "Save All")
            Loop

            lOK = FindWindowEx(lSaveAll, 0&, 0&, "&Yes")
            Do While lOK = 0
                lOK = FindWindowEx(lSaveAll, 0&, 0&, "&Yes")
            Loop

            SendMessage(lOK, BM_CLICK, 0, 0)
            SendMessage(lOK, BM_CLICK, 0, 0)
        Catch ex As Exception
        End Try
    End Sub
    Sub SaveIC(ic As ItemCatia)

        Dim OldName As String = Nothing
        Dim Path As String = Nothing
        Dim NewName As String = Nothing

        For Each d As Document In CATIA.Documents
            If d.FullName = ic.FileName Then
                OldName = ic.FileName
                Path = d.Path
                CATIA.DisplayFileAlerts = False
                Dim sauv As String = ""
                Select Case ic.Type
                    Case "RACINE"
                        sauv = Path & "\" & ic.PartNumber & ".CATProduct"
                    Case "PRODUCT"
                        sauv = Path & "\" & ic.PartNumber & ".CATProduct"
                    Case "PART"
                        sauv = Path & "\" & ic.PartNumber & ".CATPart"
                End Select

                If Path = "" Then
                    Dim m As New MessageErreur("Une erreur s'est produite. Vérifier que le fichier soit enregistré", Notifications.Wpf.NotificationType.Error)
                Else
                    d.SaveAs(sauv)
                    ic.FileName = sauv
                    NewName = ic.FileName
                    ic.Owner = ic.Doc.Name
                    ic.PartNumber = ic.ProductCATIA.PartNumber
                    CATIA.DisplayFileAlerts = True
                End If
                Exit For
            End If
        Next

        For Each d As Document In CATIA.Documents
            If Not d.Saved Then
                Try
                    d.Save()
                Catch ex As Exception
                    'Dim m As New MessageErreur("Une erreur s'est produite lors de la sauvegarde du fichier. Vérifier que le fichier existe", Notifications.Wpf.NotificationType.Warning)
                End Try
            End If

        Next

        If OldName <> NewName Then
            If IO.File.Exists(OldName) Then CATIA.FileSystem.DeleteFile(OldName)
        End If




    End Sub

#End Region
    Dim ArrayDataSetType()
    Dim ArrayMfgProcess()
    Dim ArrayDesignAutority()
    Dim ArraySupplierName()
    Dim ArrayMajorSupplierCode()
    Dim Array3donly()
    Dim ArrayColorCoded()
    Dim ArrayMaterialForm()
    Dim ArrayFinishCode()
    Dim ArrayWireGauge()
    Dim ArrayForm()
    Dim ArrayType()
    Dim ArrayStyle()
    Dim ArrayDensity()
    Dim ArrayMaterialSpec()
    Dim ArrayMaterialDescription()
    Dim ArrayAllow()
    Dim ArrayMaterialClass()
    Dim ArrayGrade()
    Dim ArrayFinalCondition()
    Dim ArrayStandardSpecDie()
    Dim ArrayMesh()
    Dim ArrayPCCN()
    Dim ArrayProject()

    Sub LinkICtoTV()

        For Each ic As ItemCatia In ListDocuments
            For Each i As TreeViewItem In MonMainV3.MonTV.Items
                If i.Header = ic.PartNumber & " | " & ic.DescriptionRef Then
                    ic.ListTVitem_.Add(i)
                End If
            Next
        Next

    End Sub

    Sub recursiveLinkICtoTV(ic As ItemCatia, i As TreeViewItem)

        For Each i_ As TreeViewItem In i.Items
            If i_.Header = ic.PartNumber & " | " & ic.DescriptionRef Then
                ic.ListTVitem_.Add(i_)
            End If
            recursiveLinkICtoTV(ic, i_)
        Next
    End Sub


    Public Sub ResetPropoerties()
        On Error Resume Next

        For Each ic In ListDocuments
            Dim P As Parameters = ic.ProductCATIA.UserRefProperties
            Dim i As Integer = 1
            For n = 1 To P.Count
                If P.Item(i).UserAccessMode = 2 Then
                    P.Remove(i)
                Else
                    P.Item(i).Hidden = True
                    If Not P.Item(i).ReadOnly Then
                        P.Item(i).ValuateFromString("")
                        i = i + 1
                    End If
                End If
            Next
        Next
        On Error GoTo 0
    End Sub

    Function GetStringFromText(Text As String, v As DrawingView)

        Dim k As String = ""
        For Each t As DrawingText In v.Texts
            If t.Name = Text Then
                k = t.Text
            End If
        Next

        Return k
    End Function

    Sub CheckLangue(p As Product)

        Dim NomFichier As String = My.Computer.FileSystem.SpecialDirectories.Temp & "\BOM.txt"
        Dim AssConvertor As AssemblyConvertor
        AssConvertor = p.GetItem("BillOfMaterial")
        Dim nullstr(0)
        AssConvertor.SetCurrentFormat(nullstr)
        Dim VarMaListNom(0)
        AssConvertor.SetSecondaryFormat(VarMaListNom)
        AssConvertor.Print("HTML", NomFichier, p)

        Dim fs As FileStream = Nothing
        If IO.File.Exists(NomFichier) Then
            Using sr As StreamReader = New StreamReader(NomFichier, Encoding.GetEncoding("iso-8859-1"))

                While Not sr.EndOfStream
                    Dim line As String = sr.ReadLine
                    If line Like "<b>Pièces différentes :*<br*" Then
                        MaLangue = "Francais"
                    ElseIf line Like "<b>Different parts:*<br*" Then
                        MaLangue = "Anglais"
                    End If
                End While
                sr.Close()
            End Using
        End If

    End Sub
    Public Sub GetBOM(p As Product)

        Dim NomFichier As String = My.Computer.FileSystem.SpecialDirectories.Temp & "\BOM.txt"
        Dim AssConvertor As AssemblyConvertor
        AssConvertor = p.GetItem("BillOfMaterial")
        Dim nullstr(2)
        If MaLangue = "Anglais" Then
            nullstr(0) = "Part Number"
            nullstr(1) = "Quantity"
            nullstr(2) = "Type"
        ElseIf MaLangue = "Francais" Then
            nullstr(0) = "Référence"
            nullstr(1) = "Quantité"
            nullstr(2) = "Type"
        End If

        AssConvertor.SetCurrentFormat(nullstr)

        Dim VarMaListNom(1)
        If MaLangue = "Anglais" Then
            VarMaListNom(0) = "Part Number"
            VarMaListNom(1) = "Quantity"
        ElseIf MaLangue = "Francais" Then
            VarMaListNom(0) = "Référence"
            VarMaListNom(1) = "Quantité"
        End If

        AssConvertor.SetSecondaryFormat(VarMaListNom)
        AssConvertor.Print("HTML", NomFichier, p)

        ModifFichierNomenclature(My.Computer.FileSystem.SpecialDirectories.Temp & "\BOM.txt")


    End Sub

    Sub ModifFichierNomenclature(txt As String)

        Dim strtocheck As String = ""
        If MaLangue = "Francais" Then
            strtocheck = "<b>Total des p"
        Else
            strtocheck = "<b>Total parts"
        End If

        Dim FichierNomenclature As String = My.Computer.FileSystem.SpecialDirectories.Temp & "\BOM_.txt"
        If IO.File.Exists(FichierNomenclature) Then
            IO.File.Delete(FichierNomenclature)
        End If
        Dim fs As FileStream = Nothing
        fs = New FileStream(FichierNomenclature, FileMode.CreateNew)
        Using sw As StreamWriter = New StreamWriter(fs, Encoding.GetEncoding("iso-8859-1"))
            If IO.File.Exists(txt) Then
                Using sr As StreamReader = New StreamReader(txt, Encoding.GetEncoding("iso-8859-1"))
                    Dim BoolStart As Boolean = False
                    While Not sr.EndOfStream
                        Dim line As String = sr.ReadLine
                        If Left(line, 8) = "<a name=" Then
                            If MaLangue = "Francais" Then
                                line = "[" & Right(line, line.Length - 24)
                                line = Left(line, line.Length - 8)
                                line = line & "]"
                                sw.WriteLine(line)
                            Else
                                line = "[" & Right(line, line.Length - 27)
                                line = Left(line, line.Length - 8)
                                line = line & "]"
                                sw.WriteLine(line)
                            End If
                        ElseIf line Like "  <tr><td><A HREF=*</td> </tr>*" Then
                            line = Replace(line, "</td><td>Assembly</td> </tr>", "") 'pas fait
                            line = Replace(line, "</td><td>Assemblage</td> </tr> ", "")
                            line = Replace(line, "  <tr><td><A HREF=", "")
                            line = Replace(line, "</A></td><td>", ControlChars.Tab)
                            line = Replace(line, "#Bill of Material: ", "")
                            line = Replace(line, "#Nomenclature : ", "")
                            If line.Contains(">") Then
                                Dim lines() = Strings.Split(line, ">")
                                line = lines(1)
                            End If
                            Dim lines_() = Strings.Split(line, ControlChars.Tab)
                            line = lines_(0) & ControlChars.Tab & lines_(1)
                            If Strings.Left(line, 2) = "  " Then line = Strings.Right(line, line.Length - 2)
                            sw.WriteLine(line)
                        ElseIf Left(line, 14) = strtocheck Then
                            sw.WriteLine("[ALL-BOM-APPKD]")
                        ElseIf line Like "*<tr><td>*</td> </tr>*" Then
                            line = Replace(line, "<tr><td>", "")
                            line = Replace(line, "</td> </tr> ", "")
                            line = Replace(line, "</td><td>", ControlChars.Tab)
                            Dim lines_() = Strings.Split(line, ControlChars.Tab)
                            line = lines_(0) & ControlChars.Tab & lines_(1)
                            If Strings.Left(line, 2) = "  " Then line = Strings.Right(line, line.Length - 2)
                            sw.WriteLine(line)
                        Else
                            'nothing
                        End If

                    End While
                    sr.Close()
                End Using
            End If
            sw.Close()
        End Using

    End Sub

    Sub ArboDassault(n As String)

        CATIA.DisplayFileAlerts = False

        Dim pr As ProductDocument = CATIA.Documents.Open(DossierBase & "\" & INIProperties.GetString(GetEnv, "TemplateARBORESCENCE", ""))

        'produit de tete
        Dim p As Product = pr.Product
        p.Name = n
        p.PartNumber = n

        'PartCTRL
        Dim pCTRL As Product = p.Products.Item(2)
        pCTRL.Name = n & "_CTR001"
        pCTRL.PartNumber = n & "_CTR001"


        'ProductOUTILLAGE
        Dim pOUT As Product = p.Products.Item(3)
        pOUT.Name = n & "_010000"
        pOUT.PartNumber = n & "_010000"

        'PartGEO
        Dim pPARTGEO As Product = p.Products.Item(3).Products.Item(1)
        pPARTGEO.Name = n & "_999999"
        pPARTGEO.PartNumber = n & "_999999"


        Dim s As String = Path.GetTempPath
        CATIA.Documents.Item("MA00001Z00-REFBE_CTR001.CATPart").SaveAs(s & pCTRL.PartNumber & ".CATPart")
        CATIA.Documents.Item("MA00001Z00-REFBE_999999.CATPart").SaveAs(s & pPARTGEO.PartNumber & ".CATPart")
        CATIA.Documents.Item("MA00001Z00-REFBE_010000.CATProduct").SaveAs(s & pOUT.PartNumber & ".CATProduct")
        CATIA.Documents.Item("MA00001Z00-REFBE_000000.CATProduct").SaveAs(s & n & ".CATProduct")


        CATIA.DisplayFileAlerts = True



    End Sub

    Sub ArboAirbus(n As String) '185D15123123456789
        Try
            Dim prd As ProductDocument = CATIA.Documents.Add("Product")
            Dim p As Product = prd.Product
            p.Name = n
            p.PartNumber = n
            p.Products.AddNewProduct("ENV")
            Dim p2 As Product = p.Products.AddNewComponent("Product", n & "000")
            Dim p3 As Product = p2.Products.AddNewComponent("Product", n & "A100")
            p3.Products.AddNewComponent("Part", n & "1002")
            p3.Products.AddNewComponent("Part", n & "1004")
            p3.Products.AddNewComponent("Part", n & "1006")

            Dim m As New MessageErreur("L'arborescence a été générée avec succès", Notifications.Wpf.NotificationType.Information)
        Catch ex As Exception
            Dim m As New MessageErreur("Une erreur s'est produite. Vérifier qu'aucun autre élément existe déjà à ce nom.", Notifications.Wpf.NotificationType.Error)
        End Try


    End Sub
    Public Sub GetCATIA()

        CATIA = GetObject(, "CATIA.Application")

    End Sub
    Sub CheckActiveDoc()
        Try
            MonActiveDoc = CATIA.ActiveDocument
            If Right(MonActiveDoc.FullName, 4) = "Part" Then TypeActiveDoc = "PART"
            If Right(MonActiveDoc.FullName, 7) = "Product" Then TypeActiveDoc = "PRODUCT"
            If Right(MonActiveDoc.FullName, 7) = "Drawing" Then TypeActiveDoc = "DRAWING"

            If TypeActiveDoc = "PRODUCT" Then
                CheckLangue(MonActiveDoc.product)
                GetBOM(MonActiveDoc.product)
            End If

        Catch ex As Exception
            MonActiveDoc = Nothing
            'err aucun document n'est ouvert
        End Try

        If MonActiveDoc Is Nothing Then
            'err le document ouvert n'est pas un product
        End If
    End Sub
    Sub AddParamatres(MyOwner As String, ic As ItemCatia) '#MODIF 08/03/2017

        Try
            If ic Is Nothing Or ic.Type = "PART" Or ic.Type = "PRODUCT" Or ic.Type = "RACINE" Then

                Dim ListeDesParamsToAdd As New List(Of String)
                Dim ListeDesParams As New List(Of String)
                Dim ListeDesValeurs As New List(Of String)
                Dim ListParamAPasSupprimer As New List(Of String)

                Dim d As Document = CATIA.Documents.Item(MyOwner)
                Dim P As Product = d.Product
                Dim MesParams As Parameters = P.UserRefProperties


                'save param + values
                ListeDesParams.Clear()
                For i = 1 To MesParams.Count
                    Dim k = Len(MesParams.Item(i).Name)
                    Dim e2 = InStrRev(MesParams.Item(i).Name, "\")
                    Dim str As String = Strings.Right(MesParams.Item(i).Name, k - e2)
                    ListeDesParams.Add(str)
                    ListeDesValeurs.Add(MesParams.Item(i).ValueAsString)
                Next


                For Each item As ItemProperties In MonMainV3.DataGridProperties.Items

                    Dim NameParam As String = item.Properties.ToString
                    Dim ValueParam As String = item.Valeur.ToString
                    Dim TypeParam As String = UCase(item.Type.ToString)
                    If NameParam = Nothing Then GoTo Boucle
                    If TypeParam = Nothing Then GoTo Boucle

                    ListParamAPasSupprimer.Add(NameParam)
                    ListeDesParamsToAdd.Add(NameParam)

                    If Not ListeDesParams.Contains(NameParam) Then

                        Dim MonParametreToUse As Parameter

                        Select Case TypeParam
                            Case "INTEGER"
                                Dim intParam1 As IntParam = MesParams.CreateInteger(NameParam, ValueParam.ToString)
                                MonParametreToUse = intParam1
                            Case "STRING"
                                Dim strParam1 As StrParam = MesParams.CreateString(NameParam, ValueParam.ToString)
                                MonParametreToUse = strParam1

                            Case "MASS"
                                Dim dimension1 As Dimension
                                dimension1 = MesParams.CreateDimension(NameParam.ToString, "MASS", 0#)
                                dimension1.ValuateFromString(ValueParam.ToString)
                                dimension1.Rename(NameParam.ToString)
                                MonParametreToUse = dimension1
                            Case "VOLUME"
                                Dim dimension1 As Dimension
                                dimension1 = MesParams.CreateDimension(NameParam.ToString, "VOLUME", 0#)
                                dimension1.ValuateFromString(ValueParam.ToString)
                                dimension1.Rename(NameParam.ToString)
                                MonParametreToUse = dimension1
                            Case "AREA"
                                Dim dimension1 As Dimension
                                dimension1 = MesParams.CreateDimension(NameParam.ToString, "AREA", 0#)
                                dimension1.ValuateFromString(ValueParam.ToString)
                                dimension1.Rename(NameParam.ToString)
                                MonParametreToUse = dimension1
                            Case "LENGTH"
                                Dim dimension1 As Dimension
                                dimension1 = MesParams.CreateDimension(NameParam.ToString, "LENGTH", 0#)
                                dimension1.ValuateFromString(ValueParam.ToString)
                                dimension1.Rename(NameParam.ToString)
                                MonParametreToUse = dimension1
                            Case "BOOLEAN"
                                Dim boolParam1 As BoolParam
                                boolParam1 = MesParams.CreateBoolean(NameParam, True)
                                boolParam1.ValuateFromString(ValueParam.ToString)
                                MonParametreToUse = boolParam1
                        End Select

                    Else
                    End If

Boucle:
                Next

                'material

                If ic.Type = "PART" Then
                    Dim MaPartDoc As PartDocument = d
                    Dim MaPart As Part = MaPartDoc.Part
                    Dim MonMateriau As Material = Nothing
                    Dim c As String = DossierBase & "\" & INIProperties.GetString(GetEnv, "BibliothequeMaterial", "")
                    c = MonMainV3.GetAs(c)
                    Try
                        Dim MonDocMaterial = CATIA.Documents.Read(c)
                        MaPart.GetItem("CATMatManagerVBExt").GetMaterialOnPart(MaPart, MonMateriau)
                        MaPart.Update()
                    Catch ex As Exception
                        Dim newm As New MessageErreur("Une erreur s'est produite lors de l'insertion du matériau. Vérifier votre fichier Environnement.ini", Notifications.Wpf.NotificationType.Error)
                    End Try
                End If

            End If
        Catch ex_ As Exception
            Dim newm As New MessageErreur("Une erreur s'est produite ; impossible de continuer. Vérifier l'état de CATIA", Notifications.Wpf.NotificationType.Error)
        End Try


    End Sub

    Sub AppliqueMaterial(MaterialName As String, MyOwner As String, ic As ItemCatia)
        Dim d As Document = CATIA.Documents.Item(MyOwner)
        Dim MaPartDoc As PartDocument = d
        Dim MaPart As Part = MaPartDoc.Part
        Dim c As String = DossierBase & "\" & INIProperties.GetString(GetEnv, "BibliothequeMaterial", "")
        c = MonMainV3.GetAs(c)

        Dim MonDocMaterial
        Try
            MonDocMaterial = CATIA.Documents.Read(c)
        Catch ex As Exception
            Dim newm As New MessageErreur("Une erreur s'est produite lors de l'insertion du matériau. Vérifier votre fichier Environnement.ini", Notifications.Wpf.NotificationType.Error)
            Exit Sub
        End Try


        Dim MonMateriau As CATMat.Material = Nothing
        Dim MonItemMaterial As ItemMaterial = Nothing


        For Each m As ItemMaterial In ListMaterials
            If m.Name = MaterialName Then
                MonItemMaterial = m
                Exit For
            End If
        Next


        Dim p As Parameters = ic.ProductCATIA.UserRefProperties
        Try
            Dim test As String = p.Item(INIProperties.GetString(GetEnv, "ProprieteMATERIAL", "")).ValueAsString
        Catch ex As Exception
            FctionCATIA.AddParamatres(ic.Owner, ic)
        End Try

        If MonItemMaterial Is Nothing Then
            Exit Sub
        Else
            MonMateriau = MonDocMaterial.Families.Item(MonItemMaterial.Famille).Materials.Item(MonItemMaterial.Name)
            MaPart.GetItem("CATMatManagerVBExt").ApplyMaterialOnPart(MaPart, MonMateriau, 0)
            Try
                p.Item(INIProperties.GetString(GetEnv, "ProprieteMATERIAL", "")).Value = MonItemMaterial.Name
            Catch ex As Exception
                Dim merr As New MessageErreur("Une erreur s'est produite lors du remplissable du paramètre MATERIAL. Vérifier le fichier Environnments.ini", Notifications.Wpf.NotificationType.Error)
            End Try
        End If

    End Sub


    Sub getMATERIAL()

        Dim c As String = DossierBase & "\" & INIProperties.GetString(GetEnv, "BibliothequeMaterial", "")
        c = MonMainV3.GetAs(c)

        Dim MonDocMatariel
        Try
            MonDocMatariel = CATIA.Documents.Read(c)
            For Each f As MaterialFamily In MonDocMatariel.families
                Dim _f As New ItemFamille(f.Name)
                For Each m As CATMat.Material In f.Materials
                    Dim im As ItemMaterial = New ItemMaterial(m.Name, f.Name)
                    _f.Materials.Add(im)
                Next
            Next
        Catch e As Exception
            Dim newm As New MessageErreur("Une erreur s'est produite lors de la lecture du CATMatarial. Vérifier votre fichier Environnement.ini", Notifications.Wpf.NotificationType.Error)
        End Try





    End Sub

    Sub OpenFile(file As String)
        Try
            AppActivate("CATIA V5 - [" & CATIA.ActiveDocument.Name & "]")
            CATIA.Documents.Open(file)
        Catch ex As Exception
            Dim MsgErr As New MessageErreur("Impossible d'ouvrir le fichier " & file, Notifications.Wpf.NotificationType.Error)
        End Try
    End Sub

    Sub CreatefromBibliotheque(path As String)
        Try
            AppActivate("CATIA V5 - [" & CATIA.ActiveDocument.Name & "]")
            Dim d As Document = CATIA.Documents.NewFrom(path)
        Catch ex As Exception
            Dim MsgErr As New MessageErreur("Une erreur s'est produite. Vérifier que le fichier soit enregistré", Notifications.Wpf.NotificationType.Error)
        End Try
    End Sub
    Sub Createfrom(ic As ItemCatia)

        Try
            AppActivate("CATIA V5 - [" & CATIA.ActiveDocument.Name & "]")
            Dim d As Document = CATIA.Documents.NewFrom(ic.Doc.FullName)
            d.product.partNumber = ic.PartNumber & "_tmp"

        Catch ex As Exception
            Dim MsgErr As New MessageErreur("Une erreur s'est produite. Vérifier que le fichier soit enregistré", Notifications.Wpf.NotificationType.Error)
        End Try



    End Sub
    Sub SelectCATIA(ic As ItemCatia)
        Try
            AppActivate("CATIA V5 - [" & CATIA.ActiveDocument.Name & "]")
            If Not ic Is Nothing Then
                CATIA.ActiveDocument.Selection.Clear()
                If MaLangue = "Anglais" Then
                    If ic.Type = "PART" Then
                        CATIA.ActiveDocument.Selection.Search("Name='" & ic.PartNumber & "';all")
                    Else
                        CATIA.ActiveDocument.Selection.Search("Name=" & ic.PartNumber & ";all")
                    End If
                    CATIA.StartCommand("Reframe On")
                    CATIA.StartCommand("Center graph")
                ElseIf MaLangue = "Francais" Then
                    If ic.Type = "PART" Then
                        CATIA.ActiveDocument.Selection.Search("Nom='" & ic.PartNumber & "';tout")
                    Else
                        CATIA.ActiveDocument.Selection.Search("Nom=" & ic.PartNumber & ";tout")
                    End If
                    CATIA.StartCommand("Centrer sur")
                    CATIA.StartCommand("Centrer le graphe")
                End If

            End If
        Catch ex As Exception
            Dim MsgErr As New MessageErreur("Erreur lors de la séléction de l'élément", Notifications.Wpf.NotificationType.Error)
        End Try
    End Sub

#Region "TreeView"
    Sub SauvegardeTreeCatiaToTxtFile(FichierBaseTree As String, FichierTransform As String)

        If IO.File.Exists(FichierBaseTree) Then My.Computer.FileSystem.DeleteFile(FichierBaseTree)
        If IO.File.Exists(FichierTransform) Then My.Computer.FileSystem.DeleteFile(FichierTransform)


        MonActiveDoc.ExportData(FichierBaseTree, "txt")
        Dim Count As Integer = 0
        Dim fs As FileStream = Nothing
        fs = New FileStream(FichierTransform, FileMode.CreateNew)
        Using sw As StreamWriter = New StreamWriter(fs, Text.Encoding.GetEncoding("iso-8859-1"))
            Try
                Using sr As StreamReader = New StreamReader(FichierBaseTree, Text.Encoding.GetEncoding("iso-8859-1"))
                    Dim Line As String
                    Do
                        Line = sr.ReadLine()
                        If Count > 3 Then
                            If Count Mod 2 = 0 Then
                                Dim str As String = Line
                                str = Replace(str, "RootProduct : ", "")
                                str = Replace(str, "|- Product : ", "")
                                Dim NbEspaces As Integer = 0
                                Dim MesTabs As String = ""
                                Dim MonCaractere As String = ""
                                Do
                                    MonCaractere = str.Chars(NbEspaces)
                                    If MonCaractere = " " Then
                                        NbEspaces = NbEspaces + 1
                                    End If
                                Loop Until MonCaractere <> " "

                                If NbEspaces < 3 Then
                                    For i = 0 To NbEspaces - 2
                                        MesTabs = MesTabs & ControlChars.Tab
                                    Next
                                Else
                                    Dim Espaces As Integer = 0
                                    Espaces = ((NbEspaces - 2) / 6) + 0
                                    For i = 0 To Espaces
                                        MesTabs = MesTabs & ControlChars.Tab
                                    Next
                                End If
                                str = MesTabs & Strings.Right(str, Len(str) - NbEspaces)
                                str = Strings.Left(str, Len(str) - 1)
                                sw.WriteLine(str)
                            End If
                        End If
                        Count += 1
                    Loop Until Line Is Nothing
                    sr.Close()
                End Using
            Catch ex As Exception

            End Try
            sw.Close()
        End Using


        If IO.File.Exists(FichierBaseTree) Then My.Computer.FileSystem.DeleteFile(FichierBaseTree)


    End Sub


    Sub RemplirTreeViewFromTxtFile(ByVal file_name As String, ByVal trv As TreeView)

        Dim stream_reader As New System.IO.StreamReader(file_name, System.Text.Encoding.GetEncoding("iso-8859-1"))
        Dim file_contents As String = stream_reader.ReadToEnd()
        stream_reader.Close()
        file_contents = file_contents.Replace(vbLf, "")

        Const charCR As Char = CChar(vbCr)


        Dim lines() As String = file_contents.Split(charCR)

        Dim IT As New ItemTV(lines(0)) With {
            .Level = 0
        }
        ListItemTV.Add(IT)
        ITRacine = IT
        ChercheEnfants(lines, 0, IT)



    End Sub

    Sub ChercheEnfants(lines() As String, i As Integer, it As ItemTV)

        Const charTab As Char = CChar(vbTab)

        For _i = i + 1 To lines.Count - 1

            If lines(_i).Trim.Length > 0 Then
                Dim Lvl As Integer = lines(_i).Length - lines(_i).TrimStart(charTab).Length
                Dim _it As ItemTV = New ItemTV(Replace(lines(_i), charTab, "")) With {
                    .Level = Lvl
                }
                ListItemTV.Add(_it)

                If Lvl = it.Level + 1 = True Then
                    it.TVitem.Items.Add(_it.TVitem)
                    If VerifNextLevel(lines, _i, _it.Level) = True Then
                        ChercheEnfants(lines, _i, _it)
                    Else
                        _it.MajHeader("CATIAPart.ico")
                    End If
                Else
                    If _it.Level = it.Level Then
                        Exit For
                    End If
                End If
            End If


        Next


    End Sub
    Function MatchItemCatia(str As String, ligne As String) As ItemCatia

        Dim MonItemMatched As ItemCatia = Nothing

        Const charTab As Char = CChar(vbTab)


        Dim str_() As String = Strings.Split(str, " (")
        Dim PN As String = str_(0)
        For i = 1 To str_.Length - 2
            PN = PN & str_(i)
        Next


        For Each ic As ItemCatia In ListDocuments
            If ic.PartNumber.ToString = PN Then
                MonItemMatched = ic
                '   If ic.Level = Nothing Then
                ic.Level = ligne.Length - ligne.TrimStart(charTab).Length
                '  End If
                Exit For
            End If
        Next




        Return MonItemMatched


    End Function

    Sub GoListDocuments()
        NbElements = 1

        Dim s As Selection
        Try

            s = MonActiveDoc.Selection
            Dim query As String = Nothing
            query = "((CATProductSearch.Product + CATAsmSearch.Product) + CATPcsSearch.Product),all"
            s.Search(query)
        Catch ex As Exception
            Dim MsgErr As New MessageErreur("Erreur lors de la recherche d'éléments. Vérifier que CATIA soit en Anglais.", Notifications.Wpf.NotificationType.Warning)
            Exit Sub
        End Try

        ListPartNumber.Clear()

        For i = 1 To s.Count
            Try
                If Not ListPartNumber.Contains(s.Item(i).Value.partnumber) Then
                    ListPartNumber.Add(s.Item(i).Value.partnumber)
                    Bgw.ReportProgress(0, "ITEM [" & s.Item(i).Value.partnumber & "] récupéré")
                End If
            Catch ex As Exception
                'partnumber impossible à trouver
            End Try
        Next

        s.Clear()

        NbElements = ListPartNumber.Count + 1

        For Each d As Document In MonActiveDoc.Application.Documents
            If Right(d.FullName, 4) = "Part" Or Right(d.FullName, 7) = "Product" Then
                Dim p As Product = d.product
                If ListPartNumber.Contains(p.PartNumber) Then
                    Dim ic As ItemCatia = New ItemCatia(d)
                    If ic.FileName = MonActiveDoc.FullName Then
                        ic.Type = "RACINE"
                        ICRacine = ic
                    End If
                End If
            End If
        Next


    End Sub

    Sub GoPart()

        NbElements = 2
        Dim p As Product = MonActiveDoc.product
        Dim ic As ItemCatia = New ItemCatia(MonActiveDoc)
        ICRacine = ic

        ListPropertiesPartEnCours = GetAllItemProperties(p)
    End Sub

    Sub RemplirTVPart()
        MonMainV3.TVCatpart.Items.Clear()
        Dim p As PartDocument = MonActiveDoc
        If Not p Is Nothing Then
            Dim MaPart As Part = p.Part
            Dim N As New TreeViewItem
            Dim Lab As New Label
            Dim st As New StackPanel

            Dim Im As New Image With {
            .Width = 16,
            .Height = 16
        }

            Im.Source = New BitmapImage(New Uri(DossierImage & "CATIAPart.ico"))
            Lab.Content = Strings.Left(p.Name, p.Name.Length - 8)
            st.Orientation = Orientation.Horizontal
            st.Children.Add(Im)
            st.Children.Add(Lab)
            N.Header = st

            For Each item As Body In MaPart.Bodies
                If item.InBooleanOperation = False Then
                    Dim tn As New TreeViewItem
                    Dim tnn As New TextBlock
                    Dim tnlabel As New Label
                    Dim tnim As New Image With {
                        .Width = 16,
                        .Height = 16,
                        .Source = New BitmapImage(New Uri(DossierImage & "BODY.png"))
                        }
                    Dim tnstack As New StackPanel
                    tnlabel.Content = item.Name
                    tnstack.Orientation = Orientation.Horizontal
                    tnstack.Children.Add(tnim)
                    tnstack.Children.Add(tnlabel)
                    tn.Header = tnstack
                    N.Items.Add(tn)
                End If
            Next


            MonMainV3.TVCatpart.Items.Add(N)

        End If
        MonMainV3.TVCatpart.Items(0).IsExpanded = True
    End Sub
    Function GetAllItemProperties(p As Product) As List(Of String)

        Dim l As New List(Of String)
        For Each item As Parameter In p.UserRefProperties
            Dim par As New PropertiesPart(item)
            ListPropertiesPart.Add(par)
            l.Add(item.Name)
        Next
        Return l

    End Function

    Function VerifNextLevel(lines() As String, i As Integer, levelActuel As String) As Boolean
        VerifNextLevel = False
        Const charTab As Char = CChar(vbTab)
        Dim nextLevel As Integer

        nextLevel = lines(i + 1).Length - lines(i + 1).TrimStart(charTab).Length
        If nextLevel = levelActuel + 1 Then
            Return True
        Else
            Return False
        End If


    End Function


#End Region


    Sub CreerPlan(ic As ItemCatia)

        If Env = "[SPIRIT AEROSYSTEMS]" Then
            CreationPlanSpirit(ic)
        ElseIf Env = "[AIRBUS]" Then
            CreationPlanA320(ic)
        ElseIf Env = "[DASSAULT AVIATION]" Then
            CreerPlanDassault(ic)
        Else
            OuvrirPlan(ic)
        End If
    End Sub
    Sub OuvrirPlan(ic)

        Dim m As New MessageErreur("Plans indisponibles", Notifications.Wpf.NotificationType.Error)
    End Sub
    Sub CreerPlanDassault(ic As ItemCatia)
        '   On Error Resume Next


        Dim k = DialogPlanA320.ShowDialog()

        If k = Forms.DialogResult.OK Then
            Dim Draw As DrawingDocument = Nothing
            If DialogPlanA320.RadioA0.Checked = True Then
                Draw = CATIA.Documents.NewFrom(MonMainV3.GetAs(DossierBase & "\" & INIProperties.GetString(GetEnv, "Template2DPetitFormat", "")))
            ElseIf DialogPlanA320.RadioA2.Checked = True Then
                Draw = CATIA.Documents.NewFrom(MonMainV3.GetAs(DossierBase & "\" & INIProperties.GetString(GetEnv, "Template2DGrandFormat", "")))
            End If
            Dim MaSheet As DrawingSheet = Draw.Sheets.ActiveSheet
            Dim Da As String = ""
            Dim Jour As Integer = Today.Day
            Dim Mois As Integer = Today.Month
            Dim annee As String = Today.Year
            Da = Format(Jour, "00") & "-" & GetMonth(Mois) & "-" & Right(annee, 2)

            For Each T As DrawingText In MaSheet.Views.Item(2).Texts
                'DASSAULT        
                If T.Name = "TitleBlock_Data_Rights6" Then
                    T.Text = My.Settings.DRN.ToString
                End If
                If T.Name = "TitleBlock_Data_Rights7" Then
                    T.Text = Da
                End If
                If T.Name = "TitleBlock_Data_Rights8" Then
                    T.Text = My.Settings.TITLE1.ToString
                End If
                If T.Name = "TitleBlock_Data_Rights9" Then
                    T.Text = My.Settings.TITLE2.ToString
                End If

                Dim n As String = ic.PartNumber

                If n.Length > 11 Then
                    Dim Usine As String = Strings.Left(n, 2)
                    If T.Name = "TitleBlock_Data_Rights12" Then
                        T.Text = Usine
                    End If
                    If T.Name = "TitleBlock_Text_Rights13" Then
                        Select Case Usine
                            Case "MA"
                                T.Text = "Etablissement de Martignas"
                            Case "BZ"
                                T.Text = "Etablissement de Biaritz"
                        End Select
                    End If
                    If T.Name = "TitleBlock_Data_Rights13" Then
                        Dim f As String = Strings.Left(n, 5)
                        f = Strings.Right(f, 3)
                        T.Text = f
                    End If
                    If T.Name = "TitleBlock_Data_Rights14" Then
                        Dim f As String = Strings.Left(n, 7)
                        f = Strings.Right(f, 2)
                        T.Text = f
                    End If
                    If T.Name = "TitleBlock_Data_Rights15" Then
                        Dim f As String = Strings.Left(n, 8)
                        f = Strings.Right(f, 1)
                        T.Text = f
                    End If
                    If T.Name = "TitleBlock_Data_Rights16" Then
                        Dim f As String = Strings.Left(n, 10)
                        f = Strings.Right(f, 2)
                        T.Text = f
                    End If
                    If T.Name = "TitleBlock_Data_Rights17" Then
                        Dim f() As String = Strings.Split(n, "_")
                        Dim ss As String = f(0)
                        If ss.Length > 12 Then ss = Strings.Right(ss, ss.Length - 11)
                        T.Text = ss
                    End If
                Else
                    Dim m As New MessageErreur("Le PartNumber de l'élément CATIA n'est pas au format de Dassault. Impossible de générer le plan correctement.", Notifications.Wpf.NotificationType.Error)
                End If
                If T.Name = "TitleBlock_Data_Grille_date_18" Then
                    T.Text = GetMonth(Mois) & "-" & Right(annee, 2)
                End If

                If T.Name = "TitleBlock_Data_Tableau_1_0" Then
                    T.Text = ic.ProductCATIA.UserRefProperties.Item("NomPuls_Planche").Value
                End If

                If T.Name = "TitleBlock_Data_Tableau_3_0" Then
                    T.Text = ic.DescriptionRef
                End If
                If T.Name = "TitleBlock_Data_Tableau_4_0" Then
                    T.Text = ic.ProductCATIA.UserRefProperties.Item(INIProperties.GetString(GetEnv, "ProprieteMATERIAL", "")).Value
                End If

                If T.Name = "TitleBlock_Data_Tableau_5_0" Then
                    Select Case ic.Source
                        Case "Fabriqué"
                            T.Text = "FAB"
                        Case "Acheté"
                            T.Text = "ACH"
                        Case "Inconnu"
                            T.Text = "INC"
                    End Select

                End If
                If T.Name = "TitleBlock_Data_Tableau_6_0" Then
                    T.Text = ic.ProductCATIA.UserRefProperties.Item(INIProperties.GetString(GetEnv, "ProprieteMASSE", "")).Value & " Kg"
                End If
                If T.Name = "TitleBlock_Data_Tableau_7_0" Then
                    T.Text = ic.ProductCATIA.UserRefProperties.Item(INIProperties.GetString(GetEnv, "ProprieteTTS", "")).Value
                End If


            Next


            Dim i As Integer = 1
            Dim NFichier As String = ic.Doc.Path & "\" & ic.PartNumber & "-" & Format(i, "00") & ".CATDrawing"
            Do While IO.File.Exists(NFichier)
                NFichier = ic.Doc.Path & "\" & ic.PartNumber & "-" & Format(i, "00") & ".CATDrawing"
                i += 1
            Loop
            Draw.SaveAs(NFichier)
            Dim MsgErr As New MessageErreur("Le plan " & NFichier & " a été créé", Notifications.Wpf.NotificationType.Information)
        Else

            Exit Sub

        End If


        On Error GoTo 0
    End Sub
    Sub CreationPlanA320(ic As ItemCatia)

        On Error Resume Next


        Dim k = DialogPlanA320.ShowDialog()

        If k = Forms.DialogResult.OK Then
            Dim Draw As DrawingDocument

            If DialogPlanA320.RadioA0.Checked = True Then
                Draw = CATIA.Documents.NewFrom(MonMainV3.GetAs(DossierBase & "\" & INIProperties.GetString(GetEnv, "Template2DPetitFormat", "")))
            ElseIf DialogPlanA320.RadioA2.Checked = True Then
                Draw = CATIA.Documents.NewFrom(MonMainV3.GetAs(DossierBase & "\" & INIProperties.GetString(GetEnv, "Template2DGrandFormat", "")))
            End If

#Disable Warning BC42104 ' La variable 'Draw' est utilisée avant qu'une valeur ne lui ait été assignée. Une exception de référence null peut se produire au moment de l'exécution.
            Dim MaSheet As DrawingSheet = Draw.Sheets.ActiveSheet
#Enable Warning BC42104 ' La variable 'Draw' est utilisée avant qu'une valeur ne lui ait été assignée. Une exception de référence null peut se produire au moment de l'exécution.
            Dim Da As String = ""
            Dim Jour As Integer = Today.Day
            Dim Mois As Integer = Today.Month
            Dim annee As String = Today.Year
            Da = Format(Jour, "00") & "-" & GetMonth(Mois) & "-" & Right(annee, 2)

            For Each T As DrawingText In MaSheet.Views.Item(2).Texts
                'AIRBUS
                If T.Name = "TXT_Date_REV" Then
                    T.Text = GetMonth(Mois) & "-" & annee
                End If
                If T.Name = "AUKTbkText_JAT_ALL_DRN_DATE" Or T.Name = "AUKTbkText_JAT_ALL_CHKD_DATE" Or T.Name = "AUKTbkText_JAT_ALL_APPD_DATE" Then
                    T.Text = Da
                End If
                If T.Name = "AUKTbkText_JAT_ALL_DRAWING_NUMBER" Then
                    If ic.PartNumber.Length > 14 Then
                        T.Text = Left(ic.PartNumber, 14) & "000"
                    End If
                End If
                If T.Name = "AUKTbkText_JAT_AIF_STRUCTURE_DETAIL_L3" Then
                    T.Text = ic.PartNumber
                End If
                If T.Name = "AUKTbkText_JAT_AIF_STRUCTURE_DETAIL_L4" Then
                    Dim str As String = ""
                    If Not ic.ProductCATIA.UserRefProperties.Item("SYM").Value Is Nothing Then
                        str = ic.ProductCATIA.UserRefProperties.Item("SYM").Value.ToString
                    End If
                    If str <> "" Then
                        T.Text = str & " (SYM)"
                    End If
                End If
                If T.Name = "AUKTbkText_JAT_AIF_STRUCTURE_NBER_L3" Then
                    T.Text = 1
                End If

                If T.Name = "AUKTbkText_JAT_AIF_DESIGNATION_L3" Then
                    T.Text = ic.DescriptionRef
                End If
                If T.Name = "AUKTbkText_JAT_AIF_TOOL_MASS" Then
                    T.Text = ic.ProductCATIA.UserRefProperties.GetItem(INIProperties.GetString(GetEnv, "ProprieteMASSE", "")).valueasstring
                End If
                If T.Name = "AUKTbkText_JAT_ALL_PAINT_PROTECTION_L1" Then
                    T.Text = ic.ProductCATIA.UserRefProperties.GetItem(INIProperties.GetString(GetEnv, "ProprieteTTS", "")).valueasstring
                End If
                If T.Name = "AUKTbkText_JAT_ALL_PAINT_PROTECTION_L2" Then
                    If Strings.UCase(ic.ProductCATIA.UserRefProperties.GetItem(INIProperties.GetString(GetEnv, "ProprieteTTS", "")).valueasstring) = "PEINTURE" Then
                        T.Text = "RAL " & "XXXX"
                    End If
                End If
                If T.Name = "AUKTbkText_JAT_ALL_REMARK_L1" Then
                    T.Text = "BREAK SHARP EDGES / CASSER LES ARETES VIVES"
                End If
                If T.Name = "AUKTbkText_JAT_ALL_REMARK_L2" Then

                    T.Text = ic.ProductCATIA.UserRefProperties.GetItem(INIProperties.GetString(GetEnv, "ProprieteTTS", "")).valueasstring

                End If
                If T.Name = "AUKTbkText_JAT_AIF_AFFECTATION_MAT" Then

                    T.Text = ic.ProductCATIA.UserRefProperties.GetItem(INIProperties.GetString(GetEnv, "ProprieteMATERIAL", "")).valueasstring

                End If
                If T.Name = "AUKTbkText_JAT_ALL_SCALE" Then
                    If MaSheet.Views.Count > 2 Then
                        T.Text = MaSheet.Views.Item(3).Scale
                    End If
                    T.Text = "Not To SCALE"
                End If
                If T.Name = "AUKTbkText_JAT_ALL_GENERAL_TOLERANCES_L2" Then
                    T.Text = "DIN ISO 2768-m-K"
                End If
                If T.Name = "AUKTbkText_JAT_ALL_PLANT" Then
                    T.Text = My.Settings.PLANT.ToString
                End If
                If T.Name = "AUKTbkText_JAT_ALL_DRN" Or T.Name = "AUKTbkText_JAT_ALL_CHKD" Or T.Name = "AUKTbkText_JAT_ALL_APPD" Then
                    T.Text = My.Settings.DRN.ToString
                End If
                If T.Name = "AUKTbkText_JAT_ALL_TITLE_L1" Then
                    T.Text = My.Settings.TITLE1.ToString
                End If
                If T.Name = "AUKTbkText_JAT_ALL_TITLE_L2" Then
                    T.Text = My.Settings.TITLE2.ToString
                End If
                If T.Name = "AUKTbkText_JAT_AIF_PROGRAM_L3" Then
                    T.Text = My.Settings.PROGRAM.ToString
                End If

                If T.Name = "DA_SOC" Then
                    T.Text = My.Settings.DRN.ToString
                End If
            Next


            Dim i As Integer = 1
            Dim NFichier As String = ic.Doc.Path & "\" & ic.PartNumber & "-" & Format(i, "00") & ".CATDrawing"
            Do While IO.File.Exists(NFichier)
                NFichier = ic.Doc.Path & "\" & ic.PartNumber & "-" & Format(i, "00") & ".CATDrawing"
                i += 1
            Loop
            Draw.SaveAs(NFichier)
            Dim MsgErr As New MessageErreur("Le plan " & NFichier & " a été créé", Notifications.Wpf.NotificationType.Information)
        Else

            Exit Sub

        End If


        On Error GoTo 0

    End Sub
    Sub CreationPlanSpirit(ic As ItemCatia)

        On Error Resume Next

        Dim k = DialogPlanA320.ShowDialog()
        Dim textSheet As Boolean = False

        If k = Forms.DialogResult.OK Then
            Dim Draw As DrawingDocument

            If DialogPlanA320.RadioA0.Checked = True Then
                Draw = CATIA.Documents.NewFrom(MonMainV3.GetAs(DossierBase & "\" & INIProperties.GetString(GetEnv, "Template2DPetitFormat", "")))
                textSheet = False
            ElseIf DialogPlanA320.RadioA2.Checked = True Then
                Draw = CATIA.Documents.NewFrom(MonMainV3.GetAs(DossierBase & "\" & INIProperties.GetString(GetEnv, "Template2DPetitFormat", "")))
                textSheet = True

            End If

#Disable Warning BC42104 ' La variable 'Draw' est utilisée avant qu'une valeur ne lui ait été assignée. Une exception de référence null peut se produire au moment de l'exécution.
            Dim MaSheet As DrawingSheet = Draw.Sheets.ActiveSheet
#Enable Warning BC42104 ' La variable 'Draw' est utilisée avant qu'une valeur ne lui ait été assignée. Une exception de référence null peut se produire au moment de l'exécution.
            Dim Da As String = ""
            Dim Jour As Integer = Today.Day
            Dim Mois As Integer = Today.Month
            Dim annee As String = Today.Year
            Da = Format(Jour, "00") & "-" & Format(Mois, "00") & "-" & Right(annee, 2)

            For Each T As DrawingText In MaSheet.Views.Item(2).Texts
                If T.Name = "TextDRAWNBy" Then
                    T.Text = My.Settings.DRN.ToString
                End If
                If T.Name = "TextDRAWNByDATE" Then
                    T.Text = Da
                End If
                If T.Name = "TextMODEL" Then
                    T.Text = My.Settings.PROGRAM.ToString
                End If
                If T.Name = "TextSECTION" Then
                    T.Text = "15"
                End If
                If T.Name = "TextNumber" Then
                    Dim s As String = ic.PartNumber
                    Dim s_() As String = Strings.Split(s, "_")
                    s = s_(0)
                    s = Replace(s, "_", "")
                    T.Text = s
                End If
                If T.Name = "TextREV" Then
                    T.Text = ic.Revision
                End If
                If T.Name = "TextNamePart" Then
                    T.Text = ic.Nomenclature & Chr(10) & My.Settings.TITLE1.ToString
                End If
                If T.Name = "TEXTCageCode" Then
                    T.Text = My.Settings.CAGECODE.ToString
                End If
                If T.Name = "TextSHEET" Then
                    If textSheet = False Then
                        T.Text = ic.ProductCATIA.UserRefProperties.Item("DETAIL NUMBER").Value & " Of " & "_"
                    Else
                        T.Text = "SD" & ic.ProductCATIA.UserRefProperties.Item("DETAIL NUMBER").Value
                    End If
                End If
                If T.Name = "TextDetailNb" Then
                    T.Text = ic.ProductCATIA.UserRefProperties.Item("DETAIL NUMBER").Value
                End If
                If T.Name = "TextMAT-TTS-POIDS" Then
                    T.Text = Strings.UCase(ic.ProductCATIA.UserRefProperties.GetItem(INIProperties.GetString(GetEnv, "ProprieteMATERIAL", "")).valueasstring) & " - " & ic.DescriptionRef & Chr(10) & Strings.UCase(ic.ProductCATIA.UserRefProperties.GetItem("STOCK SIZE").valueasstring)
                End If
            Next

            Dim i As Integer = 1
            Dim s1 As String = ic.PartNumber
            Dim s_1() As String = Strings.Split(s1, "_")
            s1 = s_1(0)
            s1 = Replace(s1, "_", "")

            If textSheet = True Then
                s1 = s1 & "_SD" & ic.ProductCATIA.UserRefProperties.Item("DETAIL NUMBER").Value & "_" & ic.Revision
            Else
                s1 = s1 & "_SHT" & ic.ProductCATIA.UserRefProperties.Item("DETAIL NUMBER").Value & "_" & ic.Revision
            End If

            Dim NFichier As String = ic.Doc.Path & "\" & s1 & ".CATDrawing"
            Do While IO.File.Exists(NFichier)
                NFichier = ic.Doc.Path & "\" & s1 & "-" & Format(i, "00") & ".CATDrawing"
                i += 1
            Loop
            Draw.SaveAs(NFichier)
            Dim MsgErr As New MessageErreur("Le plan " & NFichier & " a été créé", Notifications.Wpf.NotificationType.Information)
        Else

            Exit Sub

        End If


        On Error GoTo 0

    End Sub

    Function GetMonth(Mois As Integer) As String
        Select Case Mois
            Case 1
                Return "JAN"
            Case 2
                Return "FEB"
            Case 3
                Return "MAR"
            Case 4
                Return "APR"
            Case 5
                Return "MAY"
            Case 6
                Return "JUN"
            Case 7
                Return "JUL"
            Case 8
                Return "AUG"
            Case 9
                Return "SEP"
            Case 10
                Return "OCT"
            Case 11
                Return "NOV"
            Case 12
                Return "DEC"
            Case Else
                Return Nothing
        End Select
    End Function

    Sub MajCartoucheDassault()



        Dim d As DrawingDocument = CATIA.ActiveDocument
        Dim s As DrawingSheet = d.Sheets.ActiveSheet

        If s.Views.Count < 3 Then
            Dim m As New MessageErreur("Aucune vue n'est détectée pour pouvoir créer les liens avec le 3D. Créer une vue réessayer.", Notifications.Wpf.NotificationType.Error)
            Exit Sub
        End If

        DeleteBOMExistante(s.Views.Item("Background View"))

        Dim maVue As DrawingView = s.Views.Item(3)
        Dim MonDoc As Document = maVue.GenerativeBehavior.Document.Parent
        Dim MonIC As ItemCatia = Nothing


        For Each ic As ItemCatia In ListDocuments
            If ic.FileName = MonDoc.FullName Then
                MonIC = ic
                Exit For
            End If
        Next

        If MonIC Is Nothing Then
            CATIA.DisplayFileAlerts = False
            CATIA.Documents.Open(MonDoc.FullName)
            CATIA.DisplayFileAlerts = True
        End If

        CreerLignesBOM(s.Views.Item("Background View"), s.Views.Item("Background View").Texts, MonIC, 1, s.PaperSize, Nothing)
        CreerTitreBOM(s.Views.Item("Background View"), s.Views.Item("Background View").Texts, s.PaperSize)


        Dim Da As String = ""
        Dim Jour As Integer = Today.Day
        Dim Mois As Integer = Today.Month
        Dim annee As String = Today.Year
        Da = Format(Jour, "00") & "-" & GetMonth(Mois) & "-" & Right(annee, 2)

        For Each T As DrawingText In s.Views.Item(2).Texts
            'DASSAULT        
            If T.Name = "TitleBlock_Data_Rights6" Then
                T.Text = My.Settings.DRN.ToString
            End If
            If T.Name = "TitleBlock_Data_Rights7" Then
                T.Text = Da
            End If
            If T.Name = "TitleBlock_Data_Rights8" Then
                T.Text = My.Settings.TITLE1.ToString
            End If
            If T.Name = "TitleBlock_Data_Rights9" Then
                T.Text = My.Settings.TITLE2.ToString
            End If

            Dim n As String = MonIC.PartNumber

            If n.Length > 11 Then
                Dim Usine As String = Strings.Left(n, 2)
                If T.Name = "TitleBlock_Data_Rights12" Then
                    T.Text = Usine
                End If
                If T.Name = "TitleBlock_Text_Rights13" Then
                    Select Case Usine
                        Case "MA"
                            T.Text = "Etablissement de Martignas"
                        Case "BZ"
                            T.Text = "Etablissement de Biaritz"
                    End Select
                End If
                If T.Name = "TitleBlock_Data_Rights13" Then
                    Dim f As String = Strings.Left(n, 5)
                    f = Strings.Right(f, 3)
                    T.Text = f
                End If
                If T.Name = "TitleBlock_Data_Rights14" Then
                    Dim f As String = Strings.Left(n, 7)
                    f = Strings.Right(f, 2)
                    T.Text = f
                End If
                If T.Name = "TitleBlock_Data_Rights15" Then
                    Dim f As String = Strings.Left(n, 8)
                    f = Strings.Right(f, 1)
                    T.Text = f
                End If
                If T.Name = "TitleBlock_Data_Rights16" Then
                    Dim f As String = Strings.Left(n, 10)
                    f = Strings.Right(f, 2)
                    T.Text = f
                End If
                If T.Name = "TitleBlock_Data_Rights17" Then
                    Dim f() As String = Strings.Split(n, "_")
                    Dim ss As String = f(0)
                    ss = Strings.Right(ss, ss.Length - 11)
                    T.Text = ss
                End If
            Else
                Dim m As New MessageErreur("Le PartNumber de l'élément CATIA n'est pas au format de Dassault. Impossible de générer le plan correctement.", Notifications.Wpf.NotificationType.Error)
            End If
            If T.Name = "TitleBlock_Data_Grille_date_18" Then
                T.Text = GetMonth(Mois) & "-" & Right(annee, 2)
            End If

            If T.Name = "TitleBlock_Data_Tableau_1_0" Then
                T.Text = MonIC.ProductCATIA.UserRefProperties.Item("NomPuls_Planche").Value
            End If

            If T.Name = "TitleBlock_Data_Tableau_3_0" Then
                T.Text = MonIC.DescriptionRef
            End If
            If T.Name = "TitleBlock_Data_Tableau_4_0" Then
                T.Text = MonIC.ProductCATIA.UserRefProperties.Item(INIProperties.GetString(GetEnv, "ProprieteMATERIAL", "")).Value
            End If

            If T.Name = "TitleBlock_Data_Tableau_5_0" Then
                Select Case MonIC.Source
                    Case "Fabriqué"
                        T.Text = "FAB"
                    Case "Acheté"
                        T.Text = "ACH"
                    Case "Inconnu"
                        T.Text = "INC"
                End Select

            End If
            If T.Name = "TitleBlock_Data_Tableau_6_0" Then
                T.Text = MonIC.ProductCATIA.UserRefProperties.Item(INIProperties.GetString(GetEnv, "ProprieteMASSE", "")).Value & " Kg"
            End If
            If T.Name = "TitleBlock_Data_Tableau_7_0" Then
                T.Text = MonIC.ProductCATIA.UserRefProperties.Item(INIProperties.GetString(GetEnv, "ProprieteTTS", "")).Value
            End If


        Next


    End Sub


#Region "NOMECNALTURE 2D"




    Dim ListIC As New List(Of ItemCatia)
    Dim BoolMessWarningComponents As Boolean = False
    Sub GetBomDetails(Product As String, multiple As Integer, level As Integer)


        Dim ListComponents As New List(Of String)
        Dim FichierNomenclature As String = My.Computer.FileSystem.SpecialDirectories.Temp & "\BOM_.txt"
        Const charTab As Char = CChar(vbTab)
        Using sr As StreamReader = New StreamReader(FichierNomenclature, Encoding.GetEncoding("iso-8859-1"))
            Dim BoolStart As Boolean = False
            While Not sr.EndOfStream
                Dim line As String = sr.ReadLine
                If BoolStart = True Then
                    Dim l_() As String = Strings.Split(line, charTab)
                    If l_.Count = 2 Then
                        Dim monPN As String = l_(0)
                        Dim quantite As String = l_(1)
                        Dim isComponents As Boolean = True
                        For Each ic As ItemCatia In ListDocuments
                            If ic.PartNumber = monPN Then
                                Dim level_ As Integer
                                If level = 0 Then
                                    level_ = 1
                                Else
                                    level_ = level
                                End If
                                ic.Qte = Convert.ToInt32(ic.Qte) + (Convert.ToInt32(quantite) * multiple * level_)
                                ListIC.Add(ic)
                                isComponents = False
                                Exit For
                            End If
                        Next
                        If isComponents = True Then
                            GetBomDetails(monPN, Convert.ToInt32(quantite), level + 1)
                            BoolMessWarningComponents = True
                        End If
                    Else
                        BoolStart = False
                    End If
                End If

                If line = "[" & Product & "]" Then BoolStart = True


            End While
            sr.Close()
        End Using
    End Sub

    Sub GoNomenclature(Product As String)

        BoolMessWarningComponents = False
        For Each item In ListDocuments
            item.Qte = Nothing
        Next
        ListIC.Clear()
        If Product = "Ensemble des éléments" Then Product = "ALL-BOM-APPKD"
        GetBomDetails(Product, 1, 0)


        For Each item In ListDocuments
            If ListIC.Contains(item) Then
                item.Visible = True
            Else
                item.Visible = False
            End If
        Next

        If BoolMessWarningComponents = True Then
            Dim m As New MessageErreur("Des composants ont été détéctés lors de la génération de la nomenclature. Assurez-vous d'avoir un nom différent pour chaque composant afin de ne pas fausser le calcul des quantités de pièces", Notifications.Wpf.NotificationType.Warning)
        End If


    End Sub


    Sub RecursComponents(l As List(Of String))

        For Each item In l
            Dim ListComponents As New List(Of String)
            Const charTab As Char = CChar(vbTab)
            Dim FichierNomenclature As String = My.Computer.FileSystem.SpecialDirectories.Temp & "\BOM_.txt"

            Using sr As StreamReader = New StreamReader(FichierNomenclature, Encoding.GetEncoding("iso-8859-1"))
                Dim BoolStart As Boolean = False
                While Not sr.EndOfStream
                    Dim line As String = sr.ReadLine
                    If line Like "*" & charTab & "*" Then
                        Dim l_() As String = Strings.Split(line, charTab)
                        Dim monPN As String = l_(0)
                        Dim quantite As String = l_(1)
                        Dim CompOK As Boolean = True
                        For Each ic As ItemCatia In ListDocuments
                            If ic.PartNumber = monPN Then
                                ic.Qte = Convert.ToInt32(ic.Qte) + Convert.ToInt32(quantite)
                                ListIC.Add(ic)
                                CompOK = False
                                Exit For
                            End If
                        Next
                        If CompOK = True Then '<= c'est un components
                            ListComponents.Add(monPN)
                            '   MsgBox(monPN)
                            'composant non pris en compte - à modifier
                        End If
                    End If
                End While
                sr.Close()
            End Using
            RecursComponents(ListComponents)

        Next


    End Sub

    Sub CreerNomenclature2D(listIc As List(Of ItemCatia), EachElements As Boolean)



        listIc = listIc.OrderBy(Function(x) x.Nomenclature).ToList()
        Try
            If CheckSi2D(CATIA.ActiveDocument) = False Then
                Dim MsgErr As New MessageErreur("Un Drawing doit être ouvert pour pouvoir générer la nomenclature", Notifications.Wpf.NotificationType.Error)
                Exit Sub
            End If
        Catch ex As Exception
            Dim MsgErr As New MessageErreur("Impossible de générer une nomenclature", Notifications.Wpf.NotificationType.Error)
            Exit Sub
        End Try


        Dim D As DrawingDocument = CATIA.ActiveDocument
        Dim S As DrawingSheet = D.Sheets.ActiveSheet
        Dim V As DrawingView = S.Views.Item("Background View")
        Dim MesTexts As DrawingTexts = V.Texts
        Dim Type As Integer = S.PaperSize




        DeleteBOMExistante(MesTexts)
        If listIc.Count > 0 Then
            CreerTitreBOM(V, MesTexts, Type)
            Dim i As Integer = 1
            For Each ic As ItemCatia In listIc
                If EachElements = True Then
                    If ic.Type = "RACINE" Or ic.Type = "PRODUCT" Then
                    Else
                        CreerLignesBOM(V, MesTexts, ic, i, Type, Nothing)
                        i += 1
                    End If
                Else
                    If ic.Type = "RACINE" Then
                    Else
                        CreerLignesBOM(V, MesTexts, ic, i, Type, Nothing)
                        i += 1
                    End If
                End If

            Next
            Dim MsgErr As New MessageErreur("La nomenclature a été générée avec succès", Notifications.Wpf.NotificationType.Information)
        Else
            Dim MsgErr As New MessageErreur("Aucun item dans le tableau : la nomenclature a été supprimée avec succès", Notifications.Wpf.NotificationType.Information)
        End If
        S.Views.Item(1).Activate()
    End Sub

    Sub DeleteBOMExistante(v As DrawingView)
        If Env = "[AIRBUS]" Then
            Dim s As Selection = CATIA.ActiveDocument.Selection
            s.Clear()
            s.Search("Name=Nomenclature*_*;all")


            If s.Count > 0 Then
                s.Delete()
                s.Clear()
            End If
        End If
        If Env = "[DASSAULT AVIATION]" Then
            Dim s As Selection = CATIA.ActiveDocument.Selection
            s.Clear()
            s.Search("Name=TitleBlock*Tableau*;all") 'TitleBlock_Text_Tableau_1_0
            If s.Count > 0 Then
                s.Delete()
                s.Clear()
            End If
        End If

    End Sub
    Sub CreerTitreBOM(v As DrawingView, mestexts As DrawingTexts, TYPEPlan As Integer)

        If Env = "[AIRBUS]" Then
            Dim X, Y As Integer
            Select Case TYPEPlan
                Case 2
                    X = 979
                    Y = 151
                Case 3
                    X = 384
                    Y = 151
                Case 4
                    X = 384
                    Y = 151
                Case 5
                    X = 210
                    Y = 151
            End Select

            '   CATIA.StartWorkbench("CS0WKS")
            Dim F As Factory2D = v.Factory2D
            Dim L As Line2D = F.CreateLine(X, Y, X + 200, Y)
            L.Name = "NomenclatureLine_0_1_1"
            L = F.CreateLine(X, Y - 5, X, Y)
            L.Name = "NomenclatureLine_0_1_2"
            L = F.CreateLine(X + 8, Y - 5, X + 8, Y)
            L.Name = "NomenclatureLine_0_1_3"
            L = F.CreateLine(X + 18, Y - 5, X + 18, Y)
            L.Name = "NomenclatureLine_0_1_4"
            L = F.CreateLine(X + 27, Y - 5, X + 27, Y)
            L.Name = "NomenclatureLine_0_1_5"
            L = F.CreateLine(X + 78, Y - 5, X + 78, Y)
            L.Name = "NomenclatureLine_0_1_11"
            L = F.CreateLine(X + 114, Y - 5, X + 114, Y)
            L.Name = "NomenclatureLine_0_1_6"
            L = F.CreateLine(X + 130, Y - 5, X + 130, Y)
            L.Name = "NomenclatureLine_0_1_7"
            L = F.CreateLine(X + 140, Y - 5, X + 140, Y)
            L.Name = "NomenclatureLine_0_1_8"
            L = F.CreateLine(X + 160, Y - 5, X + 160, Y)
            L.Name = "NomenclatureLine_0_1_9"
            L = F.CreateLine(X, Y + 5, X, Y + 5)
            L.Name = "NomenclatureLine_0_1_0"


            Dim T As DrawingText = mestexts.Add("REP", X + 1, Y - 1)
            T.AnchorPosition = CatTextAnchorPosition.catTopLeft
            T.SetFontSize(0, 0, 2)
            T.Name = "NomenclatureTitle_repere_1"

            T = mestexts.Add("Pl", X + 9, Y - 1)
            T.AnchorPosition = CatTextAnchorPosition.catTopLeft
            T.SetFontSize(0, 0, 2)
            T.Name = "NomenclatureTitle_planche_1"

            T = mestexts.Add("Qté", X + 19, Y - 1)
            T.AnchorPosition = CatTextAnchorPosition.catTopLeft
            T.SetFontSize(0, 0, 2)
            T.Name = "NomenclatureTitle_quantite_1"

            T = mestexts.Add("Désignation", X + 28, Y - 1)
            T.AnchorPosition = CatTextAnchorPosition.catTopLeft
            T.SetFontSize(0, 0, 2)
            T.Name = "NomenclatureTitle_Designation_1"

            T = mestexts.Add("PartName", X + 79, Y - 1)
            T.AnchorPosition = CatTextAnchorPosition.catTopLeft
            T.SetFontSize(0, 0, 2)
            T.Name = "NomenclatureTitle_PartNumber_1"

            T = mestexts.Add("Matière", X + 115, Y - 1)
            T.AnchorPosition = CatTextAnchorPosition.catTopLeft
            T.SetFontSize(0, 0, 2)
            T.Name = "NomenclatureTitle_matiere_1"

            T = mestexts.Add("Etat", X + 131, Y - 1)
            T.AnchorPosition = CatTextAnchorPosition.catTopLeft
            T.SetFontSize(0, 0, 2)
            T.Name = "NomenclatureTitle_etat_1"

            T = mestexts.Add("Dim. Brutes", X + 141, Y - 1)
            T.AnchorPosition = CatTextAnchorPosition.catTopLeft
            T.SetFontSize(0, 0, 2)
            T.Name = "NomenclatureTitle_dimensions_brutes_1"

            T = mestexts.Add("Observations", X + 161, Y - 1)
            T.AnchorPosition = CatTextAnchorPosition.catTopLeft
            T.SetFontSize(0, 0, 2)
            T.Name = "NomenclatureTitle_observations_1"

            T = mestexts.Add(ICRacine.PartNumber.ToString, 0, 0)
            T.AnchorPosition = CatTextAnchorPosition.catTopLeft
            T.SetFontSize(0, 0, 2)
            T.Name = "NomenclatureText_NomProduitBase"


            Dim sel As Selection = CATIA.ActiveDocument.Selection
            sel.Add(T)
            sel.VisProperties.SetShow(1)
            sel.Clear()
        ElseIf Env = "[DASSAULT AVIATION" Then
            Dim X, Y As Integer
            Select Case TYPEPlan
                Case 2
                    X = 979
                    Y = 151
                Case 3
                    X = 384
                    Y = 151
                Case 4
                    X = 384
                    Y = 151
                Case 5
                    X = 210
                    Y = 151
            End Select


        Else
        End If

        If Env = "[DASSAULT AVIATION]" Then
            'rien à faire
        End If


    End Sub
    Sub CreerLignesBOM(v As DrawingView, mestexts As DrawingTexts, ic As ItemCatia, i As Integer, TYPEPlan As String, LineB As LineBOM)

        If Env = "[AIRBUS]" Then
            Dim X, Y As Integer
            Select Case TYPEPlan
                Case 2
                    X = 979
                    Y = 151
                Case 3
                    X = 384
                    Y = 151
                Case 4
                    X = 384
                    Y = 151
                Case 5
                    X = 210
                    Y = 151
            End Select

            Dim F = v.Factory2D
            If F Is Nothing Then
                Dim m As New MessageErreur("Une erreur interne à l'application s'est produite liée aux références CATIA", Notifications.Wpf.NotificationType.Error)
            Else
                Dim L As Line2D
                L = F.CreateLine(X, Y + (5 * i), X + 200, Y + (5 * i))
                L.Name = "NomenclatureLine_" & i & "_1_1"
                L = F.CreateLine(X, Y + (5 * (i - 1)), X, Y + (5 * i))
                L.Name = "NomenclatureLine_" & i & "_1_2"
                L = F.CreateLine(X + 8, Y + (5 * (i - 1)), X + 8, Y + (5 * i))
                L.Name = "NomenclatureLine_" & i & "_1_3"
                L = F.CreateLine(X + 18, Y + (5 * (i - 1)), X + 18, Y + (5 * i))
                L.Name = "NomenclatureLine_" & i & "_1_4"
                L = F.CreateLine(X + 27, Y + (5 * (i - 1)), X + 27, Y + (5 * i))
                L.Name = "NomenclatureLine_" & i & "_1_5"
                L = F.CreateLine(X + 78, Y + (5 * (i - 1)), X + 78, Y + (5 * i))
                L.Name = "NomenclatureLine_" & i & "_1_11"
                L = F.CreateLine(X + 114, Y + (5 * (i - 1)), X + 114, Y + (5 * i))
                L.Name = "NomenclatureLine_" & i & "_1_6"
                L = F.CreateLine(X + 130, Y + (5 * (i - 1)), X + 130, Y + (5 * i))
                L.Name = "NomenclatureLine_" & i & "_1_7"
                L = F.CreateLine(X + 140, Y + (5 * (i - 1)), X + 140, Y + (5 * i))
                L.Name = "NomenclatureLine_" & i & "_1_8"
                L = F.CreateLine(X + 160, Y + (5 * (i - 1)), X + 160, Y + (5 * i))
                L.Name = "NomenclatureLine_" & i & "_1_9"

                If LineB Is Nothing Then
                    Dim T As DrawingText = mestexts.Add(CheckSiTextVide(ic.Nomenclature), X, Y + (5 * i))
                    T.AnchorPosition = CatTextAnchorPosition.catTopLeft
                    T.SetFontSize(0, 0, 2)
                    T.Name = "NomenclatureText_repere_" & i

                    T = mestexts.Add(".", X + 9, Y + (5 * i))
                    T.AnchorPosition = CatTextAnchorPosition.catTopLeft
                    T.SetFontSize(0, 0, 2)
                    T.Name = "NomenclatureText_planche_" & i

                    T = mestexts.Add(CheckSiTextVide(ic.Qte), X + 19, Y + (5 * i))
                    T.AnchorPosition = CatTextAnchorPosition.catTopLeft
                    T.SetFontSize(0, 0, 2)
                    T.Name = "NomenclatureText_quantite_" & i

                    Dim strDescription As String = ic.ProductCATIA.DescriptionRef
                    If strDescription.Contains(vbCrLf) Then
                        T = mestexts.Add(CheckSiTextVide(ic.ProductCATIA.DescriptionRef), X + 28, Y + 0 + (5 * i))
                        T.AnchorPosition = CatTextAnchorPosition.catTopLeft
                        T.SetFontSize(0, 0, 1.5)
                        T.Name = "NomenclatureText_Designation_" & i
                    Else
                        T = mestexts.Add(CheckSiTextVide(ic.DescriptionRef), X + 28, Y + (5 * i))
                        T.AnchorPosition = CatTextAnchorPosition.catTopLeft
                        T.SetFontSize(0, 0, 2)
                        T.Name = "NomenclatureText_Designation_" & i
                    End If


                    T = mestexts.Add(CheckSiTextVide(ic.PartNumber), X + 79, Y + (5 * i))
                    T.AnchorPosition = CatTextAnchorPosition.catTopLeft
                    T.SetFontSize(0, 0, 2)
                    T.Name = "NomenclatureText_PartNumber_" & i

                    T = mestexts.Add(CheckSiTextVide(ic.ProductCATIA.UserRefProperties.Item(INIProperties.GetString(GetEnv, "ProprieteMATERIAL", "")).Value), X + 115, Y + (5 * i))
                    T.AnchorPosition = CatTextAnchorPosition.catTopLeft
                    T.SetFontSize(0, 0, 2)
                    T.Name = "NomenclatureText_matiere_" & i

                    Dim MaSourceRe As String = ic.Source
                    Select Case MaSourceRe
                        Case "Inconnu"
                            MaSourceRe = ""
                        Case "Fabriqué"
                            MaSourceRe = "FAB"
                        Case "Acheté"
                            MaSourceRe = "ACHAT"
                    End Select
                    T = mestexts.Add(CheckSiTextVide(MaSourceRe), X + 131, Y + (5 * i))
                    T.AnchorPosition = CatTextAnchorPosition.catTopLeft
                    T.SetFontSize(0, 0, 2)
                    T.Name = "NomenclatureText_etat_" & i

                    T = mestexts.Add(".", X + 141, Y + (5 * i))
                    T.AnchorPosition = CatTextAnchorPosition.catTopLeft
                    T.SetFontSize(0, 0, 2)
                    T.Name = "NomenclatureText_dimensions_brutes_" & i

                    T = mestexts.Add(CheckSiTextVide(ic.ProductCATIA.UserRefProperties.Item("OBSERVATIONS").Value), X + 161, Y + (5 * i))
                    T.AnchorPosition = CatTextAnchorPosition.catTopLeft
                    T.SetFontSize(0, 0, 2)
                    T.Name = "NomenclatureText_observations_" & i


                    T = mestexts.Add(ic.PartNumber, 385, 155 + (5 * i) - 5)
                    T.SetFontSize(0, 0, 2)
                    T.Name = "NomenclatureText_LinkPartNumber_" & i
                End If

            End If
        End If


        If Env = "[DASSAULT AVIATION]" Then
            Dim X, Y As Integer
            Select Case TYPEPlan
                Case 2
                    X = 893
                    Y = 128
                Case 3 'OK
                    X = 299
                    Y = 128
                Case 4 'OK
                    X = 299
                    Y = 128
                Case 5
                    X = 210
                    Y = 151
            End Select

            Dim F = v.Factory2D
            If F Is Nothing Then
                Dim m As New MessageErreur("Une erreur interne à l'application s'est produite liée aux références CATIA", Notifications.Wpf.NotificationType.Error)
            Else
                Dim L As Line2D
                L = F.CreateLine(X, Y + (6 * (i - 1)), X + 285, Y + (6 * (i - 1)))
                L.Name = "TitleBlock_Line_Row_Tableau_1" & i
                L = F.CreateLine(X, Y + (6 * (i - 1)), X, Y + (6 * (i)))
                L.Name = "TitleBlock_Line_Col_Tableau_1" & i - 1
                L = F.CreateLine(X + 12, Y + (6 * (i - 1)), X + 12, Y + (6 * (i)))
                L.Name = "TitleBlock_Line_Col_Tableau_2" & i - 1
                L = F.CreateLine(X + 37, Y + (6 * (i - 1)), X + 37, Y + (6 * (i)))
                L.Name = "TitleBlock_Line_Col_Tableau_3" & i - 1
                L = F.CreateLine(X + 49, Y + (6 * (i - 1)), X + 49, Y + (6 * (i)))
                L.Name = "TitleBlock_Line_Col_Tableau_4" & i - 1
                L = F.CreateLine(X + 144, Y + (6 * (i - 1)), X + 144, Y + (6 * (i)))
                L.Name = "TitleBlock_Line_Col_Tableau_5" & i - 1
                L = F.CreateLine(X + 174, Y + (6 * (i - 1)), X + 174, Y + (6 * (i)))
                L.Name = "TitleBlock_Line_Col_Tableau_6" & i - 1
                L = F.CreateLine(X + 190, Y + (6 * (i - 1)), X + 190, Y + (6 * (i)))
                L.Name = "TitleBlock_Line_Col_Tableau_7" & i - 1
                L = F.CreateLine(X + 220, Y + (6 * (i - 1)), X + 220, Y + (6 * (i)))
                L.Name = "TitleBlock_Line_Col_Tableau_8" & i - 1
                L = F.CreateLine(X + 250, Y + (6 * (i - 1)), X + 250, Y + (6 * (i)))
                L.Name = "TitleBlock_Line_Col_Tableau_9" & i - 1

                If LineB Is Nothing Then
                    Dim T As DrawingText = mestexts.Add(Format(i, "00"), X + 6, Y + 2.5 + (6 * (i - 1)))
                    T.AnchorPosition = CatTextAnchorPosition.catMiddleCenter
                    T.TextProperties.Justification = CatJustification.catCenter
                    T.SetParameterOnSubString(CatTextProperty.catBold, 0, 0, 1)
                    T.SetParameterOnSubString(CatTextProperty.catItalic, 0, 0, 1)
                    T.SetFontSize(0, 0, 4)
                    T.Name = "TitleBlock_Text_Tableau_1_" & i - 1

                    T = mestexts.Add(CheckSiTextVide(ic.ProductCATIA.UserRefProperties.Item("NomPuls_Planche").Value), X + 24.5, Y + 2.5 + (6 * (i - 1)))
                    T.WrappingWidth = 25
                    T.AnchorPosition = CatTextAnchorPosition.catMiddleCenter
                    T.TextProperties.Justification = CatJustification.catCenter
                    T.SetFontSize(0, 0, 2.5)
                    T.Name = "TitleBlock_Data_Tableau_1_" & i - 1

                    T = mestexts.Add(CheckSiTextVide(ic.Qte), X + 43, Y + 2.5 + (6 * (i - 1)))
                    T.WrappingWidth = 12
                    T.AnchorPosition = CatTextAnchorPosition.catMiddleCenter
                    T.TextProperties.Justification = CatJustification.catCenter
                    T.SetFontSize(0, 0, 2.5)
                    T.Name = "TitleBlock_Data_Tableau_2_" & i - 1

                    Dim strDescription As String = ic.ProductCATIA.DescriptionRef
                    If strDescription.Contains(vbCrLf) Then
                        T = mestexts.Add(CheckSiTextVide(ic.ProductCATIA.UserRefProperties.Item("NomPuls_Designation").Value), X + 97, Y + 2.5 + (6 * (i - 1)))
                        T.WrappingWidth = 93
                        T.AnchorPosition = CatTextAnchorPosition.catMiddleCenter
                        T.TextProperties.Justification = CatJustification.catCenter
                        T.SetFontSize(0, 0, 1.5)
                        T.Name = "TitleBlock_Data_Tableau_3_" & i - 1
                    Else
                        T = mestexts.Add(CheckSiTextVide(ic.ProductCATIA.UserRefProperties.Item("NomPuls_Designation").Value), X + 97, Y + 2.5 + (6 * (i - 1)))
                        T.WrappingWidth = 93
                        T.AnchorPosition = CatTextAnchorPosition.catMiddleCenter
                        T.TextProperties.Justification = CatJustification.catCenter
                        T.SetFontSize(0, 0, 2.5)
                        T.Name = "TitleBlock_Data_Tableau_3_" & i - 1
                    End If


                    T = mestexts.Add(CheckSiTextVide(ic.ProductCATIA.UserRefProperties.Item(INIProperties.GetString(GetEnv, "ProprieteMATERIAL", "")).Value), X + 159, Y + 2.5 + (6 * (i - 1)))
                    T.WrappingWidth = 30
                    T.AnchorPosition = CatTextAnchorPosition.catMiddleCenter
                    T.TextProperties.Justification = CatJustification.catCenter
                    T.SetFontSize(0, 0, 2.5)
                    T.Name = "TitleBlock_Data_Tableau_4_" & i - 1



                    Dim MaSourceRe As String = ic.Source
                    Select Case MaSourceRe
                        Case "Inconnu"
                            MaSourceRe = "-"
                        Case "Fabriqué"
                            MaSourceRe = "FAB"
                        Case "Acheté"
                            MaSourceRe = "ACHAT"
                    End Select
                    T = mestexts.Add(CheckSiTextVide(MaSourceRe), X + 182, Y + 2.5 + (6 * (i - 1)))
                    T.WrappingWidth = 16
                    T.AnchorPosition = CatTextAnchorPosition.catMiddleCenter
                    T.TextProperties.Justification = CatJustification.catCenter
                    T.SetFontSize(0, 0, 2.5)
                    T.Name = "TitleBlock_Data_Tableau_5_" & i - 1

                    Dim s As String = ic.ProductCATIA.UserRefProperties.Item("NomPuls_Dim_Brutes").Value
                    If s = "" And ic.ProductCATIA.UserRefProperties.Item(INIProperties.GetString(GetEnv, "ProprieteMASSE", "")).Value <> "" Then s = ic.ProductCATIA.UserRefProperties.Item(INIProperties.GetString(GetEnv, "ProprieteMASSE", "")).Value & " Kg"
                    T = mestexts.Add(CheckSiTextVide(s), X + 205, Y + 2.5 + (6 * (i - 1))) 'dim brutes
                    T.WrappingWidth = 30
                    T.AnchorPosition = CatTextAnchorPosition.catMiddleCenter
                    T.TextProperties.Justification = CatJustification.catCenter
                    T.SetFontSize(0, 0, 2.5)
                    T.Name = "TitleBlock_Data_Tableau_6_" & i - 1

                    T = mestexts.Add(CheckSiTextVide(ic.ProductCATIA.UserRefProperties.Item(INIProperties.GetString(GetEnv, "ProprieteTTS", "")).Value), X + 235, Y + 2.5 + (6 * (i - 1)))
                    T.WrappingWidth = 30
                    T.AnchorPosition = CatTextAnchorPosition.catMiddleCenter
                    T.TextProperties.Justification = CatJustification.catCenter
                    T.SetFontSize(0, 0, 2.5)
                    T.Name = "TitleBlock_Data_Tableau_7_" & i - 1

                    T = mestexts.Add(".", X + 267.5, Y + 2.5 + (6 * (i - 1)))
                    T.WrappingWidth = 35
                    T.AnchorPosition = CatTextAnchorPosition.catMiddleCenter
                    T.TextProperties.Justification = CatJustification.catCenter
                    T.SetFontSize(0, 0, 2.5)
                    T.Name = "TitleBlock_Data_Tableau_8_" & i - 1

                End If

            End If
        End If





    End Sub

    Function CheckSiTextVide(s As String) As String
        On Error Resume Next
        If s Is Nothing Then
            Return "."
        End If
        Dim k As String = Replace(s, " ", "")
        If s = "" Then
            Return "."
        Else
            Return s
        End If
        On Error GoTo 0
    End Function

    Function CheckSi2D(d As Document) As Boolean
        If Right(d.FullName, 7) = "Drawing" Then
            Return True
        Else
            Return False
        End If
    End Function


#End Region

    Sub BRename()

        On Error Resume Next

        Dim exists As Integer

        For i = 1 To CATIA.Documents.Count
            If Right(CATIA.Documents.Item(i).Name, 10) = "CATProduct" Then
                If ListPartNumber.Contains(CATIA.Documents.Item(i).Product.Partnumber) Then
                    Dim CurrentProduct2 = CATIA.Documents.Item(i).Product.Products
                    For Each Item In CurrentProduct2
                        Dim k As Integer = 1
                        For j As Integer = 1 To CurrentProduct2.Count
                            If CurrentProduct2.Item(j).PartNumber = Item.PartNumber Then
                                exists = 0
                                Dim NewInstanceName As String = "_"
                                If Env = "[SPIRIT AEROSYSTEMS]" Then
                                    If Item.PartNumber Like "*_*_*" Then
                                        Dim a() As String = Strings.Split(Item.PartNumber, "_")
                                        Dim Nber = a(1)
                                        Nber = Replace(Nber, "_", "")
                                        NewInstanceName = Nber & "." & k
                                    Else
                                        NewInstanceName = Item.PartNumber & "." & k
                                    End If
                                Else
                                    NewInstanceName = Item.PartNumber & "." & k
                                End If
                                For m = 1 To CurrentProduct2.Count
                                    If CurrentProduct2.Item(m).Name = NewInstanceName Then
                                        exists = 1
                                        Exit For
                                    End If
                                Next
                                If exists = 0 Then
                                    CurrentProduct2.Item(j).Name = NewInstanceName
                                End If
                                k = k + 1
                            End If
                        Next
                    Next
                End If
            End If
        Next


        Dim m_ As New MessageErreur("Les instances ont été renommées avec succès", Notifications.Wpf.NotificationType.Information)

        On Error GoTo 0
    End Sub

End Class

Public Class ItemCatia

#Region "Properties"
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Private Sub NotifyPropertyChanged(ByVal info As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(info))
    End Sub



    Public Property ListTVItem As New List(Of ItemTV)
    Public Property ListTVitem_ As New List(Of TreeViewItem)

    Public Property Type As String
    Public Property PartNumber As String
    Public Property FileName As String
    Public Property Owner As String
    Public Property Doc As Document
    Public Property ProductCATIA As Product

    Public Property ID As Integer
    Public Property Level As Integer

    Public Property Qte As String
    Public Property Nomenclature As String
    Public Property DescriptionRef As String
    Public Property Defintion As String
    Public Property Revision As String
    Public Property Source As String

    Public Property IsSelected As Integer
    Private Property _enfants As New List(Of ItemCatia)
    Public Property Visible As Boolean = True


    Property l As New List(Of itemCATIAProperties)




    ReadOnly Property Enfants As List(Of ItemCatia)
        Get
            Return _enfants
        End Get
    End Property
#End Region

    Sub New(d As Document)

        If d Is Nothing Then
            Type = "COMPOSANT"
        Else
            If Right(d.FullName, 4) = "Part" Or Right(d.FullName, 7) = "Product" Then


                ProductCATIA = d.product
                PartNumber = ProductCATIA.PartNumber
                Doc = d
                PartNumber = ProductCATIA.PartNumber
                Owner = d.Name
                FileName = d.FullName
                Nomenclature = ProductCATIA.Nomenclature
                Defintion = ProductCATIA.Definition
                Revision = ProductCATIA.Revision
                DescriptionRef = ProductCATIA.DescriptionRef


                If Right(d.FullName, 4) = "Part" Then Type = "PART"
                If Right(d.FullName, 7) = "Product" Then Type = "PRODUCT"
                If Right(d.FullName, 7) = "Drawing" Then Type = "DRAWING"
                If Right(d.FullName, 7) = "Product" And PartNumber = CATIA.ActiveDocument.product.Partnumber Then Type = "RACINE"

                Select Case ProductCATIA.Source
                    Case 0
                        Source = "Inconnu"
                    Case 1
                        Source = "Fabriqué"
                    Case 2
                        Source = "Acheté"
                End Select

                'UserRefProperties
                Dim l_ As New List(Of itemCATIAProperties)
                For Each item As Parameter In ProductCATIA.UserRefProperties
                    Dim myP As New itemCATIAProperties(item.Name, item.ValueAsString)
                    l_.Add(myP)
                Next

                For i = 1 To listPropertiesall.Count
                    l.Add(New itemCATIAProperties("", ""))
                Next
                For Each item In l_
                    For j = 0 To listPropertiesall.Count - 1
                        If item.Name = listPropertiesall(j) Then
                            l(j) = item
                            Exit For
                        End If
                    Next
                Next

                ListDocuments.Add(Me)

                PercentNbElements += 1
                Try
                    Bgw.ReportProgress(PercentNbElements / NbElements * 100, "Récupération des paramètres de l'élément [" & d.Name & "]")
                Catch ex As Exception
                End Try


            End If
        End If

    End Sub


    Sub MAJTV()

        For Each item In ListTVItem
            item.TextTV = PartNumber & " | " & DescriptionRef
        Next

    End Sub






End Class

Public Class itemCATIAProperties

    Property Name As String
    Property Value As String

    Sub New(item As String, _value As String)
        Dim s() As String = Strings.Split(item, "\")
        Name = s(UBound(s))
        Value = _value
    End Sub
End Class
Public Class PropertiesPart

    Public Property Properties As String
    Public Property FullName As String
    Public Property Value As String

    Sub New(p As Parameter)

        Dim s() As String = Strings.Split(p.Name, "\")
        Dim s_ As String = s(UBound(s))
        FullName = p.Name
        Properties = s_
        Value = p.ValueAsString

    End Sub

    Sub MajProperties()

        For Each item As Parameter In MonActiveDoc.product.UserRefProperties
            If item.Name = FullName Then
                item.ValuateFromString(Value)
            End If
        Next
    End Sub
End Class

Public Class ItemTV


    Public Property TextTV As String
    Public Property ItemCATIA As ItemCatia
    Public Property PartNumber As String
    Public Property Descritpion As String
    Public Property TVitem As TreeViewItem
    Public Property Level As Integer
    Public Property Type As String


    Sub New(name As String)
        TextTV = GetTextTv(name)
        TVitem = CreerNewNode()
    End Sub

    Function CreerNewNode() As TreeViewItem

        Dim Tn As New TreeViewItem
        Dim Lab As New Label
        Dim st As New StackPanel

        Dim Im As New Image With {
            .Width = 16,
            .Height = 16
        }

        If ItemCATIA Is Nothing Then Type = "COMPOSANT"

        Select Case Type
            Case "RACINE"
                Im.Source = New BitmapImage(New Uri(DossierImage & "CATIAProduct.ico"))
            Case "PRODUCT"
                Im.Source = New BitmapImage(New Uri(DossierImage & "CATIAProduct.ico"))
            Case "COMPOSANT"
                Im.Source = New BitmapImage(New Uri(DossierImage & "CATIAComposants.bmp"))
            Case "PART"
                Im.Source = New BitmapImage(New Uri(DossierImage & "CATIAPart.ico"))
            Case Else
                Im.Source = New BitmapImage(New Uri(DossierImage & "CATIAProduct.ico"))
        End Select

        Dim MonBind As New Binding("TextTV") With {
            .Source = Me
        }
        Lab.SetBinding(Label.ContentProperty, MonBind)
        st.Orientation = Orientation.Horizontal
        st.Children.Add(Im)
        st.Children.Add(Lab)
        Tn.Header = st
        Tn.DataContext = Me

        Return Tn

    End Function

    Sub MajHeader(SourceImage)

        Dim Tn As TreeViewItem = TVitem
        Dim Lab As New Label
        Dim st As New StackPanel

        Dim Im As New Image With {
            .Width = 16,
            .Height = 16,
            .Source = New BitmapImage(New Uri(DossierImage & SourceImage))
        }


        Dim MonBind As New Binding("TextTV") With {
            .Source = Me
        }
        Lab.SetBinding(Label.ContentProperty, MonBind)
        st.Orientation = Orientation.Horizontal
        st.Children.Add(Im)
        st.Children.Add(Lab)
        Tn.Header = st
        Tn.DataContext = Me


    End Sub

    Function GetTextTv(str As String) As String


        Dim str_() As String = Strings.Split(str, " (")
        Dim PN As String = str_(0)
        For i = 1 To str_.Length - 2
            PN = PN & str_(i)
        Next

        PartNumber = PN

        For Each ic In ListDocuments
            If ic.PartNumber = PartNumber Then
                Descritpion = ic.DescriptionRef
                ItemCATIA = ic
                Exit For
            End If
        Next

        If ItemCATIA Is Nothing Then
        Else
            ItemCATIA.ListTVItem.Add(Me)
        End If

        Return PN & " | " & Descritpion


    End Function

End Class 'Item TreeView
















Module FixAll
    Dim oList
    Dim oSelection As INFITF.Selection
    Dim oVisProp As VisPropertySet
    Dim Pint As Integer = 0

    Sub CATMain(oTopDoc As Document)

        Try
            'Declarations
            Dim oTopProd As ProductDocument = Nothing
            Dim oCurrentProd As Object

            'Check si c'est un assemblage
            If Strings.Right(oTopDoc.Name, 7) <> "Product" Then
                Dim MsgErr As New MessageErreur("Impossible de fixer un document autre qu'un assemblage", Notifications.Wpf.NotificationType.Warning)
                Exit Sub
            End If

            oSelection = oTopDoc.Selection
            oVisProp = oSelection.VisProperties
            oCurrentProd = oTopDoc.Product
            oList = CreateObject("Scripting.dictionary")

            oSelection.Clear()



            FixSingleLevel(oCurrentProd)


            Dim MsgErr2 As New MessageErreur("L'ensemble [" & MonActiveDoc.Name & "] a été fixé", Notifications.Wpf.NotificationType.Information)
        Catch ex As Exception
        End Try



    End Sub
    Private Sub FixSingleLevel(ByRef oCurrentProd As Object)

        On Error Resume Next


        Dim ItemToFix As Product
        Dim iProdCount As Integer
        Dim i As Integer
        Dim j As Integer
        Dim oConstraints As Constraints
        Dim oReference As Reference
        Dim sItemName As String
        Dim constraint1 As MECMOD.Constraint
        Dim pActivation As KnowledgewareTypeLib.Parameter
        Dim N, m As Integer
        Dim sActivationName As String

        Err.Clear()
        oCurrentProd = oCurrentProd.ReferenceProduct
        iProdCount = oCurrentProd.Products.Count
        oConstraints = oCurrentProd.Connections("CATIAConstraints")

        N = oConstraints.Count
        m = N
        For i = 1 To m
            oConstraints.Remove(N)
            N = N - 1
        Next


        For i = 1 To iProdCount
            Pint += 1
            ItemToFix = oCurrentProd.Products.Item(i)

CreateReference:

            sItemName = ItemToFix.Name

            oReference = oCurrentProd.CreateReferenceFromName(sItemName & "/!" & "/")

            constraint1 = oConstraints.AddMonoEltCst(CatConstraintType.catCstTypeReference, oReference)
            constraint1.ReferenceType = CatConstraintRefType.catCstRefTypeFixInSpace

            oSelection.Add(constraint1)
            oVisProp.SetShow(CatVisPropertyShow.catVisPropertyNoShowAttr)
            oSelection.Clear()

RecursionCall:
            If ItemToFix.Products.Count <> 0 Then
                If oList.exists(ItemToFix.PartNumber) Then GoTo Finish

                If ItemToFix.PartNumber = ItemToFix.ReferenceProduct.Parent.Product.PartNumber Then oList.Add(ItemToFix.PartNumber, 1)
                Call FixSingleLevel(ItemToFix)
            End If
Finish:
        Next

    End Sub

End Module 'Module ToutFix

Public Module ExportPDFDXF

    Sub GoToPDF(Fold As String, CheckPDF As Boolean, CheckDXF As Boolean, CheckDWG As Boolean)

        Dim fileNames = My.Computer.FileSystem.GetFiles(
        Fold, FileIO.SearchOption.SearchTopLevelOnly, "*.CATDrawing")

        Dim Errb As Boolean = False

        For Each f As String In fileNames

            If f Like "*.CATDrawing" Then
                Dim Folder As String = Path.GetDirectoryName(f)

                Dim CatDraw As DrawingDocument = CATIA.Documents.Open(f)
                Dim NamePDF As String = Strings.Left(CatDraw.Name, InStr(CatDraw.Name, ".CATDrawing") - 1)

                If CheckPDF = True Then
                    If System.IO.Directory.Exists(Folder & "\PDF") = False Then
                        System.IO.Directory.CreateDirectory(Folder & "\PDF")
                    End If

                    Try

                        Dim NomFichier As String = Folder & "\PDF\" & NamePDF & ".pdf"
                        If IO.File.Exists(NomFichier) Then
                            IO.File.Delete(NomFichier)
                        End If
                        CatDraw.ExportData(NomFichier, "pdf")
                    Catch ex As Exception
                        Errb = True
                    End Try

                End If

                If CheckDXF = True Then
                    If System.IO.Directory.Exists(Folder & "\DXF") = False Then
                        System.IO.Directory.CreateDirectory(Folder & "\DXF")
                    End If

                    Try
                        Dim NomFichier As String = Folder & "\DXF\" & NamePDF & ".DXF"
                        If IO.File.Exists(NomFichier) Then
                            IO.File.Delete(NomFichier)
                        End If
                        CatDraw.ExportData(NomFichier, "dxf")
                    Catch ex As Exception
                        Errb = True
                    End Try
                End If

                If CheckDWG = True Then
                    If System.IO.Directory.Exists(Folder & "\DWG") = False Then
                        System.IO.Directory.CreateDirectory(Folder & "\DWG")
                    End If

                    Try
                        Dim NomFichier As String = Folder & "\DWG\" & NamePDF & ".dwg"
                        If IO.File.Exists(NomFichier) Then
                            IO.File.Delete(NomFichier)
                        End If
                        CatDraw.ExportData(NomFichier, "DWG")
                    Catch ex As Exception
                        Errb = True
                    End Try
                End If

                CatDraw.Close()

            End If

        Next

        If Errb = False Then
            Dim m As New MessageErreur("La conversion des fichiers s'est terminée avec succès", Notifications.Wpf.NotificationType.Information)
        Else
            Dim m As New MessageErreur("Une erreur s'est produite lors de conversion de certains fichiers. Catia doit être ouvert.", Notifications.Wpf.NotificationType.Warning)
        End If

    End Sub

End Module 'export PDF vers DXF

Public Module GoRelinkDraw
    Sub RelinkDoc()

        Dim AC As Document

        Try

            If Right(CATIA.ActiveDocument.FullName, 7) = "Drawing" Then
                AC = CATIA.ActiveDocument
            Else
                Dim MsgErr As New MessageErreur("Ouvrir un Drawing avant de lancer la macro.", Notifications.Wpf.NotificationType.Warning)
                Exit Sub

            End If
        Catch ex As Exception
            Dim MsgErr As New MessageErreur("Ouvrir un Drawing avant de lancer la macro.", Notifications.Wpf.NotificationType.Warning)
            Exit Sub
        End Try

        Dim MesVues As DrawingView
        Dim Draw As DrawingDocument = AC
        Dim MaSheet As DrawingSheet = Draw.Sheets.ActiveSheet


        Dim File As String

        Dim BrowserFile As New Ookii.Dialogs.Wpf.VistaOpenFileDialog With {
            .Title = "Selection du fichier à link"
        }

        If BrowserFile.ShowDialog() = True Then
            File = BrowserFile.FileName
        Else
            Exit Sub
        End If

        FctionCATIA.OpenFile(File)
        Dim NewDoc
        Dim f_() As String = Strings.Split(File, "\")
        File = f_(UBound(f_))
        NewDoc = CATIA.Documents.Item(File)



        Dim i
        i = 0
        For Each MesVues In MaSheet.Views
            If i > 1 Then
                MesVues.GenerativeLinks.RemoveAllLinks()
                MesVues.GenerativeLinks.AddLink(NewDoc.Product)
            End If
            i = i + 1
        Next

        MaSheet.Update()
        MaSheet.Update()


    End Sub

    Sub RelinkDoctestSpirit()

        Dim AC As Document

        Try

            If Right(CATIA.ActiveDocument.FullName, 7) = "Drawing" Then
                AC = CATIA.ActiveDocument
            Else
                Dim MsgErr As New MessageErreur("Ouvrir un Drawing avant de lancer la macro.", Notifications.Wpf.NotificationType.Warning)
                Exit Sub

            End If
        Catch ex As Exception
            Dim MsgErr As New MessageErreur("Ouvrir un Drawing avant de lancer la macro.", Notifications.Wpf.NotificationType.Warning)
            Exit Sub
        End Try

        Dim MesVues As DrawingView
        Dim Draw As DrawingDocument = AC
        Dim MaSheet As DrawingSheet = Draw.Sheets.ActiveSheet


        Dim File As String = Draw.Name
        File = Strings.Replace(File, "SHT", "")
        File = Strings.Replace(File, "SD", "")
        File = Strings.Replace(File, "MIT15V09009", "MIT15V10009")
        File = Strings.Left(File, Len(File) - 11)

        Dim TestF() As FileInfo = New DirectoryInfo("P:\2020-A077-NZ SPIRIT AEROSYSTEMS_MIT15V10000 - STEP7\5-CAO").GetFiles(File & "*.*")
        File = TestF(0).FullName

        FctionCATIA.OpenFile(File)
        Dim NewDoc
        Dim f_() As String = Strings.Split(File, "\")
        File = f_(UBound(f_))
        NewDoc = CATIA.Documents.Item(File)



        Dim i
        i = 0
        For Each MesVues In MaSheet.Views
            If i > 1 Then
                MesVues.GenerativeLinks.RemoveAllLinks()
                MesVues.GenerativeLinks.AddLink(NewDoc.Product)
            End If
            i = i + 1
        Next

        MaSheet.Update()
        MaSheet.Update()


        For Each T As DrawingText In MaSheet.Views.Item(2).Texts


            Try
                If T.Name = "Text.82" Or T.Name = "Text.249" Then
                    T.Text = "MIT15V10009"
                End If
                If T.Name = "Text.72" Or T.Name = "Text.236" Then
                    T.Text = "02/07/20"
                End If
            Catch ex As Exception
            End Try

        Next



    End Sub

End Module

Public Module Drill

    Sub CATMain()


        Dim Diam As Double = 0

        Dim PartDoc As PartDocument
        Try
            PartDoc = CATIA.ActiveDocument
        Catch ex As Exception
            Dim newm As New MessageErreur("Une part doit être ouverte pour pouvoir continuer", Notifications.Wpf.NotificationType.Warning)
            Exit Sub
        End Try



        Try
            Diam = Replace(InputBox("Diamètre en mm :"), ".", ",")
        Catch ex As Exception
        End Try

        If Diam <= 0 Then
            Exit Sub
        End If



        Dim MaPart As Part = PartDoc.Part
        Dim HB As HybridBodies = MaPart.HybridBodies

        Dim HSF As HybridShapeFactory = MaPart.HybridShapeFactory

        Dim sel_ As Selection = CATIA.ActiveDocument.Selection
        sel_.Clear()
        Dim a_(0)
        Dim status_
        a_(0) = "HybridBody"


        AppActivate("CATIA V5 - [" & CATIA.ActiveDocument.Name & "]")
        status_ = sel_.SelectElement2(a_, "Selection du set géométrique", False)
        If status_ = "Cancel" Then
            Exit Sub
        End If


        Dim HbodiesAxes_ As HybridBody = sel_.Item(1).Value
        Dim HBodiesAxes As HybridBody = HbodiesAxes_

        Dim sel As Selection = CATIA.ActiveDocument.Selection
        sel.Clear()
        Dim a(0)
        a(0) = "BiDim"
        Dim status
        status = sel.SelectElement2(a, "Selection de la surface avion", False)
        If status = "Cancel" Then
            Exit Sub
        End If
        Dim surf As HybridShapeSurfaceExplicit = sel.Item(1).Value
        Dim SurfaceAvion As HybridShapeSurfaceExplicit = surf

        Dim i As Integer = 1

        Dim NewHB As HybridBody = HB.Add()
        NewHB.Name = "Construction axes_"

        Dim B As Bodies = MaPart.Bodies
        Dim NewB As Body = B.Add()
        NewB.Name = "Perçages"

        For Each item In HBodiesAxes.HybridShapes


            Dim Lines
            Dim Lines1 As HybridShapeCurveExplicit
            Dim Lines2 As HybridShapeLineExplicit
            Try
                Lines1 = item
            Catch ex As Exception
                Lines = Nothing
            End Try

#Disable Warning BC42104 ' La variable 'Lines1' est utilisée avant qu'une valeur ne lui ait été assignée. Une exception de référence null peut se produire au moment de l'exécution.
            Lines = Lines1
#Enable Warning BC42104 ' La variable 'Lines1' est utilisée avant qu'une valeur ne lui ait été assignée. Une exception de référence null peut se produire au moment de l'exécution.

            If Lines Is Nothing Then
                Try
                    Lines2 = item
                Catch ex As Exception
                    Lines = Nothing
                End Try
#Disable Warning BC42104 ' La variable 'Lines2' est utilisée avant qu'une valeur ne lui ait été assignée. Une exception de référence null peut se produire au moment de l'exécution.
                Lines = Lines2
#Enable Warning BC42104 ' La variable 'Lines2' est utilisée avant qu'une valeur ne lui ait été assignée. Une exception de référence null peut se produire au moment de l'exécution.
            End If


            If Lines Is Nothing Then
            Else


                Dim Ref1 As Reference
                Dim Ref2 As Reference


                Dim SetPercages As HybridBody = NewHB.HybridBodies.Add()
                SetPercages.Name = "Perçage " & i


                Ref1 = MaPart.CreateReferenceFromObject(Lines)
                Ref2 = MaPart.CreateReferenceFromObject(SurfaceAvion)

                Dim Intersection As HybridShapeIntersection = HSF.AddNewIntersection(Ref1, Ref2)
                Intersection.PointType = 0
                Intersection.Name = "Point_" & i
                SetPercages.AppendHybridShape(Intersection)
                MaPart.InWorkObject = Intersection
                '   MaPart.Update()
                sel.Clear()
                sel.Add(Intersection)
                sel.Copy()
                sel.Clear()
                sel.Add(SetPercages)
                sel.PasteSpecial("CATPrtResultWithOutLink")
                sel.Clear()

                Dim PtIntersect = SetPercages.HybridShapes.Item(1)
                Dim m 'As Measurable
                Dim Spa = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench")
                Dim Coords(2)
                Ref2 = MaPart.CreateReferenceFromObject(PtIntersect)
                m = Spa.GetMeasurable(Ref2)
                m.GetPoint(Coords)

                Ref1 = MaPart.CreateReferenceFromObject(Intersection)
                Ref2 = MaPart.CreateReferenceFromObject(SurfaceAvion)

                Dim MaDroite As HybridShapeLineNormal = HSF.AddNewLineNormal(Ref2, Ref1, 20, -20, False)
                SetPercages.AppendHybridShape(MaDroite)
                MaDroite.Name = "Axe_" & i
                MaPart.InWorkObject = MaDroite

                Ref1 = MaPart.CreateReferenceFromObject(MaDroite)
                Ref2 = MaPart.CreateReferenceFromObject(Intersection)
                Dim MonPlan As HybridShapePlaneNormal = HSF.AddNewPlaneNormal(Ref1, Ref2)
                SetPercages.AppendHybridShape(MonPlan)
                MonPlan.Name = "Plan_" & i
                MaPart.InWorkObject = MonPlan


                Ref1 = MaPart.CreateReferenceFromObject(MonPlan)
                Dim MonPlan2 As HybridShapePlaneOffset = HSF.AddNewPlaneOffset(MonPlan, 50, False)
                SetPercages.AppendHybridShape(MonPlan2)
                MonPlan2.Name = "Dec_" & i
                MaPart.InWorkObject = MonPlan2

                MaPart.InWorkObject = NewB

                Dim SF As ShapeFactory = MaPart.ShapeFactory
                Ref2 = MaPart.CreateReferenceFromObject(MonPlan2)
                MaPart.InWorkObject = B
                MaPart.Update()
                Dim MonHole As Hole
                MonHole = SF.AddNewHoleFromPoint(Coords(0), Coords(1), Coords(2), Ref2, 100)
                '     MaPart.Update()
                Dim l As Length = MonHole.Diameter
                l.Value = Diam
                '   MaPart.Update()
                MonHole.BottomType = CatHoleBottomType.catFlatHoleBottom
                MonHole.Type = CatHoleType.catSimpleHole
                Ref2 = MaPart.CreateReferenceFromObject(MaDroite)
                MonHole.SetDirection(Ref2)

                i += 1
            End If


        Next


        Dim errNotif As Boolean = False
        Try
            MaPart.Update()
        Catch ex As Exception

            errNotif = True
        End Try

        AppActivate("Application KD | 2.0.1")

        If errNotif = True Then
            Dim newm4 As New MessageErreur("Une erreur s'est produite lors de la tentative de mise à jour de la Part", Notifications.Wpf.NotificationType.Warning)
        End If
        Dim newm3 As New MessageErreur("Perçages générés avec succès", Notifications.Wpf.NotificationType.Information)

    End Sub
End Module

Public Class LineBOM

    Public Property REP As String
    Public Property QTE As String
    Public Property DESIGNATION As String
    Public Property PartNumber As String
    Public Property MATIERE As String
    Public Property ETAT As String
    Public Property DIMBRUTES As String
    Public Property OBSERVATIONS As String
    Public Property ANGLAIS As String
    Public Property PLANCHE As String

    Sub New()

    End Sub


End Class

Public Module ModifPlan

    Sub GoMajPlan(ListFiles As List(Of String), GoDate As Boolean, GoNumOutillage As Boolean, GoDessinateur As Boolean, GoNomenclature As Boolean, GoPLANT As Boolean, GoTitre As Boolean, GoPROGRAMM As Boolean, date_ As String, numOutillage As String, dessinateur As String, plant As String, titre1 As String, titre2 As String, program1 As String, program2 As String)

        Dim Boolerror As Boolean = False


        For Each item In ListFiles
            Dim CatDraw As DrawingDocument
            Try
                CatDraw = CATIA.Documents.Open(item)
            Catch ex As Exception
                Dim az3 As New MessageErreur("Une erreur s'est produite. Ouvrir CATIA.", Notifications.Wpf.NotificationType.Warning)
                Exit Sub
            End Try

            Dim MaSheet As DrawingSheet = CatDraw.Sheets.ActiveSheet
            Dim V As DrawingView = MaSheet.Views.Item("Background View")

            Dim MesTexts As DrawingTexts = V.Texts
            Dim Type As Integer = MaSheet.PaperSize


            Dim CountLine As Integer = 0

            For Each T As DrawingText In MaSheet.Views.Item(2).Texts


                Try
                    If GoDate = True Then
                        If T.Name = "AUKTbkText_JAT_ALL_DRN_DATE" Or T.Name = "AUKTbkText_JAT_ALL_CHKD_DATE" Or T.Name = "AUKTbkText_JAT_ALL_APPD_DATE" Then
                            T.Text = date_
                        End If
                    End If

                    If GoNumOutillage = True Then
                        If T.Name = "AUKTbkText_JAT_ALL_DRAWING_NUMBER" Then
                            T.Text = numOutillage
                        End If
                    End If

                    If GoPLANT = True Then
                        If T.Name = "AUKTbkText_JAT_ALL_PLANT" Then
                            T.Text = plant
                        End If
                    End If

                    If GoDessinateur = True Then
                        If T.Name = "AUKTbkText_JAT_ALL_DRN" Or T.Name = "AUKTbkText_JAT_ALL_CHKD" Or T.Name = "AUKTbkText_JAT_ALL_APPD" Or T.Name = "DA_SOC" Then
                            T.Text = dessinateur
                        End If
                    End If

                    If GoTitre = True Then
                        If T.Name = "AUKTbkText_JAT_ALL_TITLE_L1" Then
                            T.Text = titre1
                        End If
                        If T.Name = "AUKTbkText_JAT_ALL_TITLE_L2" Then
                            T.Text = titre2
                        End If
                    End If

                    If GoPROGRAMM = True Then
                        If T.Name = "AUKTbkText_JAT_AIF_PROGRAM_L3" Then
                            T.Text = program1
                        End If
                        If T.Name = "AUKTbkText_JAT_AIF_PROGRAM_L1" Then
                            T.Text = ""
                        End If
                        If T.Name = "AUKTbkText_JAT_AIF_PROGRAM_L2" Then
                            T.Text = ""
                        End If
                        If T.Name = "AUKTbkText_JAT_AIF_PROGRAM_L4" Then
                            T.Text = program2
                        End If
                        If T.Name = "AUKTbkText_JAT_AIF_PROGRAM_L5" Then
                            T.Text = ""
                        End If
                    End If

                    If T.Name Like "NomenclatureText_PartNumber_" & "*" Then
                        CountLine = CountLine + 1
                    End If
                Catch ex As Exception
                    Boolerror = True
                End Try

            Next

            If GoNomenclature = True Then

                Dim ListLineBomFrom2D As New List(Of LineBOM)

                For i = 1 To CountLine

                    Dim k As String = FctionCATIA.GetStringFromText("NomenclatureText_PartNumber_" & i, V)
                    Dim b As Boolean = False
                    Dim ICok As ItemCatia = Nothing
                    Dim ReelDescription As String = ""
                    For Each ic_ As ItemCatia In ListDocuments
                        If ic_.PartNumber = k Then
                            b = True
                            ICok = ic_
                            ReelDescription = ICok.ProductCATIA.DescriptionRef
                            Exit For
                        End If
                    Next

                    If b = True Then
                        Dim L As New LineBOM With {
                                  .REP = FctionCATIA.GetStringFromText("NomenclatureText_repere_" & i, V),
                                  .PLANCHE = FctionCATIA.GetStringFromText("NomenclatureText_planche_" & i, V),
                                  .QTE = FctionCATIA.GetStringFromText("NomenclatureText_quantite_" & i, V),
                                  .DESIGNATION = ReelDescription,
                                  .PartNumber = FctionCATIA.GetStringFromText("NomenclatureText_PartNumber_" & i, V),
                                  .MATIERE = FctionCATIA.GetStringFromText("NomenclatureText_matiere_" & i, V),
                                  .ETAT = FctionCATIA.GetStringFromText("NomenclatureText_etat_" & i, V),
                                  .DIMBRUTES = FctionCATIA.GetStringFromText("NomenclatureText_dimensions_brutes_" & i, V),
                                  .OBSERVATIONS = FctionCATIA.GetStringFromText("NomenclatureText_observations_" & i, V)
                              }
                        ListLineBomFrom2D.Add(L)
                    End If
                Next i

                V.Activate()
                FctionCATIA.DeleteBOMExistante(MaSheet.Views.Item(2).Texts)
                FctionCATIA.CreerTitreBOM(V, MesTexts, Type)
                MaSheet.Views.Item(1).Activate()

            End If


            CatDraw.Save()
            CatDraw.Close()

        Next



        If Boolerror = False Then
            Dim az3 As New MessageErreur("Les plans séléctionés ont été mis à jour avec succès", Notifications.Wpf.NotificationType.Information)
        Else
            Dim az3 As New MessageErreur("Des erreurs se sont produites lors de la modification de certains plans. Vérifier les données d'entrée.", Notifications.Wpf.NotificationType.Warning)
        End If


    End Sub
End Module

Public Module ChangeUnitsDrawing 'Change les unités d'un Draw

    Sub MainUnitsDraw(BoolMM As Boolean, BoolINCH As Boolean, BoolALL As Boolean)

        Dim AC As Document

        Try
            If Right(CATIA.ActiveDocument.FullName, 7) = "Drawing" Then
                AC = CATIA.ActiveDocument
            Else
                Dim MsgErr As New MessageErreur("Ouvrir un Drawing avant de lancer la macro.", Notifications.Wpf.NotificationType.Warning)
                Exit Sub
            End If
        Catch ex As Exception
            Dim MsgErr As New MessageErreur("Ouvrir un Drawing avant de lancer la macro.", Notifications.Wpf.NotificationType.Warning)
            Exit Sub
        End Try



        Dim MaVue 'As DrawingView
        Dim Draw As DrawingDocument = AC
        Dim MaSheet As DrawingSheet = Draw.Sheets.ActiveSheet

        Dim i As Integer = 0

        For Each MaVue In MaSheet.Views
            i = i + 1
            If i > 2 Then
                If MaVue.Dimensions.Count > 0 Then
                    For Each item In MaVue.Dimensions
                        Dim DD As DrawingDimension = item
                        Dim DV As DrawingDimValue = DD.GetValue

                        If BoolMM = True Then 'affiche toutes les cotes en mm
                            DD.DualValue = CatDimDualDisplay.catDualNone
                            DV.SetFormatUnit(1, 0)
                        End If

                        If BoolINCH = True Then 'affiche toutes les cotes en inch
                            DD.DualValue = CatDimDualDisplay.catDualNone
                            DV.SetFormatUnit(1, 1)
                        End If

                        If BoolALL = True Then 'affiche la DualValue + affiche les cotes en mm ET en Inch
                            DD.DualValue = CatDimDualDisplay.catDualFractional
                            DV.SetFormatUnit(1, 0)
                            DV.SetFormatUnit(2, 1)
                            DV.SetBaultText(2, "", "'", "", "")
                        End If
                    Next
                End If
            End If
        Next
    End Sub

End Module






'------- MATERIAL POPUP --------
Public Class ItemFamille

    Public Property Name As String
    Public Property Materials As New List(Of ItemMaterial)
    Sub New(NomFamille As String)
        Name = NomFamille
        ListFamilleMaterials.Add(Me)
    End Sub
End Class

Public Class ItemMaterial

    Public Property Famille As String
    Public Property Name As String


    Sub New(NomMaterial As String, FamilleMaterial As String)

        Famille = FamilleMaterial
        Name = NomMaterial
        ListMaterials.Add(Me)
    End Sub


End Class
