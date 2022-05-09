Imports System.Drawing
Imports System.IO
Imports Org.pdfclown.documents
Imports Org.pdfclown.documents.contents
Imports Org.pdfclown.documents.contents.colorSpaces
Imports Org.pdfclown.documents.contents.composition
Imports Org.pdfclown.documents.contents.fonts
Imports Org.pdfclown.documents.contents.xObjects
Imports Org.pdfclown.files
Imports Org.pdfclown.tools
Imports File = Org.pdfclown.files.File

Public Class ClassPDF

    'kevin

    Sub split(Chemin As String, NameFile As String)

        Using f As New File(Chemin)
            Dim oldD As Document = f.Document
            If Not Directory.Exists(Path.GetTempPath & "\PDF_Maniement_AppKD") Then MkDir(Path.GetTempPath & "\PDF_Maniement_AppKD")

            For i As Integer = 0 To oldD.Pages.Count - 1
                Dim D As Document = New PageManager(oldD).Extract(i, i + 1)
                D.File.Save(Path.GetTempPath & "\PDF_Maniement_AppKD\" & NameFile & "_" & i & ".pdf", SerializationModeEnum.Standard)
            Next i

        End Using


        Dim k As New MessageErreur("Split du fichier réussi avec succès", Notifications.Wpf.NotificationType.Information)
        Process.Start(Path.GetTempPath & "\PDF_Maniement_AppKD")

    End Sub
    Sub merge(Chemin As List(Of String), NameFile As String)



        Using f As New File(Chemin(0))
            Dim D As Document = f.Document
            If Not Directory.Exists(Path.GetTempPath & "\PDF_Maniement_AppKD") Then MkDir(Path.GetTempPath & "\PDF_Maniement_AppKD")
            Dim Pm As New PageManager(D)
            For i = 1 To Chemin.Count - 1
                Dim fi As New File(Chemin(i))
                Pm.Add(fi.Document)
            Next

            D.File.Save(Path.GetTempPath & "\PDF_Maniement_AppKD\" & NameFile & "_" & ".pdf", SerializationModeEnum.Standard)
        End Using


        Dim k As New MessageErreur("Fusion des fichiers réussie avec succès", Notifications.Wpf.NotificationType.Information)
        Process.Start(Path.GetTempPath & "\PDF_Maniement_AppKD")

    End Sub
    Sub Marker(Chemin As String, NameFile As String, TextAffiche As String)

        Using f As New File(Chemin)
            Dim d As Document = f.Document
            If Not Directory.Exists(Path.GetTempPath & "\PDF_Maniement_AppKD") Then MkDir(Path.GetTempPath & "\PDF_Maniement_AppKD")

            Dim Stamp As New PageStamper
            Dim Size As SizeF = d.GetSize
            Dim WaterMark As New FormXObject(d, Size)
            Dim compo As New PrimitiveComposer(WaterMark)
            compo.SetFont(New StandardType1Font(d, StandardType1Font.FamilyEnum.Helvetica, False, False), 120)
            compo.SetFillColor(New DeviceRGBColor(157 / 255D, 27 / 255D, 48 / 255D))
            Dim state As New ExtGState(d)
            state.FillAlpha = 0.3
            compo.ApplyState(state)

            compo.ShowText(TextAffiche, New PointF(Size.Width / 2.0F, Size.Height / 2.0F), XAlignmentEnum.Center, YAlignmentEnum.Middle, 50)
            compo.Flush()



            For i As Integer = 0 To d.Pages.Count - 1
                Stamp.Page = d.Pages(i)
                Dim Composer As PrimitiveComposer = Stamp.Foreground
                Composer.ShowXObject(WaterMark)
                Stamp.Flush()

            Next i


            d.File.Save(Path.GetTempPath & "\PDF_Maniement_AppKD\" & NameFile & "_" & ".pdf", SerializationModeEnum.Standard)

        End Using

        Dim k As New MessageErreur("Ajout du marker avec succès", Notifications.Wpf.NotificationType.Information)
        Process.Start(Path.GetTempPath & "\PDF_Maniement_AppKD")

    End Sub


End Class
