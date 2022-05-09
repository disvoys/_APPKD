
Imports System.ComponentModel
Imports System.Data
Imports System.Windows.Controls.Primitives
Imports MySqlConnector
Module Variables


    Public GoCreerPropertiesDassault As Boolean = False


#Region "Start"
    Public BoolStart As Boolean = False
    Public Bgw As BackgroundWorker
#End Region

#Region "Pages"

    Public DialogPlanA320 As DialogChoixPlan = New DialogChoixPlan
    Public WindowPDF As Window_PDxF = New Window_PDxF
    Public WindowRenameDassault_ As WindowRenameDassault = New WindowRenameDassault
    Public WindowUnitDraw As Window_UnitsDraw = New Window_UnitsDraw
    Public MonMainV3 As MainV3 = New MainV3


#End Region

#Region "CATIA"
    Public WithEvents FctionCATIA As CatiaClass = New CatiaClass
    Public WithEvents FctionLOAD As Load = New Load
    Public WithEvents FctionPDF As ClassPDF = New ClassPDF
    Public WithEvents FctionGetInfo As ClassInformationComputer = New ClassInformationComputer
    Public WithEvents FctionGetBOM As ClassGetBOMfromTV = New ClassGetBOMfromTV
    Public WithEvents FctionRenameDassault As ClassRenameDassault = New ClassRenameDassault



    Public CATIA As INFITF.Application
    Public MonActiveDoc As INFITF.Document
    Public TypeActiveDoc As String = ""
    Public ListDocuments As New List(Of ItemCatia)
    Public ListDocumentsStr As New List(Of String)

    Public ListFabriques As New List(Of ItemCatia)
    Public ListAchetes As New List(Of ItemCatia)
    Public ListInconnus As New List(Of ItemCatia)


    Public ListMaterials As New List(Of ItemMaterial)
    Public ListFamilleMaterials As New List(Of ItemFamille)
    Public ColDoc As New ListCollectionView(ListDocuments)
    Public ListItemCatia As New List(Of ItemCatia)
    Public ListItemTV As New List(Of ItemTV)


    Public NbElements As Integer = 0
    Public PercentNbElements As Integer = 1

    Public CatalogueMatieres As String

    Public ListPartNumber As New List(Of String)
    Public ListNomWindows As New List(Of String)

    Public ICRacine As ItemCatia = Nothing
    Public ITRacine As ItemTV = Nothing
#End Region

#Region "INIFILE"
    Public DossierBase As String = Replace(My.Application.Info.DirectoryPath, "\bin\Debug", "")
    Public INIFiles As GestionINIFiles
    Public INIProperties As GestionINIFiles

#End Region

#Region "Fichiers"
    Public FichierTreeTxt As String
    Public TreeSauv As String
    Public FichierListingReport As String
    Public FichierListingReportsimplifié As String
#End Region

#Region "Dossier Image"
    Public DossierImage As String = DossierBase & "\Icons\"
#End Region

    Public DossierSettingsR27 As String
    Public DossierSettingsR28 As String
    Public DefaultSettings As String

    'sql

    Public cn As New MySqlConnection
    Public dr As MySqlDataReader

    Public dt As New DataTable
    Public dtb As New DataTable
    Public sda As MySqlDataAdapter

    Public cSQL As New ClassDB


    Public ListPropertiesPart As New List(Of PropertiesPart)

    'users
    Public cn_ As New MySqlConnection
    Public dr_ As MySqlDataReader

    Public dt_ As New DataTable
    Public dtb_ As New DataTable
    Public sda_ As MySqlDataAdapter

    Public ListUsers As New List(Of ClassUsers)

    Public ColDocUsers As New ListCollectionView(ListUsers)
    Public QteUsers As Integer = 0
    Public ListUsersstr As New List(Of String)
    Public cSQLUsers As New LoadDTUsers

    'LANGUE
    Public MaLangue As String = "Anglais"

    'ENVIRONNMENTS
    Public ListEnvironnements As New List(Of String)
    Public ListToogleSettingsEnv As New List(Of ToggleButton)
    Public Env As String = ""

    Public ListPropertiesPartEnCours As New List(Of String)

    Public listPropertiesall As New List(Of String)




    Public BoolStartBOM As Boolean = False

    Function getItemListProperties(s As String)
        Dim i As Integer = 0
        For Each item In listPropertiesall
            If item = s Then
                Return i
            End If
        Next
        Return Nothing
    End Function
    Function GetEnv() As String
        Dim n As String = Env
        n = Strings.Replace(n, "[", "")
        n = Strings.Replace(n, "]", "")
        Return n
    End Function
End Module
