Imports Laserfiche.RepositoryAccess
Imports Laserfiche.ClientAutomation
Imports System.Threading
Imports Laserfiche.DocumentServices

Public Class Form1
    Dim DocumentType       'Attrib01
    Dim LFSuffix           'Attrib02
    Dim DocumentID         'Attrib03
    Dim Branch             'Attrib04
    Dim InsurerCode        'Attrib05
    Dim DocumentNumber     'Attrib06
    Dim SubscriberCode     'Attrib07
    Dim PDFFilePath        'Attrib08

    Dim FolName As String = ""
    Dim DocName As String = ""

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Log("-----------------------------------")
        Log(Now)
        'Read Parameters


        Try
            DocumentType = System.Environment.GetCommandLineArgs(1)
            LFSuffix = System.Environment.GetCommandLineArgs(2)
            DocumentID = System.Environment.GetCommandLineArgs(3)
            Branch = System.Environment.GetCommandLineArgs(4)
            InsurerCode = System.Environment.GetCommandLineArgs(5)
            DocumentNumber = System.Environment.GetCommandLineArgs(6)
            SubscriberCode = System.Environment.GetCommandLineArgs(7)
            Try
                PDFFilePath = System.Environment.GetCommandLineArgs(8)
            Catch
            End Try
            Log("DocumentType   = " & System.Environment.GetCommandLineArgs(1) & vbCrLf &
            "LFSuffix       = " & System.Environment.GetCommandLineArgs(2) & vbCrLf &
            "DocumentID     = " & System.Environment.GetCommandLineArgs(3) & vbCrLf &
            "Branch         = " & System.Environment.GetCommandLineArgs(4) & vbCrLf &
            "InsurerCode    = " & System.Environment.GetCommandLineArgs(5) & vbCrLf &
            "DocumentNumber = " & System.Environment.GetCommandLineArgs(6) & vbCrLf &
            "SubscriberCode = " & System.Environment.GetCommandLineArgs(7))

            Try
                Log("PDFFilePath = " & System.Environment.GetCommandLineArgs(8))
            Catch ex As Exception
            End Try





        Catch ex As Exception

                Log(Now & " Error reading parameters" & vbCrLf &
            "DocumentType   = " & DocumentType & vbCrLf &
            "LFSuffix       = " & LFSuffix & vbCrLf &
            "DocumentID     = " & DocumentID & vbCrLf &
            "Branch         = " & Branch & vbCrLf &
            "InsurerCode    = " & InsurerCode & vbCrLf &
            "DocumentNumber = " & DocumentNumber & vbCrLf &
            "SubscriberCode = " & SubscriberCode & vbCrLf &
             "PDFFilePath = " & PDFFilePath)


            If UCase(DocumentType) = "CLAIM" Or UCase(DocumentType) = "TITLE DEAD" Then
                'its OK, no need to log error and can continue
            Else
                MsgBox("Error reading parameters" & vbCrLf & Err.Description)
                MsgBox("DocumentType = " & DocumentType & vbCrLf &
                                "LFSuffix = " & LFSuffix & vbCrLf &
                                "DocumentID = " & DocumentID & vbCrLf &
                                "Branch = " & Branch & vbCrLf &
                                "InsurerCode = " & InsurerCode & vbCrLf &
                                "DocumentNumber = " & DocumentNumber & vbCrLf &
                                "SubscriberCode = " & SubscriberCode)
                End
            End If


        End Try

        Try
            'ReadParam
            ReadParam(DocumentType)
        Catch ex As Exception
            MsgBox("Error While Reading Param, Error: " & ex.Message)
        End Try


        Try
            'Get Target Document
            'Find Destination Folder
            FolName = ReplaceTokens(LFDestinationFolder)
            Log(Now & " DestFolder = " & FolName)

            'Select Document Name
            DocName = ReplaceTokens(LFDocumentName)

            Log(Now & " DocName = " & DocName)
        Catch ex As Exception
            MsgBox("Error While preparing Document And Folder Names, Error: " & ex.Message)
        End Try

        If UCase(DocumentType) = "TITLE DEAD" Then
            Dim finalsearch As String = ""
            Try


                For Each str As String In DocumentID.split(";")
                    'Replace(LFSearch, "Attrib01", DocumentType)
                    finalsearch = finalsearch + Replace(LFSearch, "Attrib03", str) + " | "
                Next
                finalsearch = finalsearch.Substring(0, finalsearch.Length - 3)
                finalsearch = "(" + finalsearch + " ) & ({LF:Ext=""*""} | {LF:pagecount > 0})"
                'MsgBox(finalsearch)
                Log(finalsearch)
            Catch ex As Exception
                MsgBox("Error While preparing Search syntax, Error: " & ex.Message)
            End Try

            If UCase(LFSuffix) = "VIEW ALL" Then
                Try
                    GetLFMainWindow("LaunchSearch", finalsearch)
                Catch ex As Exception
                    MsgBox("Error While launching Seach, Error: " & ex.Message)
                End Try
            Else
                Dim repository As RepositoryRegistration
                Try
                    ' log into the repository
                    repository = New RepositoryRegistration(LFServerName, LFRepName)
                    session.LogIn(LFUserName, LFUserPassw, repository)

                Catch ex As Exception
                    MsgBox("Error While opening LF Session, Error: " & ex.Message)
                End Try
                Dim LFDocID As Integer
                Try
                    Dim fieldName As String = FieldsList(0).Split("=")(0).Trim

                    Using searchResultListing As SearchResultListing = SearchInLF(finalsearch, fieldName)

                        If searchResultListing.RowsCount Then

                            Dim maxPolicyNb As Integer = CInt(searchResultListing.GetDatumAsString(1, fieldName))
                            LFDocID = searchResultListing.GetDatumAsString(1, SystemColumn.Id)
                            'For i As Integer = 1 To searchResultListing.RowsCount

                            '    Dim currentNb = CInt(searchResultListing.GetDatumAsString(i, fieldName))
                            '    If currentNb > maxPolicyNb Then
                            '        maxPolicyNb = currentNb
                            '        LFDocID = searchResultListing.GetDatumAsString(i, SystemColumn.Id)
                            '    End If
                            'Next

                        Else
                            LFDocID = CreateDocument()

                        End If


                    End Using
                Catch ex As Exception
                    MsgBox("Error While retreiving latest title deed, Error: " & ex.Message)
                End Try
                Try
                    session.Close()
                Catch ex As Exception
                    MsgBox("Error While Closing LF Session, Error: " & ex.Message)
                End Try

                Try
                    GetLFMainWindow("OpenDocumentById", LFDocID)
                Catch ex As Exception
                    MsgBox("Error While Opening document DocID " & LFDocID & " , Error: " & ex.Message)
                End Try
            End If


        Else




            Dim repository As RepositoryRegistration
            Try
                ' log into the repository
                repository = New RepositoryRegistration(LFServerName, LFRepName)
                session.LogIn(LFUserName, LFUserPassw, repository)

            Catch ex As Exception
                MsgBox("Error While opening LF Session, Error: " & ex.Message)
            End Try








            If LFSearch <> "" Then
                Try
                    Using searchResultListing As SearchResultListing = SearchInLF(ReplaceTokens(LFSearch))
                        If searchResultListing.RowsCount Then
                            LFDocID = searchResultListing.GetEntryInfo(1).Id
                        End If
                    End Using
                Catch ex As Exception
                    MsgBox("Error While searching, Error: " & ex.Message)
                End Try

            End If



            If LFDocID = 0 Then
                Try
                    Dim docinfo As DocumentInfo = Document.GetDocumentInfo(FolName & "\" & DocName, session)

                    LFDocID = docinfo.Id
                    docinfo.Dispose()
                Catch ex As Exception
                    If LFUserName = "VIEWONLY" Then
                        MsgBox("Sorry but this document does not exist in Laserfiche")
                        session.Close()
                        End
                    Else
                        Try
                            LFDocID = CreateDocument()
                        Catch ex1 As Exception
                            MsgBox("Error While creating document, Error: " & ex1.Message)
                        End Try
                    End If
                End Try
            End If
        End If




        If Not String.IsNullOrEmpty(PDFFilePath) Then
            Try
                ImportEDocument(LFDocID)
                GetLFMainWindow("OpenDocumentById", LFDocID)
                'MsgBox("Document imported successfully")
            Catch ex As Exception
                MsgBox("Error While importing Document. Error: " & ex.Message)
            End Try
        Else
            Try
                GetLFMainWindow("OpenDocumentById", LFDocID)
            Catch ex As Exception
                MsgBox("Error While opening Document. Error: " & ex.Message)
            End Try
            End If




            Try
            session.Close()
        Catch ex As Exception
            MsgBox("Error While Closing LF Session, Error: " & ex.Message)
        End Try



        End



    End Sub


    Public Function ReplaceTokens(ByVal str As String) As String
        str = Replace(str, "Attrib01", DocumentType)
        str = Replace(str, "Attrib02", LFSuffix)
        str = Replace(str, "Attrib03", DocumentID)
        str = Replace(str, "Attrib04", Branch)
        str = Replace(str, "Attrib05", InsurerCode)
        str = Replace(str, "Attrib06", DocumentNumber)
        str = Replace(str, "Attrib07", SubscriberCode)
        Return str
    End Function

    Public Sub GetLFMainWindow(ByVal task As String, ByVal Details As String)
        Dim mainwindow As MainWindow
        Dim LFisOpen As Boolean = False
        Using lfclient As New ClientManager()
            Dim clients As IEnumerable(Of ClientInstance) = lfclient.GetAllClientInstances()
            For Each client As ClientInstance In clients

                Dim windows As IEnumerable(Of ClientWindow) = client.GetAllClientWindows()
                For Each window As ClientWindow In windows
                    If window.GetWindowType() = ClientWindowType.Main Then
                        mainwindow = DirectCast(window, MainWindow)
                        If mainwindow.IsLoggedIn Then
                            If mainwindow.GetCurrentRepository.RepositoryName.ToLower = LFRepName.ToLower Then
                                LFisOpen = True
                                Exit For
                            End If
                        End If
                    End If
                Next
            Next

            If LFisOpen = False Then
                Dim options As New LaunchOptions
                options.ServerName = LFServerName
                options.RepositoryName = LFRepName
                options.UserName = LFUserName
                options.Password = LFUserPassw
                options.HiddenWindow = False

                'options.InitialFolderId = Details



                Dim clientinstance As ClientInstance = lfclient.LaunchClient(options)

                Dim windows As IEnumerable(Of ClientWindow) = clientinstance.GetAllClientWindows()
                For Each window As ClientWindow In windows
                    If window.GetWindowType() = ClientWindowType.Main Then
                        mainwindow = DirectCast(window, MainWindow)
                    End If

                Next
            End If
            If task = "OpenDocumentById" Then
                Dim openOptions As New OpenOptions
                openOptions.OpenStyle = DocumentOpenType.Default
                mainwindow.OpenDocumentById(Details, openOptions)
                If LFisOpen = False Then
                    mainwindow.Close()

                End If
            ElseIf task = "LaunchScanningFromClient" Then
                Dim openOptions As New OpenOptions
                openOptions.OpenStyle = DocumentOpenType.Default
                mainwindow.OpenDocumentById(Details, openOptions)
                Dim Scanoptions As New ScanOptions()
                Scanoptions.EntryId = Details
                Scanoptions.ScanMode = ScanMode.Standard
                Scanoptions.InsertPagesAt = CInt(InsertAt.[End])
                Scanoptions.WaitForExit = True
                Scanoptions.CloseAfterStoring = True
                mainwindow.LaunchScanningFromClient(Scanoptions)
                Dim doc As DocumentInfo = New DocumentInfo(Details, session)
                If doc.PageCount = 0 Then
                    doc.Delete()
                End If
                doc.Dispose()

                If LFisOpen = False Then
                    mainwindow.Close()
                End If

            Else
                Dim searchoptions As New SearchOptions()
                searchoptions.Query = Details
                searchoptions.NewWindow = False
                searchoptions.OpenIfOneResult = True
                mainwindow.LaunchSearch(searchoptions)
                mainwindow.SetFocus()
            End If
        End Using

    End Sub

    Public Function SearchInLF(ByVal Query As String, Optional ByVal FieldName As String = "") As SearchResultListing

        ' initialize an instance of the Search class
        Using search As New Search(session)
            ' specifiy the search query
            search.Command = Query
            ' wait until the search completes on the server
            Dim longOp As LongOperation = search.BeginRun(False)
            While Not longOp.IsCompleted
                Thread.Sleep(1000)
                search.UpdateStatus()
            End While

            ' specify the settings to use when retrieving the search results,
            ' such as the entry type filter and columns.
            Dim searchSetting As New SearchListingSettings()
            searchSetting.EntryFilter = EntryTypeFilter.AllTypes
            searchSetting.AddColumn(SystemColumn.Id)
            searchSetting.AddColumn(SystemColumn.Name)

            If FieldName <> "" Then
                searchSetting.AddColumn(FieldName)
                searchSetting.SetSortColumn(FieldName, Laserfiche.RepositoryAccess.SortDirection.Descending)
            End If


            Return search.GetResultListing(searchSetting)
            ' get the search result listing and iterate through the rows
        End Using
    End Function



    Public Sub ImportEDocument(ByVal docID As Integer)

        Dim docInfo As DocumentInfo = Document.GetDocumentInfo(docID, session)
        Dim DI As New DocumentImporter()
        DI.Document = docInfo
        DI.OcrImages = True
        DI.ImportEdoc("application/pdf", Trim(PDFFilePath))
        docInfo.Save()
        docInfo.Dispose()
    End Sub




    Public Function CreateDocument() As Integer
        Dim docInfo As New DocumentInfo(session)
        Using parentFol As FolderInfo = CreateFolderByPath(FolName)
            docInfo.Create(parentFol, DocName, LFVolumeName, EntryNameOption.AutoRename)
        End Using

        LFDocID = docInfo.Id
        docInfo.SetTemplate(LFTemplateName)
        Dim fielddata As FieldValueCollection = docInfo.GetFieldValues()
        Dim FieldName As String = ""
        Dim FieldValue As String = ""
        Try

            Log(Now & " Preparing Fields Values")
            If FieldsList.Count > 0 Then

                For Each fieldSeting As String In FieldsList

                    FieldName = fieldSeting.Split("=")(0).Trim
                    FieldValue = ReplaceTokens(fieldSeting.Split("=")(1).Trim)

                    fielddata(FieldName) = FieldValue
                Next
            End If
        Catch ex As Exception
            Log(Now & " Error while setting Fields Values = " & Err.Description)
        End Try
        docInfo.SetFieldValues(fielddata)
        docInfo.Save()
        docInfo.Dispose()
        Return LFDocID
    End Function
End Class
