Imports System.IO
Imports Laserfiche.RepositoryAccess
Imports Laserfiche.ClientAutomation
Imports System.Linq

Module General


    Public session As New Session
    Public ClientConnected As Boolean = False
    Public LFDocID As Long

    Public LFServerName As String
    Public LFRepName As String
    Public LFUserName As String
    Public LFUserPassw As String


    Public LFDestinationFolder As String
    Public LFDocumentName As String
    Public LFVolumeName As String
    Public LFTemplateName As String
    Public LFSearch As String
    Public NewType As String

    ''''Modification applied on 20180619
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Public FieldsList As New List(Of String)

    ''' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




    Public CountriesList As New ArrayList
    Public AccountMgrsList As New ArrayList
    Public LFAccountManager As String

    Public ClientWindow As ClientWindow

    Public ClientManager As ClientManager

    Public RepositoryConnection As RepositoryConnection
    Public Sub Log(ByVal text As String)
        Dim sr As New StreamWriter(Application.StartupPath & "\Logfile.log", True)
        sr.WriteLine(text)
        sr.Close()
    End Sub

    Public Function ReadParam(ByVal DocumentType As String)
        Dim a As Array
        Dim i As Integer
        Dim line As String
        Try
            Dim sr As New StreamReader(Application.StartupPath & "\Param.ini")

            a = Nothing
            a = Split(sr.ReadLine, "=", -1, CompareMethod.Text)
            LFServerName = Trim(a(1).trim)

            a = Nothing
            a = Split(sr.ReadLine, "=", -1, CompareMethod.Text)
            LFRepName = Trim(a(1).trim)

            a = Nothing
            a = Split(sr.ReadLine, "=", -1, CompareMethod.Text)
            LFUserName = Trim(a(1).trim)

            a = Nothing
            a = Split(sr.ReadLine, "=", -1, CompareMethod.Text)
            LFUserPassw = DecryptPassword(Trim(a(1).trim))
            sr.Close()
        Catch ex As Exception
            MsgBox("Sorry, cannot read Param.ini. Please check this file")
            End
        End Try
        Dim sr_DocType As StreamReader
        Try
            Try
                sr_DocType = New StreamReader(Application.StartupPath & "\DocumentTypes\" & DocumentType & ".ini")
            Catch ex As Exception
                Try
                    sr_DocType = New StreamReader(Application.StartupPath & "\DocumentTypes\Default.ini")
                Catch ex2 As Exception
                    MsgBox("Sorry, cannot read DocumentTypes\" & DocumentType & ".ini. Please check this file")
                    End
                End Try

            End Try
            a = Nothing
            a = Split(sr_DocType.ReadLine, "=", -1, CompareMethod.Text)
            LFDestinationFolder = Trim(a(1).trim)
            'MsgBox(LFDestinationFolder)
            a = Nothing
            a = Split(sr_DocType.ReadLine, "=", -1, CompareMethod.Text)
            LFDocumentName = Trim(a(1).trim)
            'MsgBox(LFDocumentName)
            a = Nothing
            a = Split(sr_DocType.ReadLine, "=", -1, CompareMethod.Text)
            LFVolumeName = Trim(a(1).trim)
            'MsgBox(LFVolumeName)
            a = Nothing
            a = Split(sr_DocType.ReadLine, "=", -1, CompareMethod.Text)
            LFTemplateName = Trim(a(1).trim)
            'MsgBox(LFTemplateName)


            a = Nothing
            a = Split(sr_DocType.ReadLine, ";", -1, CompareMethod.Text)
            LFSearch = Trim(a(1).trim)
            ' MsgBox(LFSearch)


            sr_DocType.ReadLine()    ' ''''''''''''''''''''''
            sr_DocType.ReadLine()   ' '''Fields Settings''''
            sr_DocType.ReadLine()   ' ''''''''''''''''''''''

            Do While sr_DocType.Peek() >= 0
                FieldsList.Add(sr_DocType.ReadLine)

            Loop

            sr_DocType.Close()
        Catch ex As Exception
            MsgBox("Sorry, error while reading " & DocumentType & ".ini. Please check this file")
            Log("Sorry, error while reading " & DocumentType & ".ini. Please check this file")
            End
        End Try
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    End Function
    Public Function DecryptPassword(ByVal EP As String) As String
        Dim i As Integer
        Dim PWD As String
        PWD = ""
        For i = 1 To Len(EP)
            PWD = PWD & Chr((Asc(Mid(EP, i, 1)) - i) - 12)
        Next
        DecryptPassword = PWD
    End Function

    Public Function CreateFolderByPath(ByVal folderpath As String) As FolderInfo


        Dim f As FolderInfo
        Try
            f = Folder.GetFolderInfo(folderpath, session)
        Catch ex As Exception
            Dim lastindex As Integer = folderpath.LastIndexOf("\"c)
            f = New FolderInfo(session)
            Dim parentfol As FolderInfo = CreateFolderByPath(folderpath.Substring(0, lastindex))
            f.Create(folderpath, EntryNameOption.None)
            parentfol.Dispose()
            f.Save()
        End Try

        Return f
    End Function

    Public Function CreateFolderByPathPDF(ByVal folderpath As String) As String
        ' Check if the path exists
        If Not Directory.Exists(folderpath) AndAlso Not File.Exists(folderpath) Then
            Throw New DirectoryNotFoundException($"Path '{folderpath}' not found.")
        End If

        ' Get the last directory or file name
        Dim name As String = Path.GetFileName(folderpath)
        Dim fileNameWithoutExtension As String = Path.GetFileNameWithoutExtension(name)
        Return fileNameWithoutExtension
    End Function

    Public Function CreateFolderPDF(ByVal folderpath As String) As FolderInfo

        Dim f As FolderInfo
        Try
            f = Folder.GetFolderInfo(folderpath, session)
        Catch ex As Exception
            f = New FolderInfo(session)
            f.Create(folderpath, EntryNameOption.None)
            f.Dispose()
            f.Save()

        End Try
        Return f
    End Function




End Module
