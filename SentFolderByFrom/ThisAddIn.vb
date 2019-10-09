Public Class ThisAddIn

    Private OSentItemsFolder As Outlook.Folder
    Public WithEvents OSentItems As Outlook.Items
    Private mapiNameSpace As Outlook.NameSpace
    Private primaryEmail As String
    Private SentItemFolders As Dictionary(Of String, String)
    Private binarify As New Runtime.Serialization.Formatters.Binary.BinaryFormatter()
    Private UserConfigFolder As String = Environment.GetEnvironmentVariable("appdata") & "\browlry\SentFolderByFrom"
    Private UserConfigPath As String = UserConfigFolder & "\userconfig.bin"

    'When Outlook starts:
    Private Sub ThisAddIn_Startup() Handles Me.Startup
        ' Create the folder for storing user settings, if it doesn't exist
        If (Not System.IO.Directory.Exists(UserConfigFolder)) Then
            System.IO.Directory.CreateDirectory(UserConfigFolder)
        End If
        primaryEmail = Application.Session.CurrentUser.Address
        'Monitor the messages in the Sent Items folder; trigger OSentItems_ItemAdd when a new message is added
        mapiNameSpace = Application.GetNamespace("MAPI")
        OSentItemsFolder = mapiNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail)
        OSentItems = OSentItemsFolder.Items
        AddHandler OSentItems.ItemAdd, AddressOf OSentItems_ItemAdd
        'Load the user settings file into the SentItemsFolders dictionary, or create the dictionary if the file doesn't exist or is empty.
        If IO.File.Exists(UserConfigPath) Then
            Dim fsRead As New IO.FileStream(UserConfigPath, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.None)
            If fsRead.Length > 0 Then
                SentItemFolders = binarify.Deserialize(fsRead)
            Else
                SentItemFolders = New Dictionary(Of String, String)
            End If
            fsRead.Close()
        Else
            SentItemFolders = New Dictionary(Of String, String)
        End If
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    Private Sub OSentItems_ItemAdd(ByVal myItem As Object) 'specifies the actions when a new item is added to the Sent Items folder
        'Determine if the item sent is MailItem
        If TypeName(myItem) = "MailItem" Then
            Call MoveToNewSentFolder(myItem) 'calls the ChangeSentFolder function when a mail item is sent
        End If
    End Sub

    Private Sub MoveToNewSentFolder(myItem As Object)
        ' Inspired by https://www.itprotoday.com/management-mobility/outlook-2010-move-mailitems-after-sending
        Dim NewSentItemsFolder As Outlook.Folder
        Dim NewSentItemsPath As String
        Dim fromAddress As String
        'Identify sender address
        If myItem.Sender IsNot Nothing Then
            fromAddress = myItem.Sender.Address.ToString
            ' If the item is not sent from the primary email address...
            If (fromAddress <> primaryEmail) Then
                ' See if there is a saved "Sent" folder for that address, and if not
                If Not SentItemFolders.ContainsKey(fromAddress) Then
                    ' Prompt the user to choose a folder for the sent items.
                    System.Windows.Forms.MessageBox.Show("Click OK to select the 'Sent Items' folder for items sent from " & fromAddress)
                    NewSentItemsFolder = mapiNameSpace.PickFolder
                    'If the user doesn't pick anything, stop
                    If NewSentItemsFolder Is Nothing Then
                        Exit Sub
                    End If
                    'Save the path of the folder chosen for future reference.
                    NewSentItemsPath = NewSentItemsFolder.FolderPath
                    SentItemFolders.Add(fromAddress, NewSentItemsPath)
                    Dim fs As IO.FileStream = New IO.FileStream(UserConfigPath, IO.FileMode.Create, IO.FileAccess.Write, IO.FileShare.None)
                    binarify.Serialize(fs, SentItemFolders)
                    fs.Close()
                Else
                    ' If there is a "Sent" folder associated with that email aleady, retrieve that folder. 
                    NewSentItemsPath = SentItemFolders(fromAddress)
                    NewSentItemsFolder = GetFolder(NewSentItemsPath)
                End If
                ' Move the mail item to the appropriate folder.
                myItem.Move(NewSentItemsFolder)
            End If
        End If
    End Sub

    Function GetFolder(ByVal FolderPath As String) As Outlook.Folder
        'This function courtesy of https://docs.microsoft.com/en-us/office/vba/outlook/how-to/items-folders-and-stores/obtain-a-folder-object-from-a-folder-path
        Dim TestFolder As Outlook.Folder
        Dim FoldersArray As Object
        Dim i As Integer

        On Error GoTo GetFolder_Error
        If Left(FolderPath, 2) = "\\" Then
            FolderPath = Right(FolderPath, Len(FolderPath) - 2)
        End If
        'Convert folderpath to array 
        FoldersArray = Split(FolderPath, "\")
        TestFolder = Application.Session.Folders.Item(FoldersArray(0))
        If Not TestFolder Is Nothing Then
            For i = 1 To UBound(FoldersArray, 1)
                Dim SubFolders As Outlook.Folders
                SubFolders = TestFolder.Folders
                TestFolder = SubFolders.Item(FoldersArray(i))
                If TestFolder Is Nothing Then
                    GetFolder = Nothing
                End If
            Next
        End If
        'Return the TestFolder 
        GetFolder = TestFolder
        Exit Function

GetFolder_Error:
        GetFolder = Nothing
        Exit Function
    End Function

End Class