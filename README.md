# VB.Net-Lib-Zip
VB.Net Lib - Zip / Unzip Asynchronously, with progress, and custom open file/folder dialogs for ease and flexibilitiy

---

## Requirements

- .NET Framework 4.5.1 or higher  
- System.IO.Compression (built-in)  
- Windows API Code Pack (for enhanced folder dialogs)  

---

## Description

A comprehensive, asynchronous Zipping and Unzipping module that simplifies file compression and extraction with full progress reporting and robust error handling.  
Includes built-in customizable fallback dialogs for file and folder selection, enabling seamless user interaction across different UI environments.

---

## Key Features

- Asynchronous zipping and unzipping with real-time progress percentage updates  
- Supports selection of files and folders for compression, with optional dialog fallback if no parameters provided  
- Handles both file and directory compression, preserving folder structure during extraction  
- Gracefully manages errors and cancellation scenarios with clear status messages  
- Integrated custom file/folder dialog support for easy file system browsing and selection  
- Optional callback support for post-operation hooks, ideal for UI updates or logging  
- Designed for effortless integration into any .NET project requiring reliable archive management  

---

## Main Functions

```vbnet
Await ZipAsync(Optional itemsToZip As IEnumerable(Of String) = Nothing,
               Optional zipFilePath As String = Nothing,
               Optional progress As IProgress(Of Integer) = Nothing) As Task(Of String)

Await UnzipAsync(Optional zipFilePath As String = Nothing,
                 Optional extractToFolder As String = Nothing,
                 Optional progress As IProgress(Of Integer) = Nothing) As Task(Of String)

Function SelectFiles(Optional multiSelect As Boolean = True,
                                Optional filter As String = "All files (*.*)|*.*",
                                Optional title As String = "Select File(s)") As List(Of String)

Function SelectFolder(Optional title As String = "Select Folder") 
```


## Full Module
```vbnet


Imports System.IO
Imports System.IO.Compression
Imports System.Threading
Imports System.Threading.Tasks
Imports Microsoft.WindowsAPICodePack.Dialogs



Module ModuleZipper


    Public ModuleVersion_Zipper As String = "1.0.0.1"


    ' v4 
    Public Async Function ZipAsync(Optional itemsToZip As IEnumerable(Of String) = Nothing,
                               Optional zipFilePath As String = Nothing,
                               Optional progress As IProgress(Of Integer) = Nothing, Optional callback As Action = Nothing) As Task(Of String)

        Try
            Debug.WriteLine("===== ZIP DEBUG START =====")

            ' Prompt user to select files if none provided
            If itemsToZip Is Nothing OrElse Not itemsToZip.Any() Then
                itemsToZip = SelectFiles(True, "All files (*.*)|*.*", "Select files or folders to ZIP")
                Debug.WriteLine($"User selected {itemsToZip.Count()} item(s) to zip.")
                If Not itemsToZip.Any() Then Return "ZIP cancelled: No input selected."
            Else
                Debug.WriteLine($"ItemsToZip passed in as parameter: {String.Join(";", itemsToZip)}")
            End If

            ' Default filename suggestion
            Dim defaultFileName = "Archive_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".zip"

            ' Prompt for zip destination if none was provided
            If String.IsNullOrWhiteSpace(zipFilePath) Then
                Using sfd As New SaveFileDialog()
                    sfd.Title = "Select ZIP output location"
                    sfd.Filter = "ZIP files (*.zip)|*.zip"
                    sfd.DefaultExt = "zip"
                    sfd.AddExtension = True
                    sfd.FileName = defaultFileName
                    If sfd.ShowDialog() = DialogResult.OK Then
                        zipFilePath = sfd.FileName
                    Else
                        Return "ZIP cancelled: No destination selected."
                    End If
                End Using
            End If

            ' === Ensure zipFilePath is valid and includes a .zip file name ===
            Dim fileNamePart As String = Path.GetFileName(zipFilePath)

            If String.IsNullOrWhiteSpace(fileNamePart) OrElse Not fileNamePart.ToLower().EndsWith(".zip") Then
                Debug.WriteLine("Provided path is missing a valid ZIP filename or has the wrong extension.")

                ' Get directory or fallback to current folder
                Dim dirPath As String = If(Directory.Exists(zipFilePath),
                                       zipFilePath,
                                       If(Not String.IsNullOrWhiteSpace(Path.GetDirectoryName(zipFilePath)),
                                          Path.GetDirectoryName(zipFilePath),
                                          Environment.CurrentDirectory))

                ' Append default filename
                zipFilePath = Path.Combine(dirPath, defaultFileName)
                Debug.WriteLine($"Auto-corrected ZIP path: {zipFilePath}")
            End If

            Debug.WriteLine($"ZIP output path: {zipFilePath}")

            ' Run zipping operation on background thread
            Return Await Task.Run(Function()
                                      Try
                                          ' Overwrite existing ZIP if it exists
                                          If File.Exists(zipFilePath) Then
                                              Debug.WriteLine($"Existing ZIP file found, deleting: {zipFilePath}")
                                              File.Delete(zipFilePath)
                                          End If

                                          Using archive As ZipArchive = ZipFile.Open(zipFilePath, ZipArchiveMode.Create)
                                              Dim totalItems = itemsToZip.Count()
                                              Dim currentItem = 0

                                              For Each item In itemsToZip
                                                  Debug.WriteLine($"Processing item: {item}")

                                                  If File.Exists(item) Then
                                                      AddFileToArchive(item, archive, Path.GetFileName(item))
                                                  ElseIf Directory.Exists(item) Then
                                                      AddDirectoryToArchive(item, archive, Path.GetFileName(item))
                                                  Else
                                                      Debug.WriteLine($"Item not found or invalid: {item}")
                                                  End If

                                                  currentItem += 1
                                                  progress?.Report(CInt(currentItem / totalItems * 100))
                                              Next
                                          End Using

                                          Debug.WriteLine("===== ZIP COMPLETE =====")

                                          Dim syncContext = SynchronizationContext.Current

                                          ' After unzip completes:
                                          If callback IsNot Nothing Then
                                              If syncContext IsNot Nothing Then
                                                  syncContext.Post(Sub() callback(), Nothing)
                                              Else
                                                  callback()
                                              End If
                                          End If

                                          Return "Zipping completed successfully."

                                      Catch zipEx As Exception
                                          Debug.WriteLine("ZIP TASK ERROR: " & zipEx.ToString())
                                          Return "ZIP failed during compression: " & zipEx.Message
                                      End Try
                                  End Function)

        Catch ex As Exception
            Debug.WriteLine("ZIP ERROR OUTSIDE TASK: " & ex.ToString())
            Return "ZIP failed: " & ex.Message
        End Try
    End Function






    ' V3 
    Public Async Function UnzipAsync(Optional zipFilePath As String = Nothing,
                                 Optional extractToFolder As String = Nothing,
                                 Optional progress As IProgress(Of Integer) = Nothing, Optional callback As Action = Nothing) As Task(Of String)
        Try
            Debug.WriteLine("===== UNZIP DEBUG START =====")

            If String.IsNullOrWhiteSpace(zipFilePath) Then
                Dim file = SelectFiles(False, "ZIP files (*.zip)|*.zip", "Select ZIP to extract").FirstOrDefault()
                Debug.WriteLine($"Selected ZIP file: {file}")
                If String.IsNullOrWhiteSpace(file) Then
                    Return "Unzip cancelled: No ZIP file selected."
                End If
                zipFilePath = file
            Else
                Debug.WriteLine($"ZIP file path (param): {zipFilePath}")
            End If

            If Not zipFilePath.EndsWith(".zip", StringComparison.OrdinalIgnoreCase) Then
                Debug.WriteLine("File selected is not a .zip: " & zipFilePath)
                Return "Unzip cancelled: Selected file is not a ZIP archive."
            End If

            If Not File.Exists(zipFilePath) Then
                Debug.WriteLine($"ZIP file not found: {zipFilePath}")
                Return "ZIP file not found: " & zipFilePath
            End If

            If String.IsNullOrWhiteSpace(extractToFolder) Then
                extractToFolder = SelectFolder("Select folder to extract files to")
                Debug.WriteLine($"Selected output folder: {extractToFolder}")
                If String.IsNullOrWhiteSpace(extractToFolder) Then Return "Unzip cancelled: No destination selected."
            Else
                Debug.WriteLine($"Output folder path (param): {extractToFolder}")
            End If

            If Not Directory.Exists(extractToFolder) Then
                Debug.WriteLine($"Creating output folder: {extractToFolder}")
                Directory.CreateDirectory(extractToFolder)
            End If

            ' Run the CPU-heavy zipping in background thread
            Return Await Task.Run(Function()
                                      Debug.WriteLine($"Reading archive: {zipFilePath}")
                                      Using archive As ZipArchive = ZipFile.OpenRead(zipFilePath)
                                          Dim total = archive.Entries.Count
                                          Debug.WriteLine($"Total entries: {total}")
                                          Dim index = 0

                                          For Each entry In archive.Entries
                                              Debug.WriteLine($"Processing entry: {entry.FullName}")

                                              Dim fullPath = Path.Combine(extractToFolder, entry.FullName)
                                              Dim dirPath = Path.GetDirectoryName(fullPath)
                                              If Not Directory.Exists(dirPath) Then
                                                  Debug.WriteLine($"Creating directory: {dirPath}")
                                                  Directory.CreateDirectory(dirPath)
                                              End If

                                              If Not String.IsNullOrEmpty(entry.Name) Then
                                                  Debug.WriteLine($"Extracting file to: {fullPath}")
                                                  entry.ExtractToFile(fullPath, overwrite:=True)
                                              Else
                                                  Debug.WriteLine($"Skipping directory entry: {entry.FullName}")
                                              End If

                                              index += 1
                                              progress?.Report(CInt(index / total * 100))
                                          Next
                                      End Using

                                      Debug.WriteLine("===== UNZIP COMPLETE =====")

                                      Dim syncContext = SynchronizationContext.Current

                                      ' After unzip completes:
                                      If callback IsNot Nothing Then
                                          If syncContext IsNot Nothing Then
                                              syncContext.Post(Sub() callback(), Nothing)
                                          Else
                                              callback()
                                          End If
                                      End If

                                      Return "Unzipping completed successfully."
                                  End Function)
        Catch ex As Exception
            Debug.WriteLine("UNZIP ERROR: " & ex.ToString())
            Return "Unzip failed: " & ex.Message
        End Try
    End Function





    Private Sub AddFileToArchive(filePath As String, archive As ZipArchive, entryName As String)
        Try
            archive.CreateEntryFromFile(filePath, entryName, CompressionLevel.Optimal)
        Catch ex As Exception
            Throw New Exception("Error adding file: " & filePath & " - " & ex.Message)
        End Try
    End Sub



    Private Sub AddDirectoryToArchive(folderPath As String, archive As ZipArchive, baseFolderName As String)
        Try
            Dim files = Directory.GetFiles(folderPath, "*.*", SearchOption.AllDirectories)

            For Each file In files
                Dim relativePath = Path.Combine(baseFolderName, GetRelativePath(folderPath, file))
                archive.CreateEntryFromFile(file, relativePath, CompressionLevel.Optimal)
            Next
        Catch ex As Exception
            Throw New Exception("Error zipping folder: " & folderPath & " - " & ex.Message)
        End Try
    End Sub



    Private Function GetRelativePath(basePath As String, fullPath As String) As String
        Dim baseUri = New Uri(If(basePath.EndsWith("\"), basePath, basePath & "\"))
        Dim fileUri = New Uri(fullPath)
        Return Uri.UnescapeDataString(baseUri.MakeRelativeUri(fileUri).ToString()).Replace("/", "\")
    End Function





    Public Function SelectFiles(Optional multiSelect As Boolean = True,
                                Optional filter As String = "All files (*.*)|*.*",
                                Optional title As String = "Select File(s)") As List(Of String)
        Try
            Using dlg As New OpenFileDialog()
                dlg.Multiselect = multiSelect
                dlg.Filter = filter
                dlg.Title = title

                If dlg.ShowDialog() = DialogResult.OK Then
                    Return dlg.FileNames.ToList()
                Else
                    Return New List(Of String)
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show("An error occurred while opening the file dialog: " & ex.Message, "Dialog Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return New List(Of String)
        End Try
    End Function



    ' custom FolderDialog from 'Windows API Code Pack' - behaves like an openFileDilaog, same look and feel, but for folders. 
    Public Function SelectFolder(Optional title As String = "Select Folder") As String
        Try
            Using dialog As New CommonOpenFileDialog()
                dialog.Title = title
                dialog.IsFolderPicker = True
                dialog.AllowNonFileSystemItems = False
                dialog.EnsurePathExists = True
                dialog.EnsureReadOnly = False
                dialog.EnsureValidNames = True
                dialog.Multiselect = False

                If dialog.ShowDialog() = CommonFileDialogResult.Ok Then
                    Return dialog.FileName
                Else
                    Return Nothing
                End If
            End Using
        Catch ex As Exception
            Return Nothing
        End Try
    End Function



End Module
```

