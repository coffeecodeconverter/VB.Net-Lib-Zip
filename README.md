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

Function SelectFolder(Optional title As String = "Select Folder") As String
