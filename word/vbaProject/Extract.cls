VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Extract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' For general file manipulation
Dim FSO As Object

' Helpers for pretty formatting xml files
Dim XML As Object
Dim PrettyXML As Object
Dim Writer As Object
Dim Reader As Object


' Prepare commonly used objects
Sub Initialize()

  Set FSO = CreateObject("Scripting.FilesystemObject")
  
  Set XML = CreateObject("MSXML2.DOMDocument.6.0")
  XML.async = False
  XML.resolveExternals = False
   
  ' Using DomDocument.6.0 doesn't work yet for reasons to be determined
  Set PrettyXML = CreateObject("MSXML2.DOMDocument.3.0")
  PrettyXML.async = False
  PrettyXML.resolveExternals = False
  
  Set Writer = CreateObject("MSXML2.MXXMLWriter.6.0")
  Set Reader = CreateObject("MSXML2.SAXXMLReader.6.0")
  Writer.Indent = True
  Writer.Encoding = "UTF-8"
  Writer.byteOrderMark = False
  Set Reader.contentHandler = Writer

End Sub


' Unzip active document to folder with same name
' For .bin files on vba is handled by directly exporting modules
Public Sub UnzipDocument()

  Dim ShellApp As Object: Set ShellApp = CreateObject("Shell.Application")
  Dim TmpZipFile As Variant
  Dim OutputFolderName As String
  Dim UnzipPath As Variant
  
  ' Verify file was saved and available to process
  If ActiveDocument.Path = vbNullString Then
    MsgBox "File must be saved before it can be processed"
    Exit Sub
  End If

  ' Copy document into temporary file recognized as archive
  TmpZipFile = ActiveDocument.FullName & ".zip"
  FSO.CopyFile ActiveDocument.FullName, TmpZipFile

  ' Document to be unzipped into folder with same name
  OutputFolderName = Left(ActiveDocument.Name, InStrRev(ActiveDocument.Name, ".") - 1)
  UnzipPath = FSO.BuildPath(ActiveDocument.Path, OutputFolderName)

  ' Create output folder if file/folder with same name does not exist
  If FSO.FolderExists(UnzipPath) = False Then
      FSO.CreateFolder UnzipPath
  End If

  ' Extract document contents
  ' FOF_SILENT + FOF_NOCONFIRMATION + FOF_NOERRORUI
  ShellApp.Namespace(UnzipPath).CopyHere ShellApp.Namespace(TmpZipFile).Items, (4 + 16 + 1024)
  ExportVBA CStr(UnzipPath)
  PrettyPrintXML CStr(UnzipPath)
  
  ' Cleanup
  FSO.DeleteFile TmpZipFile

End Sub

' Export all VBA components to specified folder
Sub ExportVBA(ByVal ExportPath As String)

  Dim VBC As Object
  Dim ComponentName As String

  ExportPath = FSO.BuildPath(FSO.BuildPath(ExportPath, "word"), "vbaProject")
  
  ' Create output folder if file/folder with same name does not exist
  If FSO.FolderExists(ExportPath) = False Then
      FSO.CreateFolder ExportPath
  End If
  
  For Each VBC In ActiveDocument.VBProject.VBComponents

    ComponentName = VBC.Name

    ' Select approrpiate file extension for export file
    ' Case for vbext_ct_ActiveXDesigner should not be needed
    Select Case VBC.Type
      Case 1 ' VBIDE.vbext_ct_StdModule
        ComponentName = ComponentName & ".bas"
      Case 2, 100 ' VBIDE.vbext_ct_ClassModule, VBIDE.vbext_ct_Document
        ComponentName = ComponentName & ".cls"
      Case 3 ' VBIDE.vbext_ct_MSForm
        ComponentName = ComponentName & ".frm"
    End Select
    
    VBC.Export FSO.BuildPath(ExportPath, ComponentName)
    
  Next VBC
  
End Sub

' Get listing of folder and all subfolders
Function GetFolders(Path As String) As Object()

  Dim Folder As Object
  Dim SubFolder As Object
  Dim Folders() As Object
  Dim EndIndex As Integer
  Dim CurrIndex As Integer
  
  ' Set root folder
  ReDim Folders(CurrIndex)
  Set Folders(EndIndex) = FSO.GetFolder(Path)

  ' Progressively lookup and record sub-folders
  Do While CurrIndex <= UBound(Folders)
  
    Set Folder = Folders(CurrIndex)
    ReDim Preserve Folders(EndIndex + Folder.SubFolders.Count)
  
    For Each SubFolder In Folder.SubFolders
      EndIndex = EndIndex + 1
      Set Folders(EndIndex) = SubFolder
    Next SubFolder
    
    CurrIndex = CurrIndex + 1
    
  Loop
  
  GetFolders = Folders
  
End Function

' Pretty print all xml files in given folder and its subfolders
' This method adds tabs when indenting, using manual XSLT would add unneeded closing tags
Sub PrettyPrintXML(Path As String)

  Dim Folders() As Object: Folders = GetFolders(Path)
  Dim Folder As Variant
  Dim File As Variant
  
  For Each Folder In Folders
  
    For Each File In Folder.Files
      If FSO.GetExtensionName(File.Path) = "xml" Then
        XML.Load File.Path
        Writer.output = PrettyXML
        Reader.parse XML
        PrettyXML.Save File.Path
      End If
    Next File
    
  Next Folder
  
End Sub


