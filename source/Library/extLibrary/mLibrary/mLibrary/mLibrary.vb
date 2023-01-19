Imports System
Imports System.IO
Imports System.Environment
Imports System.Windows.Forms
Imports Microsoft.Win32
Imports System.Collections.Generic
Imports System.Text
Imports System.Runtime.InteropServices
Imports System.Reflection
Imports System.Net.NetworkInformation

Public Class winLib
    Private Shared mselectedFiles() As String

    Public Shared Function BrowseForFolder(ByVal description As String, ByVal folder As String) As String
        Dim dialog As New FolderBrowserDialog()
        dialog.Description = description
        dialog.SelectedPath = folder
        dialog.ShowNewFolderButton = True
        If dialog.ShowDialog() = DialogResult.OK Then
            Return dialog.SelectedPath
        End If
        Return ""
    End Function

    ' Public Shared Sub CreateFolder(ByVal folderPath As String)
        ' Directory.CreateDirectory(folderPath)
    ' End Sub

    ' Public Shared Sub DeleteFile(ByVal filePath As String)
        ' File.Delete(filePath)
    ' End Sub

    ' Public Shared Function GetOpenFilesDlgFileName(ByVal file As Integer) As String
        ' If (mselectedFiles Is Nothing) OrElse (mselectedFiles.Length = 0) Then
            ' Return ""
        ' End If
        ' If mselectedFiles.Length < (file - 1) Then
            ' Return ""
        ' End If
        ' Return mselectedFiles(file - 1)
    ' End Function

    Public Shared Function GetOpenFilesDlgFileNames(ByVal files() As String) As Integer
        If (mselectedFiles Is Nothing) OrElse (mselectedFiles.Length = 0) Then
            Return 0
        End If
        Dim index As Integer = 0
        For Each str As String In mselectedFiles
            files(index) = str
            index += 1
        Next str
        Return files.Length
    End Function

    Public Shared Function OpenFilesDlg(ByVal title As String, ByVal startFolder As String, ByVal fileTypes As String) As Integer
        Dim dialog As New OpenFileDialog()
        dialog.Multiselect = True
        dialog.Title = title
        If (startFolder IsNot Nothing) AndAlso (startFolder.Length > 0) Then
            dialog.InitialDirectory = startFolder
        End If
        If (fileTypes IsNot Nothing) AndAlso (fileTypes.Length > 0) Then
            dialog.Filter = fileTypes
        End If
        If dialog.ShowDialog() = DialogResult.Cancel Then
            Return 0
        End If
        mselectedFiles = dialog.FileNames
        If (mselectedFiles Is Nothing) OrElse (mselectedFiles.Length = 0) Then
            Return 0
        End If
        Return mselectedFiles.Length
    End Function

    <DllImport("shell32.dll")> _
    Private Shared Function SHGetPathFromIDListW(ByVal pidl As IntPtr, <MarshalAs(UnmanagedType.LPTStr)> ByVal pszPath As StringBuilder) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function
	
    Public Shared Function worReg(ByVal sKey As String, ByVal sVal As String) As String
        Dim mru As String = sKey

        Dim myString As String = GetPathFromPIDL(CType(Registry.CurrentUser.OpenSubKey(sKey, False).GetValue(sVal), Byte()))
        Return myString
    End Function

    Private Shared Function GetPathFromPIDL(ByVal byteCode() As Byte) As String
        'MAX_PATH = 260
        Dim builder As New StringBuilder(260)

        Dim ptr As IntPtr = IntPtr.Zero
        Dim h0 As GCHandle = GCHandle.Alloc(byteCode, GCHandleType.Pinned)
        Try
            ptr = h0.AddrOfPinnedObject()
        Finally
            h0.Free()
        End Try

        SHGetPathFromIDListW(ptr, builder)

        Return builder.ToString()
    End Function

    Public Shared Function getReg(ByVal sKey As String, ByVal sVal As String) As String
        Dim MyByteArray As Byte()
        Dim MyValue As String
        Dim MyRegistryKey As RegistryKey

        MyRegistryKey = Registry.CurrentUser.OpenSubKey(sKey, False)
        MyByteArray = CType(MyRegistryKey.GetValue(sVal), Byte())

        'Dim ae As New System.Text.ASCIIEncoding
        'MyValue = ae.GetString(MyByteArray, 0, MyByteArray.Length)
        For Each MyByte As Byte In MyByteArray
            MyValue &= Hex(MyByte).ToString.PadLeft(2, "0"c) & " "
        Next
        Return MyValue
    End Function

    Public Shared Function GetVersion() As String
        Dim os_version As OperatingSystem = OSVersion
        With os_version
            Select Case .Platform
                Case .Platform.Win32Windows
                    Select Case (.Version.Minor)
                        Case 0
                            Return "Windows 95"
                        Case 10
                            If .Version.Revision.ToString() = "2222A" Then
                                Return "Windows 98 Second Edition"
                            Else
                                Return "Windows 98"
                            End If
                        Case 90
                            Return "Windows ME"
                    End Select
                Case .Platform.Win32NT
                    Select Case (.Version.Major)
                        Case 3
                            Return "Windows NT 3.51"
                        Case 4
                            Return "Windows NT 4.0"
                        Case 5
                            Select Case (.Version.Minor)
                                Case 0
                                    Return "Windows 2000"
                                Case 1
                                    Return "Windows XP"
                                Case 2
                                    Return "Windows Server 2003"
                            End Select
                        Case 6
                            Return "Windows 7"
                        Case Else
                            Return "Unknown"
                    End Select
                Case Else
                    Return "Unknown"
            End Select
        End With
    End Function
End Class