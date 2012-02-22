imports Extensibility
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Core
Imports Excel = Microsoft.Office.Interop.Excel


#Region " Read me for Add-in installation and setup information. "
' When run, the Add-in wizard prepared the registry for the Add-in.
' At a later time, if the Add-in becomes unavailable for reasons such as:
'   1) You moved this project to a computer other than which is was originally created on.
'   2) You chose 'Yes' when presented with a message asking if you wish to remove the Add-in.
'   3) Registry corruption.
' you will need to re-register the Add-in by building the $SAFEOBJNAME$Setup project, 
' right click the project in the Solution Explorer, then choose install.
#End Region

<GuidAttribute("1ABE23C2-0DE0-4FFA-8802-480071463AF9"), ProgIdAttribute("RibbonDemo.Connect")> _
Public Class Connect
	
    Implements Extensibility.IDTExtensibility2, IRibbonExtensibility

    '	Private applicationObject As Object
    Private applicationObject As Excel.Application
    Private addInInstance As Object
	
	Public Sub OnBeginShutdown(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnBeginShutdown
	End Sub
	
	Public Sub OnAddInsUpdate(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnAddInsUpdate
	End Sub
	
	Public Sub OnStartupComplete(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnStartupComplete
	End Sub
	
	Public Sub OnDisconnection(ByVal RemoveMode As Extensibility.ext_DisconnectMode, ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnDisconnection
	End Sub
	
	Public Sub OnConnection(ByVal application As Object, ByVal connectMode As Extensibility.ext_ConnectMode, ByVal addInInst As Object, ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnConnection
        applicationObject = DirectCast(application, Excel.Application)
        addInInstance = addInInst
	End Sub

    Public Function GetCustomUI(ByVal RibbonID As String) As String _
        Implements Microsoft.Office.Core.IRibbonExtensibility.GetCustomUI
        Return My.Resources.Ribbon
    End Function

    Public Function GetLabel(ByVal control As IRibbonControl) As String
        Dim strLabel As String = ""
        Select Case control.Id
            Case "button1" : strLabel = "Launch NTSYSpc Notepad for now"
            Case "button2" : strLabel = "Data Entry Form"
        End Select
        Return strLabel
    End Function

    Public Function GetScreenTip(ByVal control As IRibbonControl) As String
        Return "Launch NTSYSpc load data from current sheet"
    End Function

    Public Sub OnAction(ByVal control As IRibbonControl)

        MsgBox("Data will be saved for NTSYSpc analysis, Excel file will be closed.")

        If My.Computer.FileSystem.FileExists("c:\temp\output.csv") Then
            My.Computer.FileSystem.DeleteFile("c:\temp\output.csv")
        End If
        applicationObject.ActiveWorkbook.SaveAs("c:\temp\output.csv", Excel.XlFileFormat.xlCSV)
        applicationObject.ActiveWorkbook.Close(False)
        '        MsgBox("open output file c:\temp\output.csv use notepad")

        Shell("notepad.exe c:\temp\output.csv", AppWinStyle.MaximizedFocus, True, -1)

        '        Select Case control.Id
        '            Case "button1" : applicationObject.Range("A1").Value = _
        '               "This button inserts text."
        '          Case "button2" : applicationObject.Range("A1").Value = _
        '                "This button inserts more text."
        '        End Select
    End Sub
    Public Function GetShowLabel(ByVal control As IRibbonControl) As Boolean
        Dim bolShow As Boolean
        Select Case control.Id
            Case "button1" : bolShow = True
            Case "button2" : bolShow = False
        End Select
        Return bolShow
    End Function
    Public Function GetShowImage(ByVal control As IRibbonControl) As Boolean
        Return True
    End Function

End Class
