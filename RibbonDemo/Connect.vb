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
            Case "button1" : strLabel = "启动NTSYS "
            Case "button2" : strLabel = "启动NTEdit"
            Case "button3" : strLabel = "添加品种"
            Case "button4" : strLabel = "重算行列"
        End Select
        Return strLabel
    End Function

    Public Function GetScreenTip(ByVal control As IRibbonControl) As String
        Return "Launch NTSYSpc load data from current sheet"
    End Function

    Public Sub OnAction(ByVal control As IRibbonControl)
        Dim LaunchCommand As String
        Dim resp As Long

        LaunchCommand = "default string"
        Select Case control.Id
            Case "button1" : LaunchCommand = GetNTSYSlaunchingCommand()
            Case "button2" : LaunchCommand = GetNTEditlaunchingCommand()
        End Select

        If LaunchCommand = "NTSYS not found" Then
            MsgBox("NTSYS install not found")
            Return
        End If

        resp = MsgBox("Data in this sheet will be output to c:\output.csv for NTSYS, Change to Excel file will be saved.", MsgBoxStyle.YesNo, "NTSYS计算")

        If resp = vbNo Then
            Return
        End If

        If My.Computer.FileSystem.FileExists("c:\temp\output.csv") Then
            My.Computer.FileSystem.DeleteFile("c:\temp\output.csv")
        End If

        applicationObject.ActiveWorkbook.Save()
        applicationObject.ActiveWorkbook.SaveAs("c:\temp\output.csv", Excel.XlFileFormat.xlCSV)
        applicationObject.ActiveWorkbook.Close(False)


        LaunchCommand = LaunchCommand.Replace("%1", "")
        LaunchCommand = Chr(34) & LaunchCommand & Chr(34)

        '        MsgBox("Launch command is " & LaunchCommand & " C:\\temp\output.csv")

        Shell(LaunchCommand & " C:\\temp\output.csv", AppWinStyle.MaximizedFocus, True, -1)

    End Sub
    Public Sub OnAction_button3(ByVal control As IRibbonControl)
        Dim f As Form1
        MsgBox("button3 clicked") 'Data entry requested
        f = New Form1
        f.Show()
    End Sub
    Public Sub OnAction_button4(ByVal control As IRibbonControl)
        Dim RowCount, ColumnCount As Integer
        Dim RowRange, ColumnRange As Excel.Range
        MsgBox("button4 clicked") 'recalculate rows and columns requested

        RowRange = applicationObject.Range("A3:A10000")
        ColumnRange = applicationObject.Range("B2:HZ2")
        RowCount = applicationObject.WorksheetFunction.CountA(RowRange)
        ColumnCount = applicationObject.WorksheetFunction.CountA(ColumnRange)
        MsgBox(RowCount & ":" & ColumnCount)
        Return

    End Sub

    Public Function GetNTSYSlaunchingCommand() As String
        '        Return "NTSYS"
        Dim NTSYSCommandKey, KeyValue As String
        NTSYSCommandKey = "HKEY_CLASSES_ROOT\\.NTB\shell\\Run in NTSYS\\command"
        KeyValue = My.Computer.Registry.GetValue(NTSYSCommandKey, "", "NTSYS not found").ToString()
        Return KeyValue
    End Function

    Public Function GetNTEditlaunchingCommand() As String
        '        Return "NTEdit"
        Dim NTEditCommandKey, KeyValue As String
        NTEditCommandKey = "HKEY_CLASSES_ROOT\\.NTB\shell\\Open in NTedit\\command"
        KeyValue = My.Computer.Registry.GetValue(NTEditCommandKey, "", "NTSYS not found").ToString()
        Return KeyValue
    End Function


    Public Function GetShowLabel(ByVal control As IRibbonControl) As Boolean
        Dim bolShow As Boolean
        Select Case control.Id
            Case "button1" : bolShow = True
            Case "button2" : bolShow = True
        End Select
        '        Return bolShow
        Return True
    End Function
    Public Function GetShowImage(ByVal control As IRibbonControl) As Boolean
        Return True
    End Function

End Class
