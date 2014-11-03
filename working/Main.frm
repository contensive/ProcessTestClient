VERSION 5.00
Begin VB.Form MainForm 
   Caption         =   "Process Test Client"
   ClientHeight    =   7365
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   6660
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox argTextBox 
      Height          =   1335
      Left            =   2280
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   1920
      Width           =   3495
   End
   Begin VB.TextBox ProcessNameTextBox 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Text            =   "account billing batch process"
      Top             =   1440
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   3360
      Width           =   2055
   End
   Begin VB.TextBox AppNameTextBox 
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Text            =   "jay3-accountbilling"
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "name=value list"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label StatusLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   480
      TabIndex        =   6
      Top             =   4320
      Width           =   5775
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Process Name"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "App Name"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the Program ID of the Process Add-on, and the application name to test it on. Hit start and the Process will be run once."
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   240
      Width           =   6855
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    Dim OptionString As String
    Dim ao As Object
    Dim csv As Object
    Dim CSConn As Object
    'Dim CSConn As CSConnectionType
    Dim ProgramID As String
    Dim Result As String
    Dim CS As Long
    'Dim aoGUID As String
    Dim returnStatusOk As Boolean
    Dim fs As kmaFileSystem3.FileSystemClass
    Dim argList As String
    Dim Ptr As Long
    Dim argPairs() As String
    Dim argName As String
    Dim argValue As String
    Dim Pos As Long
    '
    On Error Resume Next
    Err.Clear
    '
    Set csv = CreateObject("ccCSrvr3.ContentServerClass")
    Call csv.OpenConnection(AppNameTextBox.Text)
    'Set CSConn = csv.OpenConnection(AppNameTextBox.Text)
    If Err.Number <> 0 Then
        StatusLabel.Caption = "There was an error during the Contensive OpenConnection [" & Err.Description & "]"
        Exit Sub
    Else
        If False Then
        'If CSConn.ApplicationStatus <> ApplicationStatusRunning Then
            StatusLabel.Caption = "Application not running"
        Else
            argList = argTextBox.Text
            If argList <> "" Then
                argPairs = Split(argList, vbCrLf)
                For Ptr = 0 To UBound(argPairs)
                    argName = argPairs(Ptr)
                    If argName <> "" Then
                        argValue = ""
                        Pos = InStr(1, argName, "=")
                        If Pos > 0 Then
                            argValue = Mid(argName, Pos + 1)
                            argName = Mid(argName, 1, Pos - 1)
                        End If
                        OptionString = OptionString & "&" & kmaEncodeRequestVariable(argName) & "=" & kmaEncodeRequestVariable(argValue)
                    End If
                Next
            End If
            If OptionString <> "" Then
                OptionString = Mid(OptionString, 2)
            End If
            StatusLabel.Caption = "Starting Addon " & ProcessNameTextBox.Text
            CS = csv.OpenCSContent("Add-ons", "name=" & KmaEncodeSQLText(ProcessNameTextBox.Text))
            If Err.Number <> 0 Then
                StatusLabel.Caption = "There was an error in OpenCSContent, " & Err.Description
            Else
            If Not csv.IsCSOK(CS) Then
                StatusLabel.Caption = "Could not find process [" & ProcessNameTextBox.Text & "] in application [" & AppNameTextBox.Text & "]"
            Else
                Result = csv.ExecuteAddon2(csv.GetCSInteger(CS, "id"), "", OptionString, ContextSimple, "", 0, "", "", False, 0, "", returnStatusOk, Nothing, "", Nothing, "", 0, False)
                'aoGUID = csv.GetCSText(CS, "ccGUID")
                'Call csv.ExecuteAddonAsProcess(aoGUID, , , True)
                If Err.Number <> 0 Then
                    StatusLabel.Caption = "There was an error executing the add-on [" & ProcessNameTextBox.Text & "], " & Err.Description
                Else
                    'Result = ao.execute(csv, Nothing, OptionString, "")
                    If Err.Number <> 0 Then
                        StatusLabel.Caption = "There was an error calling csv.ExecuteAddon2 - " & Err.Description
                    ElseIf Result = "" Then
                        StatusLabel.Caption = "Finished and returned OK (empty result)"
                    Else
                        StatusLabel.Caption = "Finished and returned Error [" & Result & "]"
                    End If
                End If
            End If
            End If
            Call csv.CloseCS(CS)
        End If
    End If
    Set csv = Nothing
    
End Sub

Private Sub Form_Load()
    Dim fs As New kmaFileSystem3.FileSystemClass
    Dim appName As String
    Dim addonname As String
    '
    addonname = fs.ReadFile(App.Path & "addonname.txt")
    ProcessNameTextBox.Text = addonname
    '
    appName = fs.ReadFile(App.Path & "appName.txt")
    AppNameTextBox.Text = appName
    '
    argTextBox.Text = fs.ReadFile(App.Path & "options.txt")
    '
End Sub

Private Sub ProcessNameTextBox_Change()
    Dim fs As New kmaFileSystem3.FileSystemClass
    Call fs.SaveFile(App.Path & "addonname.txt", ProcessNameTextBox.Text)
End Sub

Private Sub AppNameTextBox_Change()
    Dim fs As New kmaFileSystem3.FileSystemClass
    Call fs.SaveFile(App.Path & "appName.txt", AppNameTextBox.Text)
End Sub

Private Sub argTextBox_Change()
    Dim fs As New kmaFileSystem3.FileSystemClass
    Call fs.SaveFile(App.Path & "options.txt", argTextBox.Text)
End Sub

