VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmVBProjectCleaner 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB Project Cleaner"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   8655
   Icon            =   "frmVBProjectCleaner.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   8655
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check 
      Caption         =   "Add &Error handlers"
      Enabled         =   0   'False
      Height          =   300
      Index           =   9
      Left            =   3075
      TabIndex        =   11
      Top             =   3450
      Width           =   2685
   End
   Begin VB.CheckBox Check 
      Caption         =   "Remove all &Comments"
      Enabled         =   0   'False
      Height          =   300
      Index           =   8
      Left            =   3075
      TabIndex        =   10
      Top             =   3030
      Width           =   2685
   End
   Begin VB.CheckBox Check 
      Caption         =   "&Format code"
      Enabled         =   0   'False
      Height          =   300
      Index           =   7
      Left            =   3075
      TabIndex        =   9
      Top             =   2625
      Value           =   1  'Checked
      Width           =   2685
   End
   Begin VB.ListBox ListVarType 
      Height          =   255
      Left            =   9090
      TabIndex        =   35
      Top             =   165
      Width           =   2025
   End
   Begin VB.ListBox ListNSource 
      Height          =   255
      Left            =   9090
      TabIndex        =   34
      Top             =   1995
      Width           =   2025
   End
   Begin VB.ListBox ListN 
      Height          =   255
      Left            =   9090
      TabIndex        =   33
      Top             =   1695
      Width           =   2025
   End
   Begin VB.ListBox ListNames 
      Height          =   255
      Left            =   9090
      TabIndex        =   19
      Top             =   1080
      Width           =   2025
   End
   Begin VB.ListBox ListProcedure 
      Height          =   255
      Left            =   9090
      TabIndex        =   32
      Top             =   780
      Width           =   2025
   End
   Begin VB.CheckBox Check 
      Caption         =   "&Remove unused code"
      Enabled         =   0   'False
      Height          =   300
      Index           =   5
      Left            =   3075
      TabIndex        =   7
      Top             =   1950
      Width           =   2685
   End
   Begin VB.CheckBox Check 
      Caption         =   "Dont remove. &Mark as comment"
      Enabled         =   0   'False
      Height          =   300
      Index           =   6
      Left            =   3075
      TabIndex        =   8
      Top             =   2205
      Value           =   1  'Checked
      Width           =   2685
   End
   Begin VB.CheckBox Check 
      Caption         =   "&Analize entire project"
      Enabled         =   0   'False
      Height          =   300
      Index           =   0
      Left            =   3075
      TabIndex        =   2
      Top             =   435
      Width           =   2685
   End
   Begin VB.CheckBox Check 
      Caption         =   "Analize just &Selected module"
      Enabled         =   0   'False
      Height          =   300
      Index           =   1
      Left            =   3075
      TabIndex        =   3
      Top             =   690
      Value           =   1  'Checked
      Width           =   2685
   End
   Begin VB.CheckBox Check 
      Caption         =   "&Check Public declarations"
      Enabled         =   0   'False
      Height          =   300
      Index           =   4
      Left            =   3075
      TabIndex        =   6
      Top             =   1545
      Value           =   1  'Checked
      Width           =   2685
   End
   Begin VB.CheckBox Check 
      Caption         =   "Check &Private declarations"
      Enabled         =   0   'False
      Height          =   300
      Index           =   3
      Left            =   3075
      TabIndex        =   5
      Top             =   1290
      Value           =   1  'Checked
      Width           =   2685
   End
   Begin VB.CheckBox Check 
      Caption         =   "Check &Dim declarations"
      Enabled         =   0   'False
      Height          =   300
      Index           =   2
      Left            =   3075
      TabIndex        =   4
      Top             =   1065
      Value           =   1  'Checked
      Width           =   2685
   End
   Begin VB.ListBox ListIndex 
      Height          =   255
      Left            =   9090
      TabIndex        =   18
      Top             =   1380
      Width           =   2025
   End
   Begin VB.CommandButton ComOpenVBP 
      Caption         =   "&Open Project"
      Height          =   810
      Left            =   6345
      Picture         =   "frmVBProjectCleaner.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2550
      Width           =   1530
   End
   Begin VB.CommandButton ComDetectUnusedDeclarations 
      Caption         =   "&Do all the checked options"
      Enabled         =   0   'False
      Height          =   390
      Left            =   5880
      TabIndex        =   12
      Top             =   4785
      Width           =   2580
   End
   Begin VB.CommandButton ComExit 
      Caption         =   "&Exit"
      Height          =   390
      Left            =   3120
      TabIndex        =   13
      Top             =   4785
      Width           =   1110
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   210
      Left            =   3105
      TabIndex        =   17
      Top             =   135
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   370
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
   End
   Begin VB.ListBox ListObjects 
      Height          =   255
      Left            =   9090
      TabIndex        =   16
      Top             =   465
      Width           =   2025
   End
   Begin VB.ListBox ListModules 
      Height          =   4740
      Left            =   135
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   405
      Width           =   2850
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   8055
      Top             =   2595
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label LabUnusedPrivate 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   8010
      TabIndex        =   31
      Top             =   1440
      Width           =   450
   End
   Begin VB.Label LabPrivate 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   8010
      TabIndex        =   30
      Top             =   1110
      Width           =   450
   End
   Begin VB.Label LabUnusedPublic 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   8010
      TabIndex        =   29
      Top             =   2070
      Width           =   450
   End
   Begin VB.Label LabPublic 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   8010
      TabIndex        =   28
      Top             =   1755
      Width           =   450
   End
   Begin VB.Label LabUnusedDim 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   8010
      TabIndex        =   27
      Top             =   795
      Width           =   450
   End
   Begin VB.Label LabDim 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   8010
      TabIndex        =   26
      Top             =   480
      Width           =   450
   End
   Begin VB.Label LabDec 
      BackColor       =   &H80000005&
      Caption         =   " Unused Public Declarations"
      Height          =   240
      Index           =   5
      Left            =   5805
      TabIndex        =   25
      Top             =   2070
      Width           =   2310
   End
   Begin VB.Label LabDec 
      BackColor       =   &H80000005&
      Caption         =   " Public Declarations"
      Height          =   240
      Index           =   4
      Left            =   5805
      TabIndex        =   24
      Top             =   1755
      Width           =   2310
   End
   Begin VB.Label LabDec 
      BackColor       =   &H80000005&
      Caption         =   " Unused Private Declarations"
      Height          =   240
      Index           =   3
      Left            =   5805
      TabIndex        =   23
      Top             =   1440
      Width           =   2310
   End
   Begin VB.Label LabDec 
      BackColor       =   &H80000005&
      Caption         =   " Private Declarations"
      Height          =   240
      Index           =   2
      Left            =   5805
      TabIndex        =   22
      Top             =   1110
      Width           =   2310
   End
   Begin VB.Label LabDec 
      BackColor       =   &H80000005&
      Caption         =   " Unused Dim Declarations"
      Height          =   240
      Index           =   1
      Left            =   5805
      TabIndex        =   21
      Top             =   795
      Width           =   2310
   End
   Begin VB.Label LabDec 
      BackColor       =   &H80000005&
      Caption         =   " Dim Declarations"
      Height          =   240
      Index           =   0
      Left            =   5805
      TabIndex        =   20
      Top             =   480
      Width           =   2310
   End
   Begin VB.Label LabModules 
      Caption         =   "Modules"
      Height          =   210
      Left            =   135
      TabIndex        =   15
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label LabInfo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   720
      Left            =   3120
      TabIndex        =   14
      Top             =   3930
      Width           =   5340
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpenProject 
         Caption         =   "&Open VB Project..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFile0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsDetect 
         Caption         =   "&Do all the checked options"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "FrmVBProjectCleaner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'VBProjectCleaner - Franco Gomes - September 2002
' Version 1.0.5 - freeware.
Option Compare Binary
Private TotDim As Long
Private TotPrivate As Long
Private TotPublic As Long
Private UnusedDim As Long
Private UnusedPrivate As Long
Private UnusedPublic As Long
Private ModuleIndex As Integer
Private VBPName As String
Private VBDir As String
Private RetVal As Boolean
Private ProjectFolder As String
Private OKFolder As String
Private UserFile As String
Private NewLine As String
Private ChkEnable As Boolean
Private OldCheckValue As Integer
Private Fs As Object
Private f As Object
Private ts As Object

Private Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long
Private Sub GetUserData()


On Error GoTo GetUserData_Error
' Opens the user file that contains the last option made.
Dim FileNum As Integer
Dim Ca As Integer
Dim Cb As Integer

    If Dir(UserFile) <> vbNullString Then
        FileNum = FreeFile
        Open UserFile For Random As #FileNum
            For Ca = 0 To 9
                Get #FileNum, Ca + 1, Cb
                Check(Ca).Value = Cb
            Next Ca
        Close #FileNum
        Else
        Check(0).Value = 0 ' if no user file found the program use this values.
        Check(2).Value = 1
        Check(3).Value = 1
        Check(4).Value = 1
        Check(5).Value = 0
        Check(7).Value = 1
        Check(8).Value = 0
        Check(9).Value = 1
    End If
    Exit Sub
GetUserData_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(GetUserData) of (frmVBProjectCleaner.frm).", vbCritical

End Sub
Private Sub SaveUserData()

On Error GoTo SaveUserData_Error
' Saves the user option in file
Dim FileNum As Integer
Dim Ca As Integer
Dim Cb As Integer
    FileNum = FreeFile

    Open UserFile For Random As #FileNum
        For Ca = 0 To 9
            Cb = Check(Ca).Value
            Put #FileNum, Ca + 1, Cb
        Next Ca
    Close #FileNum
    Exit Sub
SaveUserData_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(SaveUserData) of (frmVBProjectCleaner.frm).", vbCritical

End Sub
Private Sub Check_Click(Index As Integer)

On Error GoTo Check_Click_Error
Dim Ch As Long

    If ChkEnable = True Then
        If Check(2).Value = 0 And Check(3).Value = 0 And Check(4).Value = 0 Then
            Check(5).Enabled = False
            Check(6).Enabled = False
            Else
            Check(5).Enabled = True
            Check(6).Enabled = True
        End If
    End If
    If Index = 0 Then
        Check(1).Value = 1 - Check(0).Value
        Exit Sub
    End If
    If Index = 1 Then
        Check(0).Value = 1 - Check(1).Value
        Exit Sub
    End If
    If Index = 5 Then
        Check(6).Value = 1 - Check(5).Value
        OldCheckValue = Check(5).Value  ' We keep the user decision for this option.
        Exit Sub
    End If
    If Index = 6 Then
        Check(5).Value = 1 - Check(6).Value
        OldCheckValue = Check(5).Value
        If Check(6).Value = 1 Then Check(8).Value = 0
        Exit Sub
    End If
    If Index = 8 Then
' if this option is checked, the option "Dont delete: Mark as comment" will be unchecked.
        Ch = OldCheckValue
        If Check(8).Value = 1 Then
            Check(5).Value = 1
            Else
            Check(5).Value = OldCheckValue
        End If
        OldCheckValue = Ch ' The previous state of the option "Remove unused code" is reestablished.
    End If
    Exit Sub
Check_Click_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(Check_Click) of (frmVBProjectCleaner.frm).", vbCritical

End Sub
Private Sub ComDetectUnusedDeclarations_Click()

On Error GoTo ComDetectUnusedDeclarations_Click_Error
'---------------------------------
    If Check(1).Value = 1 Then       ' if the user chose "Analyze just selected module",
        If ListModules.ListIndex = -1 Then  ' we verifies if some module is selected.
            MsgBox "Please, select some module from the modules list.", vbInformation, Me.Caption
            Exit Sub
        End If
    End If
'---------------------------------
' inicialization of some counters
    TotDim = 0
    TotPrivate = 0
    TotPublic = 0
    UnusedDim = 0
    UnusedPrivate = 0
    UnusedPublic = 0
'---------------------------------
    EmptyLabels         ' Write empty strings in some labels
    ChecksEnabled False ' Disabling all the options.
'---------------------------------
' Analyze all project or just selected module?
    If Check(0).Value = 1 Then
        For ModuleIndex = 0 To ListModules.ListCount - 1
            ListModules.ListIndex = ModuleIndex
            CheckModule ModuleIndex     ' analyse entire project
        Next ModuleIndex
        ModuleIndex = 0
        Else: CheckModule ModuleIndex   ' analyze just selected module
    End If
'---------------------------------
' Set this stuff to a normal situation
    Me.MousePointer = vbNormal
    ProgressBar.Value = 0
    Me.Caption = "VB Project Cleaner"
'---------------------------------
    ChecksEnabled True ' Enabling the options
    Exit Sub
ComDetectUnusedDeclarations_Click_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(ComDetectUnusedDeclarations_Click) of (frmVBProjectCleaner.frm).", vbCritical

End Sub
Private Sub CheckModule(MIndex As Integer)

On Error GoTo CheckModule_Error
Dim Final As String
    LabInfo.Caption = vbNullString

    ProgressBar.Value = 0
    Me.Caption = "VB Project Cleaner"
'----------------------------------------
' Obtaining the names of all objects in module and doing the first cleaning
    LabInfo.ForeColor = &HC0&
    LabInfo.Caption = "First cleaning of [" & ListModules.List(MIndex) & "]."
    FirstCleaning ListModules.List(MIndex)
    LabInfo.ForeColor = &HC00000
    LabInfo.Caption = "First cleaning of [" & ListModules.List(MIndex) & "]. Done."
'----------------------------------------
    If Check(2).Value = 1 Then ' Searching unused Dim declarations in current module
        LabInfo.ForeColor = &HC0&
        LabInfo.Caption = "Searching Dim declarations in [" & ListModules.List(MIndex) & "]."
        SearchDim ListModules.List(MIndex)
        LabInfo.ForeColor = &HC00000
        LabInfo.Caption = "Searching Dim declarations in [" & ListModules.List(MIndex) & "]. Done."
        LabDim.Caption = TotDim
        LabUnusedDim.Caption = UnusedDim
    End If
    If Check(3).Value = 1 Then ' Searching unused Private declarations in current module
        LabInfo.ForeColor = &HC0&
        LabInfo.Caption = "Searching Private declarations in [" & ListModules.List(MIndex) & "]."
        SearchPrivate ListModules.List(MIndex)
        LabInfo.ForeColor = &HC00000
        LabInfo.Caption = "Searching Private declarations in [" & ListModules.List(MIndex) & "]. Done."
        LabPrivate.Caption = TotPrivate
        LabUnusedPrivate.Caption = UnusedPrivate
    End If
    If Check(4).Value = 1 Then ' Searching unused Public declarations in current module
        LabInfo.ForeColor = &HC0&
        LabInfo.Caption = "Searching Public declarations in [" & ListModules.List(MIndex) & "]."
        SearchPublic ListModules.List(MIndex)
        LabInfo.ForeColor = &HC00000
        LabInfo.Caption = "Searching Public declarations in [" & ListModules.List(MIndex) & "]. Done."
        LabPublic.Caption = TotPublic
        LabUnusedPublic.Caption = UnusedPublic
    End If
    If Check(7).Value = 1 Then
        LabInfo.ForeColor = &HC0&
        LabInfo.Caption = "Formating Code in [" & ListModules.List(MIndex) & "]."
        Final = FormatCode(ListModules.List(MIndex))
        WriteCode ListModules.List(MIndex), Final
        LabInfo.ForeColor = &HC00000
        LabInfo.Caption = "Formating Code in [" & ListModules.List(MIndex) & "]. Done."
    End If
    If Check(9).Value = 1 Then
        LabInfo.ForeColor = &HC0&
        LabInfo.Caption = "Writing Error Handlers in [" & ListModules.List(MIndex) & "]."
        Final = WriteErrorHandlers(ListModules.List(MIndex))
        WriteCode ListModules.List(MIndex), Final
        LabInfo.ForeColor = &HC00000
        LabInfo.Caption = "Writing Error Handlers in [" & ListModules.List(MIndex) & "]. Done."
    End If
    LabInfo.ForeColor = &HC00000
    LabInfo.Caption = "All done."
    Exit Sub
CheckModule_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(CheckModule) of (frmVBProjectCleaner.frm).", vbCritical

End Sub
Private Sub ComExit_Click()

On Error GoTo ComExit_Click_Error
    Unload Me
    Exit Sub
ComExit_Click_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(ComExit_Click) of (frmVBProjectCleaner.frm).", vbCritical

End Sub
Private Sub ComOpenVBP_Click()

On Error GoTo ComOpenVBP_Click_Error

Dim FName As String
Dim VBPFileName As String
Dim SName As String
Dim DName As String
Dim TName As String
Dim RName As String
Dim Final As String
Dim DPos As Long
Dim Ca As Integer
Dim Cb As Integer
Dim RetVal As Boolean

    LabInfo.Caption = vbNullString
    ProgressBar.Value = 0
    Me.Caption = "VB Project Cleaner"
    ChecksEnabled False
    Me.MousePointer = vbHourglass
    RetVal = OpenVBP
    If RetVal = False Then
        Me.MousePointer = vbNormal
        Exit Sub
    End If
    VBDir = GetPath(VBPName)
    OKFolder = "Clean_" & ProjectFolder ' Building the name for the "working folder"
    VBPFileName = Right(VBPName, Len(VBPName) - Len(VBDir))
    LabInfo.ForeColor = &HC0&
    LabInfo.Caption = "Opening the VB Project [" & VBPFileName & "]."
    FName = Dir(VBDir & OKFolder, vbDirectory)
    If FName = vbNullString Then MkDir VBDir & OKFolder ' Creat Directory for working
    LabModules.Caption = "Modules in [" & VBPFileName & "] project."
'copy vbp and vbw files to OKFolder
    CopyFile VBPName, VBDir & OKFolder & "\" & VBPFileName
    CopyFile VBDir & Left(VBPFileName, Len(VBPFileName) - 3) & "vbw", VBDir & OKFolder & "\" & Left(VBPFileName, Len(VBPFileName) - 3) & "vbw"
    GetModulesNames VBDir & OKFolder & "\" & VBPFileName ' Obtains all the names of the forms and modules.
    Set Fs = CreateObject("Scripting.FileSystemObject") ' Open the VBP file
    Set f = Fs.GetFile(VBDir & OKFolder & "\" & VBPFileName)
    Set ts = f.OpenAsTextStream(1, -2)
    Final = ts.readall
    For Ca = 0 To ListModules.ListCount - 1 ' all forms and modules are copied to OKFolder
        TName = ListModules.List(Ca)
        SName = TName
        DPos = InStr(TName, "\")
        If DPos > 0 Then
            For Cb = Len(TName) To 1 Step -1
                If Mid(TName, Cb, 1) = "\" Then
                    RName = Right(TName, Len(TName) - Cb)
                    ListModules.List(Ca) = RName
                    Final = Replace(Final, TName, RName) ' Corrects the path in VBP file
                    TName = RName
                    Exit For
                End If
            Next Cb
        End If
        DName = VBDir & OKFolder & "\" & TName
        CopyFile SName, DName
        If UCase(Right(SName, 3)) = "FRM" Then
            SName = Left(SName, Len(SName) - 3) & "frx"
            DName = Left(DName, Len(DName) - 3) & "frx"
            CopyFile SName, DName
        End If
        If UCase(Right(SName, 3)) = "CTL" Then
            SName = Left(SName, Len(SName) - 3) & "ctx"
            DName = Left(DName, Len(DName) - 3) & "ctx"
            CopyFile SName, DName
        End If
    Next Ca
    ts.Close
    WriteCode VBPFileName, Final ' Writes the corrected VBP file
    LabInfo.ForeColor = &HC00000
    LabInfo.Caption = vbNullString
    Me.MousePointer = vbNormal
    ChecksEnabled True
    Exit Sub
ComOpenVBP_Click_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(ComOpenVBP_Click) of (frmVBProjectCleaner.frm).", vbCritical

End Sub
Private Sub FirstCleaning(FName As String)

On Error GoTo FirstCleaning_Error
' Obtains all the names of the object in module "FName"
' We need this information later.
Dim Txt As String
Dim Lxt As String
Dim Rxt As String
Dim LenLine As Long
Dim LenFic As Long
Dim Ca As Long
Dim Cb As Long
Dim DPos As Long
Dim DeclZone As Boolean
Dim ObjectZone As Boolean
Dim Final As String

    LenFic = FileLen(VBDir & OKFolder & "\" & FName)
    ProgressBar.Value = 0
    Me.Caption = "VB Project Cleaner"
    ListObjects.Clear
    If UCase(Right(FName, 3)) = "CLS" Then ListObjects.AddItem "Class"
    If UCase(Right(FName, 3)) = "FRM" Then ListObjects.AddItem "Form"
    Set Fs = CreateObject("Scripting.FileSystemObject")
    Set f = Fs.GetFile(VBDir & OKFolder & "\" & FName)
    Set ts = f.OpenAsTextStream(1, -2)
    LenLine = 0
    ObjectZone = True
    DeclZone = True
    Do While ts.AtEndOfStream <> True
        Txt = ts.Readline
        LenLine = LenLine + (Len(Txt) + 2)
        DPos = InStr(Txt, "Implements ")
        If DPos > 0 Then
            Lxt = GetWord(Txt, DPos + 11)
            ListObjects.AddItem Lxt
        End If
        If UCase(Right(FName, 3)) = "CTL" Then
            ListObjects.AddItem "UserControl"
            DPos = InStr(Txt, "Begin VB.UserControl ")
            If DPos > 0 Then
                Lxt = (Mid(Txt, 22, Len(Txt) - 22))
                ListObjects.AddItem Lxt
            End If
        End If
        If Left(Txt, 12) = "Attribute VB" Or Left(Txt, 15) = "Option Explicit" Then ObjectZone = False
        If Check(8).Value = 1 Then ' Remove comments
            If Left(Txt, 1) = "'" Then
                Txt = vbNullString
                Else
                For Ca = 1 To Len(Txt)
                    If Mid(Txt, Ca, 2) = " '" Then
                        If Mid(Txt, Ca, 3) <> " '""" And Mid(Txt, Ca, 4) <> " ' """ Then
                            Txt = Left(Txt, Ca - 1)
                            Exit For
                        End If
                        Exit For
                    End If
                Next Ca
            End If
        End If
        If ObjectZone = False Then
            If Txt <> vbNullString And Txt <> Space(Len(Txt)) Then
                If Right(Txt, 2) = " _" Then ' reconstitues "fragmented" lines
                    Txt = Left(Txt, Len(Txt) - 1)
                    Do While ts.AtEndOfStream <> True
                        Lxt = ts.Readline
                        LenLine = LenLine + (Len(Lxt) + 2)
                        For Cb = 1 To Len(Lxt) 'removes spaces on left side of the line
                            If Mid(Lxt, Cb, 1) <> " " Then
                                Lxt = Right(Lxt, (Len(Lxt) - Cb) + 1)
                                Exit For
                            End If
                        Next Cb
                        Txt = Txt & Lxt
                        If Right(Txt, 2) = " _" Then
                            Txt = Left(Txt, Len(Txt) - 1)
                            Else: Exit Do
                        End If
                    Loop
                End If
                If DeclZone = True Then ' We are in declarations zone of the module.
'The declarations zone is ended if one of the followed occurrences is founded.
                    If Left(Txt, 4) = "Sub " Or Left(Txt, 12) = "Private Sub " Or Left(Txt, 11) = "Public Sub " Or Left(Txt, 9) = "Function " Or Left(Txt, 17) = "Private Function " Or Left(Txt, 16) = "Public Function " Or Left(Txt, 9) = "Property " Or Left(Txt, 17) = "Private Property " Or Left(Txt, 16) = "Public Property " Then DeclZone = False 'the end of declarations zone is founded.
                End If
                Txt = VerifyLine(Txt, DeclZone) ' We call a Function to verify the line,
            End If
            Else: CheckObject Txt
        End If
        Final = Final & Txt & NewLine
        DoEvents
        SetProgressBar LenFic, LenLine
    Loop
    ts.Close
    WriteCode FName, Final
    Exit Sub
FirstCleaning_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(FirstCleaning) of (frmVBProjectCleaner.frm).", vbCritical

End Sub
Private Function GetWord(Sx As String, SPos As Long) As String

Dim Ga As Long
Dim Gx As String

    Gx = Sx & " "
    For Ga = SPos To Len(Gx)
        If Mid(Gx, Ga, 1) = " " Then
            GetWord = Mid(Gx, SPos, Ga - SPos)
            Exit Function
        End If
    Next Ga
        
End Function
Private Sub WriteCode(FName As String, FTxt As String)

On Error GoTo WriteCode_Error

    Set Fs = CreateObject("Scripting.FileSystemObject")
    Set f = Fs.CreateTextFile(VBDir & OKFolder & "\" & FName)
    f.Write (FTxt)
    f.Close ' Writing the file with the transformed text and close it.
    ProgressBar.Value = 0
    Me.Caption = "VB Project Cleaner"
    Exit Sub
WriteCode_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(WriteCode) of (frmVBProjectCleaner.frm).", vbCritical

End Sub
Private Function FormatCode(FName As String) As String

On Error GoTo FormatCode_Error
Dim Ca As Long
Dim Txt As String
Dim LenLine As Long
Dim LenFic As Long
Dim DPos As Long
Dim ObjectZone As Boolean
Dim LineBefore As Boolean
Dim LineAfter As Boolean
Dim HasLineBefore As Boolean
Dim ThisLine As Integer
Dim NextLine As Integer
Dim Found As Boolean
Dim Final As String
    LenFic = FileLen(VBDir & OKFolder & "\" & FName)

    ProgressBar.Value = 0
    Me.Caption = "VB Project Cleaner"
    Set Fs = CreateObject("Scripting.FileSystemObject")
    Set f = Fs.GetFile(VBDir & OKFolder & "\" & FName)
    Set ts = f.OpenAsTextStream(1, -2)
    LenLine = 0
    ObjectZone = True
    Do While ts.AtEndOfStream <> True
        Txt = ts.Readline
        If Txt <> vbNullString And Txt <> Space(Len(Txt)) Then ' removes empty lines
            If Left(Txt, 12) = "Attribute VB" Then ObjectZone = False
            If ObjectZone = False Then
                If Txt <> vbNullString And Txt <> Space(Len(Txt)) Then
                    For Ca = 1 To Len(Txt) 'removes spaces on left side of the line
                        If Mid(Txt, Ca, 1) <> " " Then
                            Txt = Right(Txt, (Len(Txt) - Ca) + 1)
                            Exit For
                        End If
                    Next Ca
                    Found = False
                    If Left(Txt, 4) = "Sub " Or Left(Txt, 12) = "Private Sub " Or Left(Txt, 11) = "Public Sub " Or Left(Txt, 9) = "Property " Or Left(Txt, 17) = "Private Property " Or Left(Txt, 16) = "Public Property " Or Left(Txt, 9) = "Function " Or Left(Txt, 17) = "Private Function " Or Left(Txt, 16) = "Public Function " Or Left(Txt, 8) = "End Type" Then
                        Found = True
                        ThisLine = 0
                        NextLine = 4
                        DPos = InStr(Txt, " _")
                        If DPos = 0 Then LineAfter = True
                        LineBefore = False
                        HasLineBefore = True
                    End If
                    If Found = False Then
                        If Left(Txt, 1) = "'" Then
                            ThisLine = 0
                            LineAfter = False
                            LineBefore = False
                            HasLineBefore = True
                        End If
                    End If
                    If Found = False Then
                        If Left(Txt, 7) = "End Sub" Or Left(Txt, 12) = "End Function" Or Left(Txt, 12) = "End Property" Then
                            Found = True
                            ThisLine = 0
                            NextLine = 0
                            LineAfter = False
                            LineBefore = True
                        End If
                    End If
                    If Found = False Then
                        If Left(Txt, 16) = "Private Declare " Or Left(Txt, 15) = "Public Declare " Or Left(Txt, 5) = "Type " Or Left(Txt, 13) = "Private Type " Or Left(Txt, 12) = "Public Type " Then
                            LineAfter = False
                            If HasLineBefore = False Then
                                LineBefore = True
                                HasLineBefore = True
                            End If
                            Found = True
                            LineAfter = False
                            ThisLine = 0
                            NextLine = 4
                        End If
                    End If
                    If Found = False Then
                        If Left(Txt, 4) = "Dim " Or Left(Txt, 7) = "Static " Or Left(Txt, 8) = "Private " Or Left(Txt, 7) = "Public " Or Left(Txt, 7) = "Option " Or Left(Txt, 13) = "Attribute VB_" Then
                            Found = True
                            ThisLine = 0
                            NextLine = 4
                            HasLineBefore = False
                            LineBefore = False
                            LineAfter = False
                        End If
                    End If
                    If Found = False Then
                        If Left(Txt, 11) = "Select Case" Or Left(Txt, 5) = "Open " Or Left(Txt, 3) = "Do " Or Left(Txt, 4) = "For " Or Left(Txt, 6) = "While " Or Left(Txt, 5) = "With " Then
                            NextLine = ThisLine + 4
                            LineAfter = False
                            If HasLineBefore = False Then
                                LineBefore = True
                                HasLineBefore = True
                            End If
                            Found = True
                        End If
                    End If
                    If Found = False Then
                        If Left(Txt, 3) = "If " Then
                            If Right(Txt, 5) = " Then" Then
                                Found = True
                                Else
                                DPos = InStr(Txt, " Then ") ' search for the end of the condition "If"
                                If DPos > 0 Then
                                    For Ca = DPos + 6 To Len(Txt)
                                        If Mid(Txt, Ca, 1) <> "'" And Mid(Txt, Ca, 1) <> " " Then Exit For
                                        If Mid(Txt, Ca, 1) = "'" Then
                                            Found = True
                                            Exit For
                                        End If
                                    Next Ca
                                End If
                            End If
                            If Found = True Then
                                NextLine = ThisLine + 4
                                LineAfter = False
                                If HasLineBefore = False Then
                                    LineBefore = True
                                    HasLineBefore = True
                                End If
                            End If
                        End If
                    End If
                    If Found = False Then
                        If Right(Txt, 5) = " Then" And Left(Txt, 7) <> "ElseIf " Then  ' Search for the keyword "Then"
                            Found = True
                            Else
                            DPos = InStr(Txt, " Then ")
                            If DPos > 0 Then
                                For Ca = DPos + 6 To Len(Txt)
                                    If Mid(Txt, Ca, 1) <> "'" And Mid(Txt, Ca, 1) <> " " Then Exit For
                                    If Mid(Txt, Ca, 1) = "'" Then
                                        Found = True
                                        Exit For
                                    End If
                                Next Ca
                            End If
                        End If
                        If Found = True Then
                            NextLine = ThisLine + 4
                            LineAfter = False
                            If HasLineBefore = False Then
                                LineBefore = True
                                HasLineBefore = True
                            End If
                        End If
                    End If
                    If Found = False Then
                        If Left(Txt, 7) = "End If " Or Txt = "End If" Or Left(Txt, 11) = "End Select " Or Txt = "End Select" Or Left(Txt, 5) = "Loop " Or Txt = "Loop" Or Left(Txt, 5) = "Next " Or Txt = "Next" Or Left(Txt, 6) = "Close " Or Txt = "Close" Or Left(Txt, 9) = "End With " Or Txt = "End With" Or Left(Txt, 6) = "Whend " Or Txt = "Whend" Then
                            ThisLine = ThisLine - 4
                            NextLine = ThisLine
                            LineAfter = False
                            LineBefore = False
                            Found = True
                        End If
                    End If
                    If Found = False Then
                        If HasLineBefore = False Then
                            DPos = InStr(Txt, " _")
                            If DPos = 0 Then
                                LineAfter = True
                                HasLineBefore = True
                                Else: LineBefore = False
                            End If
                        End If
                    End If
                    If Txt <> vbNullString Then
                        If ThisLine > 0 Then
                            Txt = Space(ThisLine) & Txt
                        End If
                    End If
                    If LineBefore = True Then Txt = NewLine & Txt
                    If LineAfter = True Then Txt = Txt & NewLine
                    LineBefore = False
                    LineAfter = False
                    ThisLine = NextLine
                End If
            End If
            Final = Final & Txt & NewLine
        End If
        LenLine = LenLine + (Len(Txt) + 2)
        DoEvents
        SetProgressBar LenFic, LenLine
    Loop
    ts.Close
    FormatCode = Final
    Exit Function
FormatCode_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(FormatCode) of (frmVBProjectCleaner.frm).", vbCritical

End Function
Private Function WriteErrorHandlers(FName As String) As String

On Error GoTo WriteErrorHandlers_Error
' Write error handlers in procedures
Dim Txt As String
Dim LenLine As Long
Dim LenFic As Long
Dim Final As String
Dim InProc As String
Dim EndProc As Boolean
    ListProcedure.Clear

    LenFic = FileLen(VBDir & OKFolder & "\" & FName)
    ProgressBar.Value = 0
    Me.Caption = "VB Project Cleaner"
    Set Fs = CreateObject("Scripting.FileSystemObject") ' Opening the module "FName"
    Set f = Fs.GetFile(VBDir & OKFolder & "\" & FName)
    Set ts = f.OpenAsTextStream(1, -2)
    Do While ts.AtEndOfStream <> True
        Txt = ts.Readline
        LenLine = LenLine + Len(Txt) + 2
        InProc = ProcedureBegin(Txt) ' search for the begin of the procedure
        If InProc <> vbNullString Then
            ListProcedure.Clear
            ListProcedure.AddItem Txt ' "ListProcedures" keeps all the lines of the procedure
            EndProc = False
            Do Until EndProc = True Or ts.AtEndOfStream = True
                Txt = ts.Readline
                LenLine = LenLine + Len(Txt) + 2
                ListProcedure.AddItem Txt
                EndProc = ProcedureEnd(Txt)
            Loop
            Txt = WriteProcedureErrorHandle(InProc, FName) ' Write error handlers in the procedure
        End If
        Final = Final & Txt & NewLine
        DoEvents
        SetProgressBar LenFic, LenLine
    Loop
    ts.Close
    WriteErrorHandlers = Final
    Exit Function
WriteErrorHandlers_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(WriteErrorHandlers) of (frmVBProjectCleaner.frm).", vbCritical

End Function
Private Sub CheckObject(Ctx As String)

On Error GoTo CheckObject_Error
' Search for the names of objects in the form, analysing the line "Ctx"
Dim Cpos As Long
Dim Ka As Long
Dim Kb As Long
Dim ObjName As String
    Cpos = InStr(Ctx, "Begin VB.MDIForm ")

    If Cpos <> 0 Then
        ListObjects.AddItem "MDIForm"
        Exit Sub
    End If
    Cpos = InStr(Ctx, "Begin ") ' The name of the object maybe is found after the keyword "Begin"
    If Cpos = 0 Then Cpos = InStr(Ctx, "Public WithEvents ")    ' In this case, is not an object but
    If Cpos = 0 Then Cpos = InStr(Ctx, "Private WithEvents ")   ' acts like one.
    If Cpos <> 0 Then
        If Right(Ctx, 1) = " " Then Ctx = Left(Ctx, Len(Ctx) - 1)
        For Ka = Len(Ctx) To 1 Step -1
            If Mid(Ctx, Ka, 1) = " " Then
                ObjName = Right(Ctx, Len(Ctx) - Ka)
                If ListObjects.ListCount = 0 Then
                    ListObjects.AddItem ObjName
                    Exit Sub
                End If
                For Kb = 0 To ListObjects.ListCount - 1 ' verify if the name is already in the list
                    If ObjName = ListObjects.List(Kb) Then Exit Sub
                Next Kb
                ListObjects.AddItem ObjName ' Write the object name in the list.
                Exit Sub
            End If
        Next Ka
    End If
    Exit Sub
CheckObject_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(CheckObject) of (frmVBProjectCleaner.frm).", vbCritical

End Sub
Private Sub SearchDim(FName)

On Error GoTo SearchDim_Error
' Search for Dim declarations
Dim Txt As String
Dim LenLine As Long
Dim LenFic As Long
Dim Final As String
Dim InProc As String
Dim EndProc As Boolean
    ListProcedure.Clear

    LenFic = FileLen(VBDir & OKFolder & "\" & FName)
    ProgressBar.Value = 0
    Me.Caption = "VB Project Cleaner"
    Set Fs = CreateObject("Scripting.FileSystemObject") ' Opening the module "FName"
    Set f = Fs.GetFile(VBDir & OKFolder & "\" & FName)
    Set ts = f.OpenAsTextStream(1, -2)
    Do While ts.AtEndOfStream <> True
        Txt = ts.Readline
        LenLine = LenLine + Len(Txt) + 2
        InProc = ProcedureBegin(Txt) ' search for the begin of the procedure
        If InProc <> vbNullString Then
            ListProcedure.Clear
            ListProcedure.AddItem Txt ' "ListProcedures" keeps all the lines of the procedure
            EndProc = False
            Do Until EndProc = True Or ts.AtEndOfStream = True
                Txt = ts.Readline
                ListProcedure.AddItem Txt
                EndProc = ProcedureEnd(Txt)
            Loop
            Txt = ScanProcedure ' search for Dim declarations und unused variables
        End If
        Final = Final & Txt & NewLine
        DoEvents
        SetProgressBar LenFic, LenLine
    Loop
    ts.Close
    Set Fs = CreateObject("Scripting.FileSystemObject") ' Writing the alterations
    Set f = Fs.CreateTextFile(VBDir & OKFolder & "\" & FName)
    f.Write (Final)
    f.Close
    Exit Sub
SearchDim_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(SearchDim) of (frmVBProjectCleaner.frm).", vbCritical

End Sub
Private Sub SearchPrivate(FName)

On Error GoTo SearchPrivate_Error
' Search for private declarations
Dim Txt As String
Dim LenLine As Long
Dim LenFic As Long
Dim Final As String
    ListProcedure.Clear

    LenFic = FileLen(VBDir & OKFolder & "\" & FName)
    ProgressBar.Value = 0
    Me.Caption = "VB Project Cleaner"
    Set Fs = CreateObject("Scripting.FileSystemObject")
    Set f = Fs.GetFile(VBDir & OKFolder & "\" & FName)
    Set ts = f.OpenAsTextStream(1, -2)
    Do While ts.AtEndOfStream <> True
        Txt = ts.Readline
        LenLine = LenLine + Len(Txt) + 2
        ListProcedure.AddItem Txt ' "ListProcedure" keeps all module after the Object's declarations zone
        DoEvents
        SetProgressBar LenFic, LenLine
    Loop
    Txt = ScanPrivate ' We call a routine that will make all the hard work.
    Final = Final & Txt
    ts.Close
    Set Fs = CreateObject("Scripting.FileSystemObject")
    Set f = Fs.CreateTextFile(VBDir & OKFolder & "\" & FName)
    f.Write (Final)
    f.Close
    Exit Sub
SearchPrivate_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(SearchPrivate) of (frmVBProjectCleaner.frm).", vbCritical

End Sub
Private Sub SearchPublic(FName)

On Error GoTo SearchPublic_Error
' Search of Public declarations
Dim Txt As String
Dim LenLine As Long
Dim LenFic As Long
Dim Final As String
Dim VName As String
Dim LineTx As String
Dim IsVar As Boolean
Dim InProcedure As Boolean
Dim Ca As Long
Dim Cb As Long
    ListProcedure.Clear

    LenFic = FileLen(VBDir & OKFolder & "\" & FName)
    ProgressBar.Value = 0
    Me.Caption = "VB Project Cleaner"
    Set Fs = CreateObject("Scripting.FileSystemObject")
    Set f = Fs.GetFile(VBDir & OKFolder & "\" & FName)
    Set ts = f.OpenAsTextStream(1, -2)
    Do While ts.AtEndOfStream <> True
        Txt = ts.Readline
        LenLine = LenLine + Len(Txt) + 2
        ListProcedure.AddItem Txt   ' "ListProcedure" keeps all lines of the module
        DoEvents                    '  after the Object's declarations zone
        SetProgressBar LenFic, LenLine
    Loop
    ts.Close
    ScanPublic  ' we call a routine that list all the public declarations
' and obtains the name and usage of the Public declarations in the current module
    If ListNames.ListCount > 0 Then
'   Scan the other modules of the project for the Public declarations listed in "ListNames"
        For Ca = 0 To ListModules.ListCount - 1
            If ListModules.List(Ca) <> FName Then ' do not search in the "Source" module
                LenFic = FileLen(VBDir & OKFolder & "\" & ListModules.List(Ca))
                ProgressBar.Value = 0
                Me.Caption = "VB Project Cleaner"
                LenLine = 0
                LabInfo.ForeColor = &HC0&
                LabInfo.Caption = "Scanning [" & ListModules.List(Ca) & "]" & NewLine & "for the usage of Public declarations made in [" & FName & "]."
                Set Fs = CreateObject("Scripting.FileSystemObject")
                Set f = Fs.GetFile(VBDir & OKFolder & "\" & ListModules.List(Ca))
                Set ts = f.OpenAsTextStream(1, -2)
                Do While ts.AtEndOfStream <> True
                    Txt = ts.Readline
                    LenLine = LenLine + Len(Txt) + 2
                    For Cb = 0 To ListNames.ListCount - 1
                        VName = ListNames.List(Cb)
                        IsVar = IsVarTrue(Txt, VName)
'"ListN" Keeps the amount of the Public declarations usage
                        If IsVar = True Then ListN.List(Cb) = CInt(ListN.List(Cb)) + 1
                    Next Cb
                    DoEvents
                    SetProgressBar LenFic, LenLine
                Loop
                ts.Close
            End If
        Next Ca
        ProgressBar.Value = 0
        Me.Caption = "VB Project Cleaner"
        For Ca = 0 To ListN.ListCount - 1
' "ListN" has the amount of declarations used outside the source module;
' "ListNSource" has the amount of declarations used inside the source module.
' If only used in the "Source" module, Public declarations are converted to Private
            If CInt(ListN.List(Ca)) < 1 And CInt(ListNSource.List(Ca)) > 1 Then
                LineTx = ListProcedure.List(ListIndex.List(Ca))
                LineTx = Replace(LineTx, "Public ", "Private ")
                ListProcedure.List(ListIndex.List(Ca)) = LineTx
                TotPrivate = TotPrivate + 1
                LabPrivate.Caption = TotPrivate
                TotPublic = TotPublic - 1
                LabPublic.Caption = TotPublic
            End If
' Unused in all modules
            If CInt(ListN.List(Ca)) < 1 And CInt(ListNSource.List(Ca)) < 2 Then
                UnusedPublic = UnusedPublic + 1
                LabUnusedPublic.Caption = UnusedPublic
                LineTx = ListProcedure.List(ListIndex.List(Ca))
                If Left(LineTx, 11) = "Public Sub " Or Left(LineTx, 12) = "Public Type " Or Left(LineTx, 16) = "Public Function " Or Left(LineTx, 11) = "Public Property " Then
                    InProcedure = True
                    Else: InProcedure = False
                End If
' The unused declarations are marked as comments
                ListProcedure.List(ListIndex.List(Ca)) = "'~~Unused~~ |" & LineTx
                If InProcedure = True Then
                    For Cb = ListIndex.List(Ca) + 1 To ListProcedure.ListCount - 1
                        LineTx = ListProcedure.List(Cb)
                        ListProcedure.List(Cb) = "'~~~~~~~~~~ |" & LineTx
                        If Left(LineTx, 7) = "End Sub" Or Left(LineTx, 8) = "End Type" Or Left(LineTx, 12) = "End Function" Or Left(LineTx, 12) = "End Property" Then
                            InProcedure = False
                            Exit For
                        End If
                    Next Cb
                End If
            End If
            DoEvents
            SetProgressBar ListN.ListCount, Ca
        Next Ca
    End If
    For Ca = 0 To ListProcedure.ListCount - 1
        Txt = ListProcedure.List(Ca) & NewLine
' If the user have choose "Remove unused code" we remove our comment lines.
        If Check(5).Value = 1 And Left(Txt, 3) = "'~~" Then Txt = vbNullString
        Final = Final & Txt
    Next Ca
    Set Fs = CreateObject("Scripting.FileSystemObject")
    Set f = Fs.CreateTextFile(VBDir & OKFolder & "\" & FName)
    f.Write (Final) ' writing the corrected staff.
    f.Close
    Exit Sub
SearchPublic_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(SearchPublic) of (frmVBProjectCleaner.frm).", vbCritical

End Sub
Private Function ProcedureBegin(Bxx As String) As String

On Error GoTo ProcedureBegin_Error
' Searchs for the beginning of a procedure
Dim Pxx As String
Dim PPos As Long
Dim QPos As Long
Dim Cp As Long
Dim Cq As Long
    Pxx = " " & Bxx & " "

    ProcedureBegin = vbNullString
    PPos = InStr(Pxx, " Private Sub ")
    If PPos = 0 Then PPos = InStr(Pxx, " Public Sub ")
    If PPos = 0 Then PPos = InStr(Pxx, " Private Function ")
    If PPos = 0 Then PPos = InStr(Pxx, " Public Function ")
    If PPos = 0 Then PPos = InStr(Pxx, " Private Property ")
    If PPos = 0 Then PPos = InStr(Pxx, " Public Property ")
    If PPos = 0 Then PPos = InStr(Pxx, " Sub ")
    If PPos = 0 Then Exit Function
    QPos = InStr(Pxx, "'")
    If QPos > 0 And QPos < PPos Then Exit Function ' The line is a comment, not code.
    QPos = InStr(Pxx, """")
    If QPos > 0 Then ' The ocurrency is text? Not code?
        For Cp = PPos To Len(Pxx)
            If Mid(Pxx, Cp, 1) = "&" Or Mid(Pxx, Cp, 1) = "+" Then Exit For
            If Mid(Pxx, Cp, 1) = """" Then Exit Function
        Next Cp
    End If
    For Cp = 1 To Len(Pxx)
        If Mid(Pxx, Cp, 1) = "(" Then
            For Cq = Cp - 1 To 1 Step -1
                If Mid(Pxx, Cq, 1) = " " Then
                    ProcedureBegin = Mid(Pxx, Cq + 1, Cp - (Cq + 1))
                    Exit Function
                End If
            Next Cq
        End If
    Next Cp
    Exit Function
ProcedureBegin_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(ProcedureBegin) of (frmVBProjectCleaner.frm).", vbCritical

End Function
Private Function ProcedureEnd(Bxx As String) As Boolean

On Error GoTo ProcedureEnd_Error
' Searchs for the end of a procedure
Dim Pxx As String
Dim PPos As Long
Dim QPos As Long
Dim Cp As Long
    Pxx = " " & Bxx & " "

    ProcedureEnd = False
    PPos = InStr(Pxx, " End Sub ")
    If PPos = 0 Then PPos = InStr(Pxx, " End Function ")
    If PPos = 0 Then PPos = InStr(Pxx, " End Property ")
    If PPos = 0 Then Exit Function
    QPos = InStr(Pxx, "'")
    If QPos > 0 And QPos < PPos Then Exit Function ' The line is a comment, not code.
    QPos = InStr(Pxx, """")
    If QPos > 0 Then
        For Cp = PPos To Len(Pxx) ' The line looks text, not code.
            If Mid(Pxx, Cp, 1) = "&" Or Mid(Pxx, Cp, 1) = "+" Then Exit For
            If Mid(Pxx, Cp, 1) = """" Then Exit Function
        Next Cp
    End If
    ProcedureEnd = True
    Exit Function
ProcedureEnd_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(ProcedureEnd) of (frmVBProjectCleaner.frm).", vbCritical

End Function
Private Function ScanProcedure() As String

On Error GoTo ScanProcedure_Error
' Verifying the entire procedure
Dim Lc As Long
Dim Sc As Long
Dim VName As String
Dim Ltx As String
Dim Xtx As String
Dim IsVar As Boolean
Dim VarCt As Long
    ScanProcedure = vbNullString

    For Lc = 0 To ListProcedure.ListCount - 1
        Ltx = ListProcedure.List(Lc)
        VName = GetDecls(Ltx, "Dim")
        If VName <> vbNullString Then ' found Dim declaration
            TotDim = TotDim + 1
            LabDim.Caption = TotDim
            VarCt = 0
            For Sc = 0 To ListProcedure.ListCount - 1
                Xtx = ListProcedure.List(Sc)
                IsVar = IsVarTrue(Xtx, VName)           ' we call a routine to confirm if the
                If IsVar = True Then VarCt = VarCt + 1  ' variable/const. is in this line.
                If VarCt > 1 Then Exit For
            Next Sc
            If VarCt < 2 Then ' variable/const. unused
                UnusedDim = UnusedDim + 1
                LabUnusedDim.Caption = UnusedDim
                ListProcedure.List(Lc) = "'~Unused~ |" & Ltx
            End If
        End If
    Next Lc
    For Lc = 0 To ListProcedure.ListCount - 1
        Ltx = ListProcedure.List(Lc) & NewLine
        If Check(5).Value = 1 And Left(Ltx, 11) = "'~Unused~ |" Then Ltx = vbNullString
        ScanProcedure = ScanProcedure & Ltx ' returns the string corrected
    Next Lc
    ScanProcedure = Left(ScanProcedure, Len(ScanProcedure) - 2) ' Removes the last line feed.
    Exit Function
ScanProcedure_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(ScanProcedure) of (frmVBProjectCleaner.frm).", vbCritical

End Function
Private Function WriteProcedureErrorHandle(ProcName As String, FName As String) As String

' Verifying the entire procedure
Dim Lc As Long
Dim WPos As Long
Dim Ltx As String
Dim Header As String
Dim Footer As String
Dim ProcType As String
Dim HasErrorHandle As Boolean
    Header = vbNullString

    Footer = vbNullString
    WriteProcedureErrorHandle = vbNullString
    HasErrorHandle = False
    Ltx = ListProcedure.List(0)
    WPos = InStr(Ltx, "Function ")
    If WPos > 0 Then
        ProcType = "Function"
        Else: WPos = InStr(Ltx, "Sub ")
        If WPos > 0 Then
            ProcType = "Sub"
            Else: WPos = InStr(Ltx, "Property ")
            If WPos > 0 Then ProcType = "Property"
        End If
    End If
    For Lc = 0 To ListProcedure.ListCount - 1
        Ltx = ListProcedure.List(Lc)
        WPos = InStr(Ltx, "on error ")
        If WPos > 0 Then
            HasErrorHandle = True
            Exit For
        End If
    Next Lc
    If HasErrorHandle = False Then
        Header = NewLine & "on error GoTo " & ProcName & "_Error" & NewLine
        Footer = NewLine & "Exit " & ProcType & NewLine & NewLine & ProcName & "_Error:" & NewLine & NewLine & "MsgBox ""Error "" & Err.Number & "" ("" & Err.Description & "") in procedure"" & Chr(13) & Chr(10) & ""(" & ProcName & ") of (" & FName & ")."", vbCritical" & NewLine & NewLine
    End If
    For Lc = 0 To ListProcedure.ListCount - 1
        Ltx = ListProcedure.List(Lc) & NewLine
        If Lc = 1 Then Ltx = Header & Ltx
        If Lc = ListProcedure.ListCount - 2 Then Ltx = Ltx & Footer
        WriteProcedureErrorHandle = WriteProcedureErrorHandle & Ltx ' returns the string corrected
    Next Lc
    WriteProcedureErrorHandle = Left(WriteProcedureErrorHandle, Len(WriteProcedureErrorHandle) - 2) ' Removes the last line feed.

End Function
Private Function ScanPrivate() As String

On Error GoTo ScanPrivate_Error
Dim Lc As Long
Dim Sc As Long
Dim Tc As Long
Dim VName As String
Dim Ltx As String
Dim Xtx As String
Dim IsVar As Boolean
Dim VarCt As Long
Dim LineTx As String
Dim InProcedure As Boolean
    ScanPrivate = vbNullString

    ProgressBar.Value = 0
    Me.Caption = "VB Project Cleaner"
    ListNames.Clear
    ListIndex.Clear
    For Sc = 0 To ListProcedure.ListCount - 1
        Ltx = ListProcedure.List(Sc)
        VName = GetDecls(Ltx, "Private")
        If VName <> vbNullString Then
            TotPrivate = TotPrivate + 1
            LabPrivate.Caption = TotPrivate
            ListNames.AddItem VName
            ListIndex.AddItem Sc
        End If
    Next Sc
    If ListNames.ListCount > 0 Then
        ListN.Clear
        For Lc = 0 To ListNames.ListCount - 1
            ListN.AddItem 0
        Next Lc
        For Sc = 0 To ListProcedure.ListCount - 1
            Xtx = ListProcedure.List(Sc)
            VarCt = 0
            For Lc = 0 To ListNames.ListCount - 1
                VName = ListNames.List(Lc)
                IsVar = IsVarTrue(Xtx, VName)
                If IsVar = True Then ListN.List(Lc) = CInt(ListN.List(Lc)) + 1
            Next Lc
            DoEvents
        Next Sc
        For Lc = 0 To ListN.ListCount - 1
            If CInt(ListN.List(Lc)) < 2 Then
                UnusedPrivate = UnusedPrivate + 1
                LabUnusedPrivate.Caption = UnusedPrivate
                LineTx = ListProcedure.List(ListIndex.List(Lc))
                If Left(LineTx, 13) = "Private Type " Or Left(LineTx, 12) = "Private Sub " Or Left(LineTx, 17) = "Private Function " Or Left(LineTx, 17) = "Private Property " Then
                    InProcedure = True
                    Else: InProcedure = False
                End If
                ListProcedure.List(ListIndex.List(Lc)) = "'~~Unused~~ |" & LineTx
                If InProcedure = True Then
                    For Tc = ListIndex.List(Lc) + 1 To ListProcedure.ListCount - 1
                        LineTx = ListProcedure.List(Tc)
                        ListProcedure.List(Tc) = "'~~~~~~~~~~ |" & LineTx
                        If Left(LineTx, 7) = "End Sub" Or Left(LineTx, 8) = "End Type" Or Left(LineTx, 12) = "End Function" Or Left(LineTx, 12) = "End Property" Then
                            InProcedure = False
                            Exit For
                        End If
                    Next Tc
                End If
            End If
            DoEvents
            SetProgressBar ListN.ListCount, Lc
        Next Lc
    End If
    For Lc = 0 To ListProcedure.ListCount - 1
        Ltx = ListProcedure.List(Lc) & NewLine
        If Check(5).Value = 1 And Left(Ltx, 3) = "'~~" Then Ltx = vbNullString
        ScanPrivate = ScanPrivate & Ltx ' returns the string corrected
    Next Lc
    If ScanPrivate <> vbNullString Then ScanPrivate = Left(ScanPrivate, Len(ScanPrivate) - 2) ' Removes the last line feed.
    ProgressBar.Value = 0
    Me.Caption = "VB Project Cleaner"
    Exit Function
ScanPrivate_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(ScanPrivate) of (frmVBProjectCleaner.frm).", vbCritical

End Function
Private Sub ScanPublic()

On Error GoTo ScanPublic_Error
Dim Lc As Long
Dim Sc As Long
Dim VName As String
Dim Ltx As String
Dim Xtx As String
Dim IsVar As Boolean
Dim VarCt As Long
    ProgressBar.Value = 0

    Me.Caption = "VB Project Cleaner"
    ListNames.Clear
    ListIndex.Clear
    For Sc = 0 To ListProcedure.ListCount - 1
        Ltx = ListProcedure.List(Sc)
        VName = GetDecls(Ltx, "Public")
        If VName = vbNullString Then VName = GetDecls(Ltx, "Static")
        If VName <> vbNullString Then
            TotPublic = TotPublic + 1
            LabPublic.Caption = TotPublic
            ListNames.AddItem VName
            ListIndex.AddItem Sc
        End If
    Next Sc
    If ListNames.ListCount = 0 Then Exit Sub
    ListN.Clear
    ListNSource.Clear
    For Lc = 0 To ListNames.ListCount - 1
        ListN.AddItem 0
        ListNSource.AddItem 0
    Next Lc
    For Sc = 0 To ListProcedure.ListCount - 1
        Xtx = ListProcedure.List(Sc)
        VarCt = 0
        For Lc = 0 To ListNames.ListCount - 1
            VName = ListNames.List(Lc)
            IsVar = IsVarTrue(Xtx, VName)
            If IsVar = True Then ListNSource.List(Lc) = CInt(ListNSource.List(Lc)) + 1
        Next Lc
        DoEvents
    Next Sc
    Exit Sub
ScanPublic_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(ScanPublic) of (frmVBProjectCleaner.frm).", vbCritical

End Sub
Private Function GetDecls(Dex As String, DType As String) As String

On Error GoTo GetDecls_Error

Dim Dxx As Long
Dim Xxx As String
Dim XName As String
Dim XType As String
Dim XPos As Long
Dim YPos As Long
Dim IsEvt As Boolean

    GetDecls = vbNullString
    Xxx = " " & Dex & " "
    XType = " " & DType & " "
    XPos = InStr(Xxx, XType)
    If XPos <> 0 Then
        YPos = InStr(Xxx, "'") ' Is comment?
        If YPos <> 0 And YPos < XPos Then Exit Function ' Yes, it is.
        Else: Exit Function
    End If
    YPos = InStr(Xxx, """") ' Is text?
    If YPos > 0 And YPos < XPos Then Exit Function
    YPos = InStr(Xxx, "Sub Main")
    If YPos > 0 Then Exit Function
' If one of the following occurrences is found in the line, probably we have a declaration
    XName = Mid(Xxx, XPos + Len(XType), Len(Xxx) - (XPos + Len(XType) - 1))
    If Left(XName, 4) = "Sub " Then XName = Right(XName, Len(XName) - 4)
    If Left(XName, 5) = "Type " Then XName = Right(XName, Len(XName) - 5)
    If Left(XName, 6) = "Const " Then XName = Right(XName, Len(XName) - 6)
    If Left(XName, 9) = "Function " Then XName = Right(XName, Len(XName) - 9)
    If Left(XName, 13) = "Property Get " Then XName = Right(XName, Len(XName) - 13)
    If Left(XName, 13) = "Property Set " Then XName = Right(XName, Len(XName) - 13)
    If Left(XName, 9) = "Property " Then XName = Right(XName, Len(XName) - 9)
    If Left(XName, 12) = "Declare Sub " Then XName = Right(XName, Len(XName) - 12)
    If Left(XName, 17) = "Declare Function " Then XName = Right(XName, Len(XName) - 17)
    XName = XName & " "
    For Dxx = 1 To Len(XName) ' Isolates the name of the declaration
        If Mid(XName, Dxx, 1) = " " Or Mid(XName, Dxx, 1) = "(" Or Mid(XName, Dxx, 1) = "'" Or _
        Mid(XName, Dxx, 1) = "!" Or Mid(XName, Dxx, 1) = "%" Or Mid(XName, Dxx, 1) = "&" Or _
        Mid(XName, Dxx, 1) = "#" Or Mid(XName, Dxx, 1) = "@" Or Mid(XName, Dxx, 1) = "$" Then
            XName = Left(XName, Dxx - 1)
            Exit For
        End If
    Next Dxx
    If XName = vbNullString Or XName = Space(Len(XName)) Then Exit Function
    IsEvt = IsEvent(XName) ' Is the declaration the name of an object?
    If IsEvt = True Then Exit Function
    GetDecls = XName ' The name of the declaration is returned.
    Exit Function
GetDecls_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(GetDecls) of (frmVBProjectCleaner.frm).", vbCritical

End Function
Private Function VerifyLine(Dex As String, DZone As Boolean) As String

On Error GoTo VerifyLine_Error
Dim Cxx As Long
Dim Xxx As String
Dim NVar As String
Dim Gpos As Long
Dim HPos As Long
Dim RetX As Boolean
    Xxx = " " & Dex & " "

    If DZone = True Then ' In declarations zone.
        Gpos = InStr(Xxx, " Dim ") 'Is a Dim declaration?
        If Gpos > 0 Then ' Yes.
            Dex = Replace(Dex, "Dim ", "Private ") 'Replaces Dim to Private
            Xxx = " " & Dex & " " ' our working string takes the new contents.
        End If
    End If
    For Cxx = 0 To ListVarType.ListCount - 1    ' This list keeps the several
        NVar = ListVarType.List(Cxx) & " "      ' kinds of declarations.
        Gpos = InStr(Xxx, NVar) 'There is a declaration is this line?
        If Gpos <> 0 Then ' Looks like a declaration...
            HPos = InStr(Xxx, "'") ' The line is comment?
            If HPos > 0 And HPos < Gpos Then ' Yes, it is.
                VerifyLine = Dex  ' the line is returned untouched.
                Exit Function
            End If
            Exit For ' the line pass all this tests
        End If
    Next Cxx
    If Gpos > 0 Then ' GPos > 0 - the declaration was founded in the previous cicle
        RetX = IsText(Xxx) ' now, we see if declaration is text (between "")
        If RetX = False Then
            RetX = IsInProced(Xxx) ' Is declaration between ()? Like arguments in a procedure statement?
' if not, there is a declaration write in the format: (Dim A as Long, B as Byte, C)
' We rewrite them, one for line.
            If RetX = False Then Dex = Replace(Dex, ", ", NewLine & NVar)
        End If
    End If
    VerifyLine = Dex
    Exit Function
VerifyLine_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(VerifyLine) of (frmVBProjectCleaner.frm).", vbCritical

End Function
Private Function IsText(Cxt As String) As Boolean

On Error GoTo IsText_Error
' verifies if the occurrence " ," is inside of diaeresis ("). In this case is text, not code
Dim XPos As Long
Dim YPos As Long
Dim ZPos As Long
Dim Cxx As String
    Cxx = " " & Cxt & " "

    IsText = True
    XPos = InStr(Cxx, ", ")
    If XPos > 0 Then
        YPos = InStr(Cxx, " """)
        If YPos > 0 And YPos < XPos Then
            ZPos = InStr(Cxx, """ ")
            If ZPos > 0 And ZPos > YPos Then Exit Function
        End If
    End If
    XPos = InStr(Cxx, "Dim ")
    If XPos = 0 Then XPos = InStr(Cxx, "Const ")
    If XPos = 0 Then XPos = InStr(Cxx, "Static ")
    If XPos = 0 Then XPos = InStr(Cxx, "Private ")
    If XPos = 0 Then XPos = InStr(Cxx, "Public ")
    If XPos = 0 Then Exit Function
    IsText = False
    Exit Function
IsText_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(IsText) of (frmVBProjectCleaner.frm).", vbCritical

End Function
Private Function IsInProced(Cx As String) As Boolean

On Error GoTo IsInProced_Error
' verifies if the occurrence " ," is inside parenthesis( ).
'In that case it is part of the arguments of a declaration.
Dim XPos As Long
Dim YPos As Long
Dim ZPos As Long
    IsInProced = True

    XPos = InStr(Cx, ", ")
    If XPos > 0 Then
        YPos = InStr(Cx, "(")
        If YPos > 0 And YPos < XPos Then
            ZPos = InStr(Cx, ")")
            If ZPos > 0 And ZPos > YPos Then Exit Function
        End If
    End If
    IsInProced = False
    Exit Function
IsInProced_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(IsInProced) of (frmVBProjectCleaner.frm).", vbCritical

End Function
Private Function IsEvent(Dec As String) As Boolean

On Error GoTo IsEvent_Error
    IsEvent = False
    If ListObjects.ListCount = 0 Then Exit Function
Dim Ia As Integer
Dim YPos As Long

    For Ia = 0 To ListObjects.ListCount - 1
        YPos = InStr(UCase(Dec), UCase(ListObjects.List(Ia)))
        If YPos <> 0 Then
            IsEvent = True
            Exit Function
        End If
    Next Ia
    Exit Function
IsEvent_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(IsEvent) of (frmVBProjectCleaner.frm).", vbCritical

End Function
Private Function OpenVBP() As Boolean

'Opens the VB Project
    With Dialog
        .CancelError = True
        .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
On Error GoTo ErrHandler
        .Filter = "Ficheiros VBP (*.vbp)|*.vbp"
        .ShowOpen
        VBPName = .FileName
    End With
    OpenVBP = True
    Exit Function
ErrHandler:
    If ListModules.ListCount > 0 Then
        ChecksEnabled True
        Else: ChecksEnabled False
    End If
    OpenVBP = False
    Err.Clear

End Function
Private Function GetPath(Nx As String) As String

On Error GoTo GetPath_Error
' Return in "GetPath" the Pathname without filename and ending with "\"
Dim i As Integer
Dim j As Integer
    ProjectFolder = vbNullString

    For i = Len(Nx) To 1 Step -1
        If Mid(Nx, i, 1) = "\" Then
            GetPath = Left(Nx, i)
            If i > 3 Then
                For j = i - 1 To 1 Step -1
                    If Mid(Nx, j, 1) = "\" Then
                        ProjectFolder = Mid(Nx, j + 1, (i - j) - 1)
                        Exit Function
                    End If
                Next j
            End If
            Exit Function
        End If
    Next i
    Exit Function
GetPath_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(GetPath) of (frmVBProjectCleaner.frm).", vbCritical

End Function
Private Sub GetModulesNames(FName As String)

On Error GoTo GetModulesNames_Error
' "ListModules" obtains the names of all modules, classmodules e forms of the project.
Dim Txt As String
Dim DPos As Long
Dim LenLine As Long
Dim LenFic As Long
    ListModules.Clear

    LenFic = FileLen(FName)
    ProgressBar.Value = 0
    Me.Caption = "VB Project Cleaner"
    Set Fs = CreateObject("Scripting.FileSystemObject")
    Set f = Fs.GetFile(FName)
    Set ts = f.OpenAsTextStream(1, -2)
    LenLine = 0
    Do While ts.AtEndOfStream <> True
        Txt = ts.Readline
        LenLine = LenLine + Len(Txt) + 2
        If Left(Txt, 5) = "Form=" Or Left(Txt, 13) = "PropertyPage=" Or _
            Left(Txt, 9) = "Designer=" Or Left(Txt, 12) = "UserControl=" Then
            DPos = InStr(Txt, "=")
            ListModules.AddItem Right(Txt, Len(Txt) - DPos)
        End If
        If Left(Txt, 7) = "Module=" Or Left(Txt, 6) = "Class=" Then
            DPos = InStr(Txt, "; ")
            ListModules.AddItem Right(Txt, Len(Txt) - (DPos + 1))
        End If
        DoEvents
        SetProgressBar LenFic, LenLine
    Loop
    ts.Close
    If ListModules.ListCount = 1 Then ListModules.ListIndex = 0
    ProgressBar.Value = 0
    Me.Caption = "VB Project Cleaner"
    Exit Sub
GetModulesNames_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(GetModulesNames) of (frmVBProjectCleaner.frm).", vbCritical

End Sub
Private Sub SetProgressBar(Fx As Long, Lx As Long)

On Error GoTo SetProgressBar_Error
Dim Vap As Integer
    Vap = CInt((ProgressBar.Max * Lx) / Fx)

    If Vap > ProgressBar.Max Then Vap = ProgressBar.Max
    ProgressBar.Value = Vap
    Me.Caption = "VB Project Cleaner - " & Vap & " %"
    Exit Sub
SetProgressBar_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(SetProgressBar) of (frmVBProjectCleaner.frm).", vbCritical

End Sub
Private Sub EmptyLabels()

On Error GoTo EmptyLabels_Error
    LabDim.Caption = vbNullString
    LabUnusedDim.Caption = vbNullString
    LabPrivate.Caption = vbNullString
    LabUnusedPrivate.Caption = vbNullString
    LabPublic.Caption = vbNullString
    LabUnusedPublic.Caption = vbNullString
    Exit Sub
EmptyLabels_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(EmptyLabels) of (frmVBProjectCleaner.frm).", vbCritical

End Sub
Private Sub ChecksEnabled(Enbl As Boolean)

On Error GoTo ChecksEnabled_Error
Dim Ch As Integer
    ComDetectUnusedDeclarations.Enabled = Enbl

    For Ch = 0 To 9
        Check(Ch).Enabled = Enbl
    Next Ch
    ChkEnable = Enbl
    If ChkEnable = True Then
        If Check(2).Value = 0 And Check(3).Value = 0 And Check(4).Value = 0 Then
            Check(5).Enabled = False
            Check(6).Enabled = False
        End If
    End If
    Exit Sub
ChecksEnabled_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(ChecksEnabled) of (frmVBProjectCleaner.frm).", vbCritical

End Sub
Private Function IsVarTrue(Stxt As String, Vtxt As String) As Boolean

On Error GoTo IsVarTrue_Error
Dim TPos As Long
Dim VPos As Long
Dim SSx As String
Dim Vxx As String
    
    SSx = Stxt
    Vxx = Vtxt
    IsVarTrue = False
    VPos = InStr(SSx, "_Error")
    If VPos > 0 Then Exit Function
    VPos = InStr(SSx, Vxx)
    If VPos = 0 Then Exit Function
    TPos = InStr(SSx, "'") ' verifies if the line is a comment, not code.
    If TPos > 0 And TPos < VPos Then Exit Function ' yes, it is.
' The following conditions guarantee the efective usage of the declaration
' and nothing else, like part of another name, text, etc.
    IsVarTrue = True
    If SSx = Vxx Then Exit Function
    TPos = InStr(SSx, "ReDim " & Vxx)
    If TPos > 0 Then Exit Function
    If Left(SSx, Len(Vxx) + 1) = Vxx & " " Then Exit Function
    If Right(SSx, Len(Vxx) + 1) = " " & Vxx Then Exit Function
    SSx = " " & SSx & " "
    TPos = InStr(SSx, " " & Vxx & " ")
    If TPos <> 0 Then Exit Function
    TPos = InStr(SSx, "(" & Vxx & ".")
    If TPos <> 0 Then Exit Function
    TPos = InStr(SSx, "." & Vxx & ")")
    If TPos <> 0 Then Exit Function
    TPos = InStr(SSx, "(" & Vxx & ",")
    If TPos <> 0 Then Exit Function
    TPos = InStr(SSx, "(" & Vxx & ")")
    If TPos <> 0 Then Exit Function
    TPos = InStr(SSx, "(" & Vxx & "(")
    If TPos <> 0 Then Exit Function
    TPos = InStr(SSx, "(" & Vxx & " ")
    If TPos <> 0 Then Exit Function
    TPos = InStr(SSx, " " & Vxx & ")")
    If TPos <> 0 Then Exit Function
    TPos = InStr(SSx, " " & Vxx & "(")
    If TPos <> 0 Then Exit Function
    TPos = InStr(SSx, " " & Vxx & ".")
    If TPos <> 0 Then Exit Function
    TPos = InStr(SSx, " " & Vxx & ",")
    If TPos <> 0 Then Exit Function
    TPos = InStr(SSx, "," & Vxx & " ")
    If TPos <> 0 Then Exit Function
    TPos = InStr(SSx, "." & Vxx & " ")
    If TPos <> 0 Then Exit Function
    TPos = InStr(SSx, "." & Vxx & ".")
    If TPos <> 0 Then Exit Function
    TPos = InStr(SSx, "," & Vxx & ".")
    If TPos <> 0 Then Exit Function
    TPos = InStr(SSx, "." & Vxx & ",")
    If TPos <> 0 Then Exit Function
    TPos = InStr(SSx, "_" & Vxx & "(")
    If TPos <> 0 Then Exit Function
    IsVarTrue = False
    Exit Function
IsVarTrue_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(IsVarTrue) of (frmVBProjectCleaner.frm).", vbCritical

End Function
Private Sub CopyFile(SourceName As String, DestName As String)

On Error GoTo CopyFile_Error
' File copy routine
    Me.MousePointer = vbHourglass
    Set Fs = CreateObject("Scripting.FileSystemObject")
    Set f = Fs.GetFile(SourceName)
    f.Copy DestName, True
    Me.MousePointer = vbNormal
    Exit Sub
CopyFile_Error:
    Err.Clear
    If UCase(Right(DestName, 3)) = "FRX" Or UCase(Right(DestName, 3)) = "CTX" Then Exit Sub
    MsgBox "File not found: " & SourceName, vbInformation
End Sub
Private Sub IniDeclarationsType()

On Error GoTo IniDeclarationsType_Error
    ListVarType.Clear
' Note: Not a random order. Decreasing order of string's size. I forgot some?
    ListVarType.AddItem "Private Declare Function"
    ListVarType.AddItem "Public Declare Function"
    ListVarType.AddItem "Private Declare Function"
    ListVarType.AddItem "Public Declare Function"
    ListVarType.AddItem "Private Declare Sub"
    ListVarType.AddItem "Public Declare Sub"
    ListVarType.AddItem "Private Property"
    ListVarType.AddItem "Private Function"
    ListVarType.AddItem "Declare Function"
    ListVarType.AddItem "Public Property"
    ListVarType.AddItem "Public Function"
    ListVarType.AddItem "Private Const"
    ListVarType.AddItem "Public Const"
    ListVarType.AddItem "Private Type"
    ListVarType.AddItem "Public Type"
    ListVarType.AddItem "Private Sub"
    ListVarType.AddItem "Declare Sub"
    ListVarType.AddItem "Public Sub"
    ListVarType.AddItem "Property"
    ListVarType.AddItem "Function"
    ListVarType.AddItem "Private"
    ListVarType.AddItem "Public"
    ListVarType.AddItem "Static"
    ListVarType.AddItem "Const"
    ListVarType.AddItem "Type"
    ListVarType.AddItem "Dim"
    ListVarType.AddItem "Sub"
    Exit Sub
IniDeclarationsType_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(IniDeclarationsType) of (frmVBProjectCleaner.frm).", vbCritical

End Sub
Private Sub Form_Load()

On Error GoTo Form_Load_Error
    ProgressBar.Max = 100
    NewLine = Chr(13) & Chr(10)
    IniDeclarationsType
    UserFile = App.Path
    If Right(UserFile, 1) <> "\" Then UserFile = UserFile & "\"
    UserFile = UserFile & "UserData.txt"
    GetUserData ' Obtains the user options file
    OldCheckValue = Check(5).Value
    Exit Sub
Form_Load_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(Form_Load) of (frmVBProjectCleaner.frm).", vbCritical

End Sub
Private Sub Form_Unload(Cancel As Integer)

On Error GoTo Form_Unload_Error
    SaveUserData
    Me.MousePointer = vbNormal
    Set Fs = Nothing
    Set f = Nothing
    Set ts = Nothing
    Set FrmVBProjectCleaner = Nothing
    End
    Exit Sub
Form_Unload_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(Form_Unload) of (frmVBProjectCleaner.frm).", vbCritical

End Sub
Private Sub ListModules_Click()

On Error GoTo ListModules_Click_Error

    If Check(1).Value = 1 Then
        ModuleIndex = ListModules.ListIndex
        Else: ListModules.ListIndex = ModuleIndex
    End If
    Exit Sub
ListModules_Click_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(ListModules_Click) of (frmVBProjectCleaner.frm).", vbCritical

End Sub
Private Sub mnuFileExit_Click()

On Error GoTo mnuFileExit_Click_Error
    Unload Me
    Exit Sub
mnuFileExit_Click_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(mnuFileExit_Click) of (frmVBProjectCleaner.frm).", vbCritical

End Sub
Private Sub mnuFileOpenProject_Click()

On Error GoTo mnuFileOpenProject_Click_Error
    ComOpenVBP_Click
    Exit Sub
mnuFileOpenProject_Click_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(mnuFileOpenProject_Click) of (frmVBProjectCleaner.frm).", vbCritical

End Sub
Private Sub mnuHelpContents_Click()

' Displays our small help
Dim HelpFile As String
Dim hwndHelp As Long
    HelpFile = App.Path

    If Right(HelpFile, 1) <> "\" Then HelpFile = HelpFile & "\"
    HelpFile = HelpFile & "VBProjectCleaner.chm"
    If FileLen(HelpFile) = 0 Then
        MsgBox "Help file missing or corrupted!", vbInformation, Me.Caption
        Else
        On Error Resume Next
        hwndHelp = HtmlHelp(hWnd, HelpFile, &H0, 0)
        If Err Then MsgBox Err.Description
    End If

End Sub
Private Sub mnuOptionsDetect_Click()

On Error GoTo mnuOptionsDetect_Click_Error
    ComDetectUnusedDeclarations_Click
    Exit Sub
mnuOptionsDetect_Click_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure" & Chr(13) & Chr(10) & "(mnuOptionsDetect_Click) of (frmVBProjectCleaner.frm).", vbCritical

End Sub
