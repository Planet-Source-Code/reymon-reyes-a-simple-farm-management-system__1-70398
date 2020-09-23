VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00FEFCDE&
   Caption         =   "Amadeus Automated Stocks Processing System |ASPS|"
   ClientHeight    =   7065
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10545
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":08CA
   ScrollBars      =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imglstToolbar 
      Left            =   1680
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":17E02C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":17EF06
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":17FD58
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":181A62
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbarMain 
      Align           =   3  'Align Left
      Height          =   6660
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   11748
      ButtonWidth     =   1958
      ButtonHeight    =   1799
      Appearance      =   1
      ImageList       =   "imglstToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Inventory"
            Key             =   "inventory"
            Object.ToolTipText     =   "Inventory"
            ImageIndex      =   2
            Value           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stock Details"
            Key             =   "stockdetails"
            Object.ToolTipText     =   "Stock Details"
            ImageIndex      =   1
            Value           =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Request Form"
            Key             =   "requestform"
            Object.ToolTipText     =   "Request Form"
            ImageIndex      =   3
            Value           =   1
         EndProperty
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "MDIForm1.frx":18273C
      Begin VB.PictureBox Picture3 
         Height          =   495
         Left            =   120
         Picture         =   "MDIForm1.frx":182A56
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   5
         Top             =   4440
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Height          =   495
         Left            =   120
         Picture         =   "MDIForm1.frx":183616
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   4
         Top             =   3840
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Height          =   495
         Left            =   120
         Picture         =   "MDIForm1.frx":1844B8
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   3
         Top             =   3240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   6570
         Left            =   -120
         Picture         =   "MDIForm1.frx":1854FA
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   1230
      End
   End
   Begin MSComDlg.CommonDialog dlgBackup 
      Left            =   240
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.StatusBar StatusBar2 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   6690
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   6
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2469
            MinWidth        =   2469
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "Today is:"
            TextSave        =   "Today is:"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "3/16/2008"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "CAPS"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NUM"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   30
      Left            =   0
      TabIndex        =   0
      Top             =   6660
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   53
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1200
      Top             =   5640
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnufieldman 
         Caption         =   " L&eadman"
      End
      Begin VB.Menu mnugrower 
         Caption         =   " &Grower"
      End
      Begin VB.Menu mnuSuppliers 
         Caption         =   " Su&ppliers"
      End
      Begin VB.Menu separator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuinventory 
         Caption         =   " I&nventory"
      End
      Begin VB.Menu mnuStockDetails 
         Caption         =   " Stock Details"
      End
      Begin VB.Menu mnuPurchaseOrder 
         Caption         =   " Purchase Order"
      End
      Begin VB.Menu mnuRequestForm 
         Caption         =   " &Request Form"
      End
      Begin VB.Menu mnuReorderStatus 
         Caption         =   " Reorder Status"
      End
      Begin VB.Menu mnuItemUsage 
         Caption         =   " Item Usage"
      End
      Begin VB.Menu separator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnusetting 
         Caption         =   "&Settings"
         Begin VB.Menu mnustocktypes 
            Caption         =   " Classification && &Kind"
         End
         Begin VB.Menu mnuunitmeasure 
            Caption         =   " &Unit of Measure"
         End
      End
      Begin VB.Menu separator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   " E&xit"
      End
   End
   Begin VB.Menu mnureports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuReportGrower 
         Caption         =   " Grower"
      End
      Begin VB.Menu mnuReportStockItem 
         Caption         =   " Stock Items"
      End
      Begin VB.Menu mnuReportWithdrawals 
         Caption         =   " Withdrawals"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuUtilities 
         Caption         =   "Utilities"
         Begin VB.Menu mnuBackup 
            Caption         =   " Backup Database"
         End
         Begin VB.Menu mnuCompactDatabase 
            Caption         =   " Compact Database"
         End
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSecurity 
         Caption         =   "Security"
         Begin VB.Menu mnuChangePassword 
            Caption         =   " Change Password"
         End
         Begin VB.Menu mnuAddNewUser 
            Caption         =   " Add New User"
            Visible         =   0   'False
         End
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuCascade 
         Caption         =   " Cascade"
      End
   End
   Begin VB.Menu mnuabout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long

Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long

Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long



Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_MAXIMIZEBOX = &H10000
Public Sub NoMaxBox(F As MDIForm)
Dim l As Long
l = GetWindowLong(F.hwnd, GWL_STYLE)
l = l And Not (WS_MAXIMIZEBOX)
l = SetWindowLong(F.hwnd, GWL_STYLE, 1)
End Sub
Public Sub NoMinBox(F As MDIForm)
Dim l As Long
l = GetWindowLong(F.hwnd, GWL_STYLE)
l = l And Not (WS_MINIMIZEBOX)
l = SetWindowLong(F.hwnd, GWL_STYLE, 1)
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
tbarMain.Buttons(1).Value = tbrPressed
tbarMain.Buttons(2).Value = tbrPressed
tbarMain.Buttons(3).Value = tbrPressed
End Sub

Private Sub MDIForm_Initialize()
envAmadeus.conReports.ConnectionString = "Provider=MSDataShape.1;Extended Properties=Jet OLEDB:Database Password=a;Persist Security Info=False;Data Source=" & App.Path & "\Database\AmadeusFarm.mdb;Data Provider=MICROSOFT.JET.OLEDB.4.0"
End Sub

Private Sub MDIForm_Load()
On Error Resume Next

Dim lngMenu As Long
Dim lngSubMenu As Long
Dim lngMenuItemID As Long
Dim lngRet As Long

StatusBar2.Panels(6).Text = modAdoconnect.user
frmReorderStatus.Show

lngMenu = GetMenu(MDIForm1.hwnd)
lngSubMenu = GetSubMenu(lngMenu, 0)
lngMenuItemID = GetMenuItemID(lngSubMenu, 4)
lngRet = SetMenuItemBitmaps(lngMenu, lngMenuItemID, 0, _
Picture1.Picture, Picture1.Picture)

lngMenu = GetMenu(MDIForm1.hwnd)
lngSubMenu = GetSubMenu(lngMenu, 0)
lngMenuItemID = GetMenuItemID(lngSubMenu, 5)
lngRet = SetMenuItemBitmaps(lngMenu, lngMenuItemID, 0, _
Picture2.Picture, Picture2.Picture)

lngMenu = GetMenu(MDIForm1.hwnd)
lngSubMenu = GetSubMenu(lngMenu, 0)
lngMenuItemID = GetMenuItemID(lngSubMenu, 7)
lngRet = SetMenuItemBitmaps(lngMenu, lngMenuItemID, 0, _
Picture3.Picture, Picture3.Picture)

End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
tbarMain.Buttons(1).Value = tbrPressed
tbarMain.Buttons(2).Value = tbrPressed
tbarMain.Buttons(3).Value = tbrPressed
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
If CheckUpdateMode = True Then
    MsgBox "Please finish all pending operations before exiting."
    Cancel = 1
    Exit Sub
End If
If MsgBox("Are you sure you want to close the program now?" & vbCr & _
          " By: The_Etc Groups", vbInformation + vbYesNo, "Warning!") = vbYes Then
    UnloadMode = vbFormControlMenu
    Unload Me
    Cancel = 0
Else
    Cancel = 1
End If

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
On Error Resume Next
If envAmadeus.conReports.State = adStateOpen Then
    envAmadeus.conReports.Close
End If
End Sub

Private Sub mnuabout_Click()
frmAbout.Show
End Sub

Private Sub mnuBackup_Click() 'database file backup routine
Dim Source As String, Destination As String, backupdate As String
Dim cpyFSO As New FileSystemObject
Dim cpyForce As Boolean

On Error GoTo BackupErrHandler

'set the dialog box to generate error when user pressed cancel
dlgBackup.CancelError = True
'format the date for appending to the destination path
backupdate = Format(Date, "mm-dd-yyyy")
'set path for the source of the db file
Source = App.Path & "\Database\AmadeusFarm.mdb"
'set db file name
Destination = "AmadeusFarm " & backupdate
'set dialog properties
dlgBackup.FileName = Destination
dlgBackup.DefaultExt = "mdb"
dlgBackup.DialogTitle = "Backup Database"
dlgBackup.Filter = "Access Database (*.mdb)|*.mdb"
dlgBackup.Flags = cdlOFNOverwritePrompt
dlgBackup.ShowSave
'backup the file
If cpyFSO.FileExists(dlgBackup.FileName) = True Then
    cpyForce = True
Else
    cpyForce = False
End If
Me.MousePointer = vbHourglass
cpyFSO.CopyFile Source, dlgBackup.FileName, cpyForce
MsgBox "Backup complete."
Me.MousePointer = vbDefault
Exit Sub
BackupErrHandler:
    'MsgBox Err.Number & " " & Err.Description
    Exit Sub
End Sub

Private Sub mnuCascade_Click()
MDIForm1.Arrange vbCascade
End Sub

Private Sub mnuChangePassword_Click()
frmChangePassword.Show
End Sub

Private Sub mnuCompactDatabase_Click()
Dim Source As String, Destination As String
Dim srcConnection As String, destConnection As String
Dim cpyFSO As New FileSystemObject
Dim cpyForce As Boolean
Dim jetDB As New JRO.JetEngine
Dim conReports As String

On Error GoTo ErrHandler

If Forms.Count > 1 Then
    MsgBox "Please close all open forms to proceed."
'    Exit Sub
'Else
 '   MsgBox "Its ok to proceed."
    Exit Sub
End If

conReports = envAmadeus.conReports.ConnectionString

envAmadeus.conReports.Close

Source = App.Path & "\Database\AmadeusFarm.mdb"
Destination = App.Path & "\Database\AmadeusFarmTemp.mdb"
srcConnection = "Data Source=" & Destination & ";Jet OLEDB:Database Password=a"
destConnection = "Data Source=" & App.Path & "\Database\AmadeusFarmCompact.mdb;Jet OLEDB:Database Password=a"

If MsgBox("Do you want to proceed with the Compact Database operation?", vbQuestion + vbYesNo, "Compact Database") = vbYes Then
    Me.MousePointer = vbHourglass
    'copy the original db file
    cpyFSO.CopyFile Source, Destination
    'compact the copy version
    jetDB.CompactDatabase srcConnection, destConnection
    'delete the copy version
    cpyFSO.DeleteFile Destination, True
    'copy the compacted file and overwrite the original file
    cpyFSO.CopyFile App.Path & "\Database\AmadeusFarmCompact.mdb", Source, True
    'delete the compacted db file
    cpyFSO.DeleteFile App.Path & "\Database\AmadeusFarmCompact.mdb", True
    Me.MousePointer = vbDefault
    MsgBox "Compact Database successfully finished."
End If

envAmadeus.conReports.Open conReports

Exit Sub
ErrHandler:
    MsgBox Err.Number & " " & Err.Description

End Sub

Private Sub mnuexit_Click()
Unload Me
End Sub

Private Sub mnufieldman_Click()
frmLeadman.Show
End Sub

Private Sub mnugrower_Click()
frmGrower.Show
End Sub

Private Sub mnuinventory_Click()
frmInventory.Show
End Sub

Private Sub mnuItemUsage_Click()
frmUsage.Show
End Sub

Private Sub mnuPurchaseOrder_Click()
frmPO.Show
End Sub

Private Sub mnuReorderStatus_Click()
frmReorderStatus.Show
End Sub

Private Sub mnuReportGrower_Click()
frmRptGrower.Show
End Sub

Private Sub mnuReportStockItem_Click()
frmRptStockItem.Show
End Sub

Private Sub mnuReportWithdrawals_Click()
frmRptWithdrawals.Show
End Sub

Private Sub mnuRequestForm_Click()
frmRequest.Show
End Sub

Private Sub mnuStockDetails_Click()
frmStockDetails.Show
End Sub

Private Sub mnustocktypes_Click()
frmClassification.Show
End Sub

Private Sub mnuSuppliers_Click()
frmSuppliers.Show
End Sub

Private Sub mnuunitmeasure_Click()
frmUnitofMeasurement.Show
End Sub

Private Sub tbarMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If (X >= 0 And Y <= 1000) Then
    tbarMain.Buttons(1).Value = tbrUnpressed
    tbarMain.Buttons(2).Value = tbrPressed
    tbarMain.Buttons(3).Value = tbrPressed
Else
    tbarMain.Buttons(1).Value = tbrPressed
End If

If (X >= 0 And Y <= 2000) And Not (X >= 0 And Y <= 1000) Then
    tbarMain.Buttons(1).Value = tbrPressed
    tbarMain.Buttons(2).Value = tbrUnpressed
    tbarMain.Buttons(3).Value = tbrPressed
Else
    tbarMain.Buttons(2).Value = tbrPressed
End If

If (X >= 0 And Y <= 3000) And Not (X >= 0 And Y <= 2000) Then
    tbarMain.Buttons(1).Value = tbrPressed
    tbarMain.Buttons(2).Value = tbrPressed
    tbarMain.Buttons(3).Value = tbrUnpressed
Else
    tbarMain.Buttons(3).Value = tbrPressed
End If

End Sub

Private Sub Timer1_Timer()
StatusBar2.Panels(1).Text = Time()
End Sub

Private Function CheckOpenForms() As Boolean
Dim cntr As Integer
Dim formStat As Boolean, rptStat As Boolean

For cntr = 2 To Forms.Count - 1
    If Forms(cntr).Visible = True Then
        formStat = True
        Exit For
    End If
    formStat = False
Next
If envAmadeus.conReports.State = adStateOpen Then
    rptStat = True
Else
    rptStat = False
End If
CheckOpenForms = formStat Or rptStat
End Function

Private Function CheckUpdateMode() As Boolean
Dim mode As Boolean

On Error Resume Next
If frmUnitofMeasurement.UpdateMode = True Then
    mode = True
ElseIf frmSuppliers.UpdateMode = True Then
    mode = True
ElseIf frmStockDetails.UpdateMode = True Then
    mode = True
ElseIf frmRequest.UpdateMode = True Then
    mode = True
ElseIf frmPO.UpdateMode = True Then
    mode = True
ElseIf frmLeadman.UpdateMode = True Then
    mode = True
ElseIf frmGrower.UpdateMode = True Then
    mode = True
ElseIf frmClassification.UpdateMode = True Then
    mode = True
ElseIf frmUsage.UpdateMode = True Then
    mode = True
End If
CheckUpdateMode = mode
End Function

Private Sub tbarMain_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Select Case Button.Key
    Case "inventory"
        MDIForm1.MousePointer = vbHourglass
        frmInventory.Show
        frmInventory.SetFocus
        MDIForm1.MousePointer = vbDefault
    Case "stockdetails"
        MDIForm1.MousePointer = vbHourglass
        frmStockDetails.Show
        frmStockDetails.SetFocus
        MDIForm1.MousePointer = vbDefault
    Case "requestform"
        MDIForm1.MousePointer = vbHourglass
        frmRequest.Show
        frmRequest.SetFocus
        MDIForm1.MousePointer = vbDefault
End Select
End Sub
