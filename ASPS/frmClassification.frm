VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmClassification 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Classification"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   6240
   Begin VB.TextBox txtKindID 
      DataField       =   "StockTypeCategory_ID"
      DataSource      =   "adoTblKind"
      Height          =   285
      Left            =   6360
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   4080
      Width           =   735
   End
   Begin MSAdodcLib.Adodc adoTblKind 
      Height          =   375
      Left            =   4800
      Top             =   4440
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "adoTblKind"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoUnit 
      Height          =   375
      Left            =   4800
      Top             =   4800
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "adoUnit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">|"
      Height          =   375
      Left            =   3360
      TabIndex        =   17
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">>"
      Height          =   375
      Left            =   2280
      TabIndex        =   16
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<<"
      Height          =   375
      Left            =   1200
      TabIndex        =   15
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "|<"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Frame fraKinds 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Kinds"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   120
      TabIndex        =   19
      Top             =   960
      Width           =   4455
      Begin MSDataListLib.DataCombo dcboUnit 
         Bindings        =   "frmClassification.frx":0000
         DataField       =   "UnitType_ID"
         DataSource      =   "adoKind"
         Height          =   360
         Left            =   2520
         TabIndex        =   9
         Top             =   3240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   635
         _Version        =   393216
         Locked          =   -1  'True
         Style           =   2
         ListField       =   "Unit"
         BoundColumn     =   "UnitType_ID"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton cmdCancelKind 
         Caption         =   "Cancel"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   13
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton cmdDeleteKind 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   12
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton cmdEditKind 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   11
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton cmdAddKind 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   3720
         Width           =   1095
      End
      Begin VB.TextBox txtKind 
         DataField       =   "Category"
         DataSource      =   "adoKind"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   3240
         Width           =   2295
      End
      Begin MSDataGridLib.DataGrid grdClassification 
         Bindings        =   "frmClassification.frx":0016
         Height          =   2895
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   5106
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   8454143
         HeadLines       =   1
         RowHeight       =   19
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "Category"
            Caption         =   "Category"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Unit"
            Caption         =   "Unit"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            Locked          =   -1  'True
            BeginProperty Column00 
               ColumnWidth     =   2429.858
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   720
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraMainCon 
      BackColor       =   &H0000C000&
      Height          =   3375
      Left            =   4800
      TabIndex        =   18
      Top             =   960
      Width           =   1455
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelClassification 
         Caption         =   "Cancel"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton cmdDeleteClassification 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdEditClassification 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdNewClassification 
         Caption         =   "&New"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc adoKind 
      Height          =   375
      Left            =   4800
      Top             =   5520
      Visible         =   0   'False
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "adoKind"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoClassification 
      Height          =   375
      Left            =   4800
      Top             =   5160
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "adoclassification"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtClassification 
      DataField       =   "Type"
      DataSource      =   "adoClassification"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   615
      Left            =   4800
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblclassification 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Classification:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      Height          =   615
      Left            =   2520
      Shape           =   4  'Rounded Rectangle
      Top             =   -240
      Width           =   4095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   375
      Left            =   720
      Shape           =   3  'Circle
      Top             =   -240
      Width           =   495
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   375
      Left            =   960
      Top             =   -240
      Width           =   2175
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   735
      Left            =   480
      Shape           =   2  'Oval
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   375
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   5760
      Width           =   5535
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      Height          =   495
      Index           =   1
      Left            =   -120
      Shape           =   4  'Rounded Rectangle
      Top             =   5640
      Width           =   1335
   End
End
Attribute VB_Name = "frmClassification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Note: Please understand the forms before trying to understand the codes
'      so that it's much easier for us to trace the flow of the forms.
'      Also consider that there are 1 and a million ways to do something
'      so the codes here might be unorthodox but the IMPORTANT thing is that
'      the desired output is achieved.
'      The approach taken here is to use adodc instead of pure ado code.
'      Don't forget to place comments in your code so that other people who sees your
'      code doesn't suffer from painfully trying to understand it.
'      Ayaw liboga imo sarili kay daghan na ta! Sayuna lang para dili maglisod!
' Peace! :D
' the_Etc Development Team
' Mabuhay ang Pilipinas!

Option Explicit
Dim loadStatus As Boolean
Public UpdateMode As Boolean

Private Sub adoClassification_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub adoKind_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub adoTblKind_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub adoUnit_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub cmdAddKind_Click()
On Error GoTo ErrHandler

If cmdAddKind.Caption = "&Add" Then
    UpdateMode = True
    cmdAddKind.Caption = "&Save"
    adoKind.Recordset.AddNew
    cmdEditKind.Enabled = False
    cmdDeleteKind.Enabled = False
    cmdCancelKind.Enabled = True
    cmdFirst.Enabled = False
    cmdPrevious.Enabled = False
    cmdNext.Enabled = False
    cmdLast.Enabled = False
    fraMainCon.Enabled = False
    adoKind.Recordset.Fields("StockType_ID") = adoClassification.Recordset.Fields("StockType_ID")
    dcboUnit.Locked = False
    txtKind.Locked = False
    txtKind.SetFocus
ElseIf cmdAddKind.Caption = "&Save" Then
    If txtKind.Text = "" Then
        MsgBox "Kind is required."
        txtKind.SetFocus
        Exit Sub
    ElseIf dcboUnit.Text = "" Then
        MsgBox "Please select a Unit."
        dcboUnit.SetFocus
        Exit Sub
    Else
        Me.MousePointer = vbHourglass
        cmdAddKind.Caption = "&Add"
        adoKind.Recordset.Update
        cmdEditKind.Enabled = True
        cmdDeleteKind.Enabled = True
        cmdCancelKind.Enabled = False
        dcboUnit.Locked = True
        txtKind.Locked = True
        cmdFirst.Enabled = True
        cmdPrevious.Enabled = True
        cmdNext.Enabled = True
        cmdLast.Enabled = True
        fraMainCon.Enabled = True
        adoTblKind.Refresh
        adoKind.Refresh
        Call qryKinds(adoClassification.Recordset.Fields("StockType_ID"))
        Me.MousePointer = vbDefault
        MsgBox "Record successfully updated."
        UpdateMode = False
    End If
End If

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdCancelClassification_Click()
On Error GoTo ErrHandler

UpdateMode = False
adoClassification.Recordset.Cancel
adoClassification.Refresh
txtClassification.Locked = True
cmdNewClassification.Caption = "&New"
cmdEditClassification.Enabled = True
cmdDeleteClassification.Enabled = True
fraKinds.Enabled = True
cmdFirst.Enabled = True
cmdPrevious.Enabled = True
cmdNext.Enabled = True
cmdLast.Enabled = True
cmdClose.Enabled = True
Call qryKinds(adoClassification.Recordset.Fields("StockType_ID"))

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdCancelClassification_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdNewClassification.FontBold = False
cmdEditClassification.FontBold = False
cmdDeleteClassification.FontBold = False
cmdCancelClassification.FontBold = True
cmdClose.FontBold = False
End Sub

Private Sub cmdCancelKind_Click()
On Error GoTo ErrHandler

UpdateMode = False
adoKind.Recordset.Cancel
adoKind.Refresh
cmdAddKind.Caption = "&Add"
cmdDeleteKind.Enabled = True
cmdEditKind.Enabled = True
cmdCancelKind.Enabled = False
cmdFirst.Enabled = True
cmdPrevious.Enabled = True
cmdNext.Enabled = True
cmdLast.Enabled = True
fraMainCon.Enabled = True
dcboUnit.Locked = True
txtKind.Locked = True

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdNewClassification.FontBold = False
cmdEditClassification.FontBold = False
cmdDeleteClassification.FontBold = False
cmdCancelClassification.FontBold = False
cmdClose.FontBold = True
End Sub

Private Sub cmdDeleteClassification_Click()
On Error GoTo errorHandler

If MsgBox("The selected RECORD and its corresponding stocks will be deleted, are you sure you want to continue?" _
    & vbCr & "This action is irreversible.", vbOKCancel + vbExclamation, "Delete") = vbOK Then
    Me.MousePointer = vbHourglass
    adoClassification.Recordset.Delete
    adoClassification.Recordset.Requery
    Call qryKinds(adoClassification.Recordset.Fields("StockType_ID"))
    Me.MousePointer = vbDefault
End If

Exit Sub
errorHandler:
    If Err.Number = -2147467259 Then
        MsgBox "You cannot delete this record because it is being used in" & vbCr & _
            "one or more existing records."
        Me.MousePointer = vbDefault
        adoClassification.Refresh
        Call qryKinds(adoClassification.Recordset.Fields("StockType_ID"))
        Me.MousePointer = vbDefault
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub cmdDeleteClassification_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdNewClassification.FontBold = False
cmdEditClassification.FontBold = False
cmdDeleteClassification.FontBold = True
cmdCancelClassification.FontBold = False
cmdClose.FontBold = False
End Sub

Private Sub cmdDeleteKind_Click()
On Error GoTo errorHandler

If MsgBox("The selected RECORD will be deleted, are you sure you want to continue?" & vbCr & "This action is irreversible.", vbOKCancel + vbExclamation, "Delete") = vbOK Then
    adoTblKind.Recordset.MoveFirst
    adoTblKind.Recordset.Find "StockTypeCategory_ID=" & adoKind.Recordset.Fields("StockTypeCategory_ID")
    adoTblKind.Recordset.Delete
    adoTblKind.Refresh
    adoKind.Refresh
End If

Exit Sub
errorHandler:
    If Err.Number = -2147467259 Then
        MsgBox "You cannot delete this record because it is being used in" & vbCr & _
            "one or more existing records."
        adoKind.Refresh
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub cmdEditClassification_Click()
On Error GoTo ErrHandler

UpdateMode = True
txtClassification.Locked = False
txtClassification.SetFocus
fraKinds.Enabled = False
cmdNewClassification.Caption = "&Save"
cmdDeleteClassification.Enabled = False
cmdEditClassification.Enabled = False
cmdCancelClassification.Enabled = True
cmdFirst.Enabled = False
cmdPrevious.Enabled = False
cmdNext.Enabled = False
cmdLast.Enabled = False
cmdClose.Enabled = False

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdEditClassification_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdNewClassification.FontBold = False
cmdEditClassification.FontBold = True
cmdDeleteClassification.FontBold = False
cmdCancelClassification.FontBold = False
cmdClose.FontBold = False
End Sub

Private Sub cmdEditKind_Click()
On Error GoTo ErrHandler

UpdateMode = True
cmdAddKind.Caption = "&Save"
cmdEditKind.Enabled = False
cmdDeleteKind.Enabled = False
cmdCancelKind.Enabled = True
cmdFirst.Enabled = False
cmdPrevious.Enabled = False
cmdNext.Enabled = False
cmdLast.Enabled = False
fraMainCon.Enabled = False
dcboUnit.Locked = False
txtKind.Locked = False
txtKind.SetFocus

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdFirst_Click()
adoClassification.Recordset.MoveFirst
cmdNext.Enabled = True
Call qryKinds(adoClassification.Recordset.Fields("StockType_ID"))
End Sub

Private Sub cmdLast_Click()
adoClassification.Recordset.MoveLast
cmdPrevious.Enabled = True
Call qryKinds(adoClassification.Recordset.Fields("StockType_ID"))
End Sub

Private Sub cmdNewClassification_Click()
On Error GoTo ErrHandler

If cmdNewClassification.Caption = "&New" Then
    UpdateMode = True
    cmdNewClassification.Caption = "&Save"
    cmdEditClassification.Enabled = False
    cmdDeleteClassification.Enabled = False
    cmdCancelClassification.Enabled = True
    fraKinds.Enabled = False
    cmdFirst.Enabled = False
    cmdPrevious.Enabled = False
    cmdNext.Enabled = False
    cmdLast.Enabled = False
    cmdClose.Enabled = False
    
    adoClassification.Recordset.AddNew
    Call qryKinds(0)

    txtClassification.Locked = False
    txtClassification.SetFocus

ElseIf cmdNewClassification.Caption = "&Save" Then
    If txtClassification.Text = "" Then
        MsgBox "Classification is required."
        txtClassification.SetFocus
        Exit Sub
    Else
        Me.MousePointer = vbHourglass
        cmdNewClassification.Caption = "&New"
        adoClassification.Recordset.Update
        cmdDeleteClassification.Enabled = True
        cmdEditClassification.Enabled = True
        cmdCancelClassification.Enabled = False
        txtClassification.Locked = True
        fraKinds.Enabled = True
        cmdFirst.Enabled = True
        cmdPrevious.Enabled = True
        cmdNext.Enabled = True
        cmdLast.Enabled = True
        cmdClose.Enabled = True
        adoClassification.Recordset.Resync
        'adoBatch.Refresh
        'adoBatch.Recordset.MoveLast
        Me.MousePointer = vbDefault
        MsgBox "Record successfully updated."
        UpdateMode = False
    End If
End If

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdNewClassification_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdNewClassification.FontBold = True
cmdEditClassification.FontBold = False
cmdDeleteClassification.FontBold = False
cmdCancelClassification.FontBold = False
cmdClose.FontBold = False
End Sub

Private Sub cmdNext_Click()
cmdPrevious.Enabled = True
adoClassification.Recordset.MoveNext
If adoClassification.Recordset.EOF Then
    cmdNext.Enabled = False
    adoClassification.Recordset.MoveLast
    Call qryKinds(adoClassification.Recordset.Fields("StockType_ID"))
End If
Call qryKinds(adoClassification.Recordset.Fields("StockType_ID"))
End Sub

Private Sub cmdPrevious_Click()
cmdNext.Enabled = True
adoClassification.Recordset.MovePrevious
If adoClassification.Recordset.BOF Then
    cmdPrevious.Enabled = False
    adoClassification.Recordset.MoveFirst
    Call qryKinds(adoClassification.Recordset.Fields("StockType_ID"))
End If
Call qryKinds(adoClassification.Recordset.Fields("StockType_ID"))
End Sub

Private Sub Form_Activate()
If UpdateMode = True Then Exit Sub
adoUnit.Recordset.Requery
dcboUnit.ListField = "Unit"
dcboUnit.ReFill
End Sub

Private Sub Form_Deactivate()
If Me.UpdateMode = True Then
    Me.ZOrder (vbBringToFront)
    Me.SetFocus
    Exit Sub
End If
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler

Call CenterForm(frmClassification, MDIForm1)
loadStatus = False
'connect and set the connectionstring, commandtype, recordset properties of the adodcs
Call ConnectDB(adoClassification)
adoClassification.CommandType = adCmdTable
adoClassification.RecordSource = "tblstocktype"
adoClassification.Refresh
Call ConnectDB(adoKind)
adoKind.CommandType = adCmdTable
adoKind.RecordSource = "qryKinds"
adoKind.Refresh
Call ConnectDB(adoUnit)
adoUnit.CommandType = adCmdTable
adoUnit.RecordSource = "tblunittype"
adoUnit.Refresh
Call ConnectDB(adoTblKind)
adoTblKind.CommandType = adCmdTable
adoTblKind.RecordSource = "tblstocktypecategory"
adoTblKind.Refresh
Call qryKinds(adoClassification.Recordset.Fields("StockType_ID"))

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Public Sub qryKinds(criteria As Long)
On Error GoTo ErrHandler
adoKind.CommandType = adCmdText
adoKind.RecordSource = "select * from qryKinds where StockType_ID=" & criteria
adoKind.Refresh

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub grdClassification_Error(ByVal DataError As Integer, Response As Integer)
Response = 0
End Sub
