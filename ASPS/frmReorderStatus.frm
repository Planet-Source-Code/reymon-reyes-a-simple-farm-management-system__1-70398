VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmReorderStatus 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reorder Status"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   5655
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
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
      Left            =   3000
      TabIndex        =   6
      ToolTipText     =   "Refresh the items"
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      DataField       =   "PODetail_ID"
      DataSource      =   "adoPODetails"
      Height          =   375
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "PODetail_ID"
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox Text1 
      DataField       =   "PO_ID"
      DataSource      =   "adoPO"
      Height          =   375
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "PO_ID"
      Top             =   5160
      Width           =   615
   End
   Begin MSAdodcLib.Adodc adoPODetails 
      Height          =   375
      Left            =   120
      Top             =   5520
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "adoPODetails"
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
   Begin MSAdodcLib.Adodc adoPO 
      Height          =   375
      Left            =   120
      Top             =   5160
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "adoPO"
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
   Begin VB.CommandButton cmdAddToPO 
      Caption         =   "Add To New Purchase Order"
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
      ToolTipText     =   "Add the items to a new purchase order"
      Top             =   4080
      Width           =   2775
   End
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
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      ToolTipText     =   "Close this form"
      Top             =   4080
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc adoReorderStatus 
      Height          =   375
      Left            =   120
      Top             =   4800
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "adoReorderStatus"
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
   Begin MSDataGridLib.DataGrid grdReorder 
      Bindings        =   "frmReorderStatus.frx":0000
      Height          =   3375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5953
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "Type"
         Caption         =   "Classification"
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
         DataField       =   "Category"
         Caption         =   "Kind"
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
      BeginProperty Column02 
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
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   2009.764
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2294.929
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   510.236
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   495
      Index           =   1
      Left            =   -240
      Shape           =   4  'Rounded Rectangle
      Top             =   4440
      Width           =   5295
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   1455
      Index           =   0
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   -1200
      Width           =   5295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The following stocks needs to be replenished."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   4770
   End
End
Attribute VB_Name = "frmReorderStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ReorderStat As Boolean

Private Sub adoPO_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub adoPODetails_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub adoReorderStatus_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub cmdAddToPO_Click()
Dim cntr As Integer

On Error GoTo ErrHandler

adoPO.Recordset.AddNew
adoPO.Recordset.Fields("Date") = Now()
adoPO.Recordset.Update
adoPO.Refresh
adoPO.Recordset.MoveLast
adoReorderStatus.Recordset.MoveFirst
Do Until adoReorderStatus.Recordset.EOF
    adoPODetails.Recordset.AddNew
    adoPODetails.Recordset.Fields("PO_ID") = adoPO.Recordset.Fields("PO_ID")
    adoPODetails.Recordset.Fields("StockType_ID") = adoReorderStatus.Recordset.Fields("StockType_ID")
    adoPODetails.Recordset.Fields("StockTypeCategory_ID") = adoReorderStatus.Recordset.Fields("StockTypeCategory_ID")
    adoPODetails.Recordset.Update
    adoReorderStatus.Recordset.MoveNext
Loop
frmPO.Show
frmPO.adoPO.Recordset.MoveLast
frmPO.qryPODetails (adoPO.Recordset.Fields("PO_ID"))
Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdAddToPO_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAddToPO.FontBold = True
cmdRefresh.FontBold = False
cmdClose.FontBold = False
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdClose.FontBold = True
cmdRefresh.FontBold = False
cmdAddToPO.FontBold = False
End Sub

Private Sub cmdRefresh_Click()
On Error GoTo ErrHandler

adoReorderStatus.Recordset.Requery
If adoReorderStatus.Recordset.RecordCount = 0 Then
    cmdAddToPO.Enabled = False
Else
    cmdAddToPO.Enabled = True
End If

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdRefresh_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdRefresh.FontBold = True
cmdClose.FontBold = False
cmdAddToPO.FontBold = False
End Sub

Private Sub Form_Activate()
Call cmdRefresh_Click
Me.grdReorder.Refresh
End Sub

Private Sub Form_GotFocus()
'MsgBox "Gotfocus"
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Left = 0
Me.Top = 0

Call ConnectDB(adoReorderStatus)
adoReorderStatus.CommandType = adCmdTable
adoReorderStatus.RecordSource = "qryReorderStatus"
adoReorderStatus.Refresh

Call ConnectDB(adoPO)
adoPO.CommandType = adCmdTable
adoPO.RecordSource = "tblPO"
adoPO.Refresh

Call ConnectDB(adoPODetails)
adoPODetails.CommandType = adCmdTable
adoPODetails.RecordSource = "tblPODetails"
adoPODetails.Refresh

If adoReorderStatus.Recordset.RecordCount = 0 Then
    cmdAddToPO.Enabled = False
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdClose.FontBold = False
cmdRefresh.FontBold = False
cmdAddToPO.FontBold = False
End Sub

Private Sub grdReorder_Error(ByVal DataError As Integer, Response As Integer)
Response = 0
End Sub
