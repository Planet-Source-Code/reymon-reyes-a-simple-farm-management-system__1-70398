VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmInventory 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventory"
   ClientHeight    =   8505
   ClientLeft      =   1875
   ClientTop       =   1980
   ClientWidth     =   13005
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   13005
   Begin VB.CommandButton cmdPreviewInventory 
      Caption         =   "Preview"
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
      Left            =   7680
      TabIndex        =   6
      Top             =   7920
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc adoInventorySummary 
      Height          =   375
      Left            =   2040
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Caption         =   "adoInventorySummary"
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
   Begin MSDataGridLib.DataGrid grdInventorySummary 
      Bindings        =   "inventory.frx":0000
      Height          =   6615
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   11668
      _Version        =   393216
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
         DataField       =   "SumOfQtyRemaining"
         Caption         =   "Stocks on Hand"
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
         RecordSelectors =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   1964.976
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1679.811
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdViewAllInventory 
      Caption         =   "&Refresh"
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
      Left            =   9000
      TabIndex        =   7
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
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
      Left            =   11640
      Picture         =   "inventory.frx":0022
      TabIndex        =   10
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton cmdStocks 
      Caption         =   "&Stocks"
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
      Left            =   10320
      TabIndex        =   8
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
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
      Left            =   10095
      TabIndex        =   3
      Top             =   560
      Width           =   975
   End
   Begin VB.TextBox txtSearch 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6960
      TabIndex        =   2
      Top             =   600
      Width           =   3015
   End
   Begin MSDataListLib.DataCombo dcboKind 
      Bindings        =   "inventory.frx":02DF
      Height          =   360
      Left            =   4575
      TabIndex        =   1
      Top             =   600
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      ListField       =   "Category"
      BoundColumn     =   ""
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
   Begin MSDataListLib.DataCombo dcboClassification 
      Bindings        =   "inventory.frx":02F5
      Height          =   360
      Left            =   2175
      TabIndex        =   0
      Top             =   600
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      ListField       =   "Type"
      BoundColumn     =   ""
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
   Begin MSDataGridLib.DataGrid grdInventory 
      Bindings        =   "inventory.frx":0315
      Height          =   6615
      Left            =   5640
      TabIndex        =   5
      Top             =   1200
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   11668
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   8454143
      HeadLines       =   1
      RowHeight       =   19
      RowDividerStyle =   1
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
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "Date"
         Caption         =   "Date"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "MM/dd/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Stock_ID"
         Caption         =   "Stock_ID"
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
      BeginProperty Column03 
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
      BeginProperty Column04 
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
      BeginProperty Column05 
         DataField       =   "Unit_Cost"
         Caption         =   "Unit_Cost"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "QtyRemaining"
         Caption         =   "CurrentQty"
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
      BeginProperty Column07 
         DataField       =   "Remarks"
         Caption         =   "Remarks"
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
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         RecordSelectors =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column01 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1769.953
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1980.284
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   3119.811
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoKind 
      Height          =   375
      Left            =   -120
      Top             =   360
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Left            =   -120
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "adoClassification"
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
   Begin MSAdodcLib.Adodc adoInventory 
      Height          =   375
      Left            =   -120
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "adoInventory"
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
   Begin VB.Shape Shape4 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   615
      Left            =   11160
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   2055
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C000&
      BorderWidth     =   25
      X1              =   2160
      X2              =   12960
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   7800
      Shape           =   4  'Rounded Rectangle
      Top             =   8280
      Width           =   5175
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1215
      Left            =   6000
      Shape           =   2  'Oval
      Top             =   7920
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   0
      Top             =   7920
      Width           =   7095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Text"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   6975
      TabIndex        =   12
      Top             =   360
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Kind"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4695
      TabIndex        =   11
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Classification"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2190
      TabIndex        =   9
      Top             =   360
      Width           =   1155
   End
End
Attribute VB_Name = "frmInventory"
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

Private Sub adoClassification_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub adoInventory_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub adoInventorySummary_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub adoKind_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdClose.FontBold = True
cmdPreviewInventory.FontBold = False
cmdViewAllInventory.FontBold = False
cmdStocks.FontBold = False
End Sub

Private Sub cmdPreviewInventory_Click()
Me.MousePointer = vbHourglass
rptInventorySummary.Show
rptInventorySummary.WindowState = vbMaximized
Me.MousePointer = vbDefault
End Sub

Private Sub cmdPreviewInventory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdPreviewInventory.FontBold = True
cmdViewAllInventory.FontBold = False
cmdStocks.FontBold = False
cmdClose.FontBold = False
End Sub

Private Sub cmdSearch_Click()
adoInventorySummary.CommandType = adCmdText
adoInventorySummary.RecordSource = "select * from qryInventorySummary where Category like '" & txtSearch.Text & "%'"
adoInventorySummary.Refresh
End Sub

Private Sub cmdSearch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdSearch.FontBold = True
End Sub

Private Sub cmdStocks_Click()
frmStockDetails.Show
frmStockDetails.SetFocus
End Sub

Private Sub cmdStocks_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdStocks.FontBold = True
cmdPreviewInventory.FontBold = False
cmdViewAllInventory.FontBold = False
cmdClose.FontBold = False
End Sub

Private Sub cmdViewAllInventory_Click()
On Error GoTo ErrHandler

Me.MousePointer = vbHourglass
adoInventory.CommandType = adCmdText
adoInventory.RecordSource = "select * from qryInventory where QtyRemaining<>0"
adoInventory.Refresh
adoInventorySummary.CommandType = adCmdTable
adoInventorySummary.RecordSource = "qryInventorySummary"
adoInventorySummary.Refresh
Me.MousePointer = vbDefault

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdViewAllInventory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdViewAllInventory.FontBold = True
cmdPreviewInventory.FontBold = False
cmdStocks.FontBold = False
cmdClose.FontBold = False
End Sub

Private Sub dcboClassification_Change()
On Error Resume Next

If loadStatus = False Then Exit Sub
adoClassification.Recordset.Bookmark = dcboClassification.SelectedItem
adoKind.CommandType = adCmdText
adoKind.RecordSource = "select * from tblstocktypecategory where StockType_ID=" & adoClassification.Recordset.Fields("StockType_ID")
adoKind.Refresh
Set dcboKind.RowSource = adoKind.Recordset
dcboKind.ListField = "Category"
dcboKind.Text = ""
dcboKind.ReFill
adoInventorySummary.CommandType = adCmdText
adoInventorySummary.RecordSource = "select * from qryInventorySummary where StockType_ID =" & adoClassification.Recordset.Fields("StockType_ID")
adoInventorySummary.Refresh
End Sub

Private Sub dcboClassification_Click(Area As Integer)
loadStatus = True
End Sub

Private Sub dcboKind_Change()
On Error Resume Next

If loadStatus = False Then Exit Sub
If dcboKind.Text = "" Then Exit Sub
'adoKind.Refresh
'MsgBox adoKind.Recordset.Fields("StockTypeCategory_ID").Value
'MsgBox dcboKind.Text
adoKind.Recordset.Bookmark = dcboKind.SelectedItem
adoInventorySummary.CommandType = adCmdText
adoInventorySummary.RecordSource = "select * from qryInventorySummary where StockType_ID =" & adoClassification.Recordset.Fields("StockType_ID") _
    & " and StockTypeCategory_ID=" & adoKind.Recordset.Fields("StockTypeCategory_ID")
adoInventorySummary.Refresh
End Sub

Private Sub dcboKind_Click(Area As Integer)
'loadStatus = True
End Sub

Private Sub dcboKind_GotFocus()
loadStatus = True
End Sub

Private Sub dcboKind_LostFocus()
loadStatus = False
End Sub

Private Sub Form_Activate()
Call CenterForm(frmInventory, MDIForm1)
Me.grdInventory.Refresh
grdInventorySummary.Refresh
'Me.grdInventory.SetFocus
End Sub

Private Sub Form_Load()
Call CenterForm(frmInventory, MDIForm1)
loadStatus = False
'setup the adodcs
Call ConnectDB(adoInventorySummary)
adoInventorySummary.CommandType = adCmdTable
adoInventorySummary.RecordSource = "qryInventorySummary"
adoInventorySummary.Refresh
Call ConnectDB(adoInventory)
adoInventory.CommandType = adCmdText
adoInventory.RecordSource = "select * from qryInventory where QtyRemaining<>0"
adoInventory.Refresh
Call ConnectDB(adoClassification)
adoClassification.CommandType = adCmdTable
adoClassification.RecordSource = "tblstocktype"
adoClassification.Refresh
Call ConnectDB(adoKind)
adoKind.CommandType = adCmdTable
adoKind.RecordSource = "tblstocktypecategory"
adoKind.Refresh
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdPreviewInventory.FontBold = False
cmdViewAllInventory.FontBold = False
cmdStocks.FontBold = False
cmdClose.FontBold = False
cmdSearch.FontBold = False
End Sub

Private Sub grdInventory_Error(ByVal DataError As Integer, Response As Integer)
Response = 0
End Sub

Private Sub grdInventorySummary_Error(ByVal DataError As Integer, Response As Integer)
Response = 0
End Sub

Private Sub grdInventorySummary_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo ErrHandler

adoInventory.CommandType = adCmdText
adoInventory.RecordSource = "select * from qryInventory where QtyRemaining<>0 and StockType_ID=" & _
    adoInventorySummary.Recordset.Fields("StockType_ID") & " and StockTypeCategory_ID= " & _
    adoInventorySummary.Recordset.Fields("StockTypeCategory_ID")
adoInventory.Refresh

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call cmdSearch_Click
End If
End Sub
