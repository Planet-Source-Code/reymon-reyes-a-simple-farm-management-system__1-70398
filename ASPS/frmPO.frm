VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPO 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Order"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7905
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   7905
   Begin MSMask.MaskEdBox mskDate 
      DataField       =   "Date"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "MM/dd/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      DataSource      =   "adoPO"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "MM/dd/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C000&
      Height          =   3855
      Left            =   6480
      TabIndex        =   25
      Top             =   1320
      Width           =   1455
      Begin VB.CommandButton cmdNewPO 
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
         ToolTipText     =   "Create new PO"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdEditPO 
         Caption         =   "Edit"
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
         ToolTipText     =   "Modify current PO"
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdDeletePO 
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
         ToolTipText     =   "Delete PO"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelPO 
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
         ToolTipText     =   "Cancel current operation"
         Top             =   2040
         Width           =   1215
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
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Close this form"
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CommandButton cmdPreviewPO 
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
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Print preview of the PO"
         Top             =   2640
         Width           =   1215
      End
   End
   Begin VB.TextBox txtPODetailID 
      DataField       =   "PODetail_ID"
      DataSource      =   "adotblPODetails"
      Height          =   375
      Left            =   9960
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   2040
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSAdodcLib.Adodc adotblPODetails 
      Height          =   375
      Left            =   7920
      Top             =   2400
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
      Caption         =   "adotblPODetails"
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Go to the last record"
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Go to the next record"
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   18
      ToolTipText     =   "Go to the previous record"
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "|<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   17
      ToolTipText     =   "Go to the first record"
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Frame fraPODetails 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Particulars"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      TabIndex        =   23
      Top             =   1440
      Width           =   6135
      Begin VB.CommandButton cmdCancelPODetail 
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
         Left            =   4320
         TabIndex        =   16
         ToolTipText     =   "Cancel current operation"
         Top             =   3600
         Width           =   1215
      End
      Begin VB.CommandButton cmdDeletePODetail 
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
         Left            =   3120
         TabIndex        =   15
         ToolTipText     =   "Delete the selected item"
         Top             =   3600
         Width           =   1215
      End
      Begin VB.CommandButton cmdEditPODetail 
         Caption         =   "Edit"
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
         Left            =   1920
         TabIndex        =   14
         ToolTipText     =   "Edit the item"
         Top             =   3600
         Width           =   1215
      End
      Begin VB.TextBox txtQuantity 
         Alignment       =   1  'Right Justify
         DataField       =   "Quantity"
         DataSource      =   "adoPODetails"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4440
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   12
         Top             =   3120
         Width           =   1575
      End
      Begin VB.CommandButton cmdAddPODetail 
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
         Left            =   720
         TabIndex        =   13
         ToolTipText     =   "Add an item to PO"
         Top             =   3600
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo dcboKind 
         Bindings        =   "frmPO.frx":0000
         DataField       =   "StockTypeCategory_ID"
         DataSource      =   "adoPODetails"
         Height          =   360
         Left            =   2400
         TabIndex        =   11
         Top             =   3120
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   635
         _Version        =   393216
         Locked          =   -1  'True
         Style           =   2
         ListField       =   "Category"
         BoundColumn     =   "StockTypeCategory_ID"
         Text            =   ""
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
      Begin MSDataListLib.DataCombo dcboClassification 
         Bindings        =   "frmPO.frx":0016
         DataField       =   "StockType_ID"
         DataSource      =   "adoPODetails"
         Height          =   360
         Left            =   120
         TabIndex        =   10
         Top             =   3120
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   635
         _Version        =   393216
         Locked          =   -1  'True
         Style           =   2
         ListField       =   "Type"
         BoundColumn     =   "StockType_ID"
         Text            =   ""
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
      Begin MSDataGridLib.DataGrid grdPODetails 
         Bindings        =   "frmPO.frx":0036
         Height          =   2655
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   4683
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "Type"
            Caption         =   "Type"
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
         BeginProperty Column03 
            DataField       =   "Quantity"
            Caption         =   "Quantity"
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
               ColumnWidth     =   2174.74
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1950.236
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   945.071
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc adoKind 
      Height          =   375
      Left            =   7920
      Top             =   3120
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
      Left            =   7920
      Top             =   2760
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
   Begin MSAdodcLib.Adodc adoPODetails 
      Height          =   375
      Left            =   7920
      Top             =   2040
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
   Begin VB.TextBox txtRemarks 
      DataField       =   "Remarks"
      DataSource      =   "adoPO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   840
      Width           =   2895
   End
   Begin VB.TextBox txtDate 
      DataField       =   "Date"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "MM/dd/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      DataSource      =   "adoPO"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc adoPO 
      Height          =   375
      Left            =   7920
      Top             =   1680
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
   Begin VB.Shape Shape3 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   375
      Left            =   1440
      Top             =   -130
      Width           =   4335
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000C000&
      BorderWidth     =   30
      X1              =   0
      X2              =   6360
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   1335
      Left            =   6480
      Top             =   0
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   495
      Left            =   5640
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C000&
      BorderWidth     =   30
      X1              =   1320
      X2              =   1440
      Y1              =   20
      Y2              =   20
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks:"
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
      Left            =   90
      TabIndex        =   22
      Top             =   960
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
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
      Left            =   480
      TabIndex        =   21
      Top             =   480
      Width           =   495
   End
End
Attribute VB_Name = "frmPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public UpdateMode As Boolean
Dim loadStatus As Boolean

Private Sub adoClassification_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub adoKind_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub adoPO_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub adoPODetails_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub adotblPODetails_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub cmdAddPODetail_Click()
On Error GoTo ErrHandler

If cmdAddPODetail.Caption = "&Add" Then
    UpdateMode = True
    cmdAddPODetail.Caption = "&Save"
    adoPODetails.Recordset.AddNew
    cmdEditPODetail.Enabled = False
    cmdDeletePODetail.Enabled = False
    cmdCancelPODetail.Enabled = True
    adoPODetails.Recordset.Fields("PO_ID") = adoPO.Recordset.Fields("PO_ID")
    dcboClassification.Locked = False
    dcboClassification.Text = ""
    dcboKind.Locked = False
    dcboKind.Text = ""
    txtQuantity.Locked = False
    dcboClassification.SetFocus
ElseIf cmdAddPODetail.Caption = "&Save" Then
    If dcboClassification.Text = "" Then
        MsgBox "Please select a classification."
        dcboClassification.SetFocus
        Exit Sub
    ElseIf dcboKind.Text = "" Then
        MsgBox "Please select a kind."
        dcboKind.SetFocus
        Exit Sub
    ElseIf txtQuantity.Text = "" Or Not IsNumeric(Val(txtQuantity.Text)) _
            Or Val(txtQuantity.Text) = 0 Then
        MsgBox "Invalid quantity."
        txtQuantity.SetFocus
        Exit Sub
    Else
        Me.MousePointer = vbHourglass
        cmdAddPODetail.Caption = "&Add"
        adoPODetails.Recordset.Update
        cmdEditPODetail.Enabled = True
        cmdDeletePODetail.Enabled = True
        cmdCancelPODetail.Enabled = False
        dcboClassification.Locked = True
        dcboKind.Locked = True
        txtQuantity.Locked = True
        'adoPODetails.Recordset.Resync
        Call qryPODetails(adoPO.Recordset.Fields("PO_ID"))
        adoKind.CommandType = adCmdTable
        adoKind.RecordSource = "tblstocktypecategory"
        adoKind.Refresh
        Set dcboKind.RowSource = adoKind.Recordset
        dcboKind.ListField = "Category"
        dcboKind.ReFill
        dcboKind.Refresh
        adoPODetails.Refresh
        Me.MousePointer = vbDefault
        MsgBox "Record successfully updated."
        UpdateMode = False
    End If
End If

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdCancelPO_Click()
On Error GoTo ErrHandler

adoPO.Recordset.Cancel
adoPO.Refresh
cmdNewPO.Caption = "&New"
cmdEditPO.Enabled = True
cmdDeletePO.Enabled = True
cmdCancelPO.Enabled = False
txtDate.Locked = True
txtRemarks.Locked = True
mskDate.Enabled = False
fraPODetails.Enabled = True
cmdFirst.Enabled = True
cmdPrevious.Enabled = True
cmdNext.Enabled = True
cmdLast.Enabled = True
cmdPreviewPO.Enabled = True
cmdClose.Enabled = True
Call qryPODetails(adoPO.Recordset.Fields("PO_ID"))
UpdateMode = False

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdCancelPO_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdNewPO.FontBold = False
cmdEditPO.FontBold = False
cmdDeletePO.FontBold = False
cmdCancelPO.FontBold = True
cmdPreviewPO.FontBold = False
cmdClose.FontBold = False
End Sub

Private Sub cmdCancelPODetail_Click()
On Error GoTo ErrHandler

adoPODetails.Recordset.Cancel
adoPODetails.Refresh
adoKind.CommandType = adCmdTable
adoKind.RecordSource = "tblstocktypecategory"
adoKind.Refresh
Set dcboKind.RowSource = adoKind.Recordset
dcboKind.ListField = "Category"
dcboKind.Refresh
cmdAddPODetail.Caption = "&Add"
cmdEditPODetail.Enabled = True
cmdDeletePODetail.Enabled = True
cmdCancelPODetail.Enabled = False
dcboClassification.Locked = True
dcboKind.Locked = True
txtQuantity.Locked = True
UpdateMode = False

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdNewPO.FontBold = False
cmdEditPO.FontBold = False
cmdDeletePO.FontBold = False
cmdCancelPO.FontBold = False
cmdPreviewPO.FontBold = False
cmdClose.FontBold = True
End Sub

Private Sub cmdDeletePO_Click()
On Error GoTo ErrHandler

If MsgBox("The selected RECORD will be deleted, are you sure you want to continue?" & vbCr & "This action is irreversible.", vbOKCancel + vbExclamation, "Delete") = vbOK Then
    Me.MousePointer = vbHourglass
    adoPO.Recordset.Delete
    adoPO.Refresh
    Call qryPODetails(adoPO.Recordset.Fields("PO_ID"))
    Me.MousePointer = vbDefault
End If

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdDeletePO_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdNewPO.FontBold = False
cmdEditPO.FontBold = False
cmdDeletePO.FontBold = True
cmdCancelPO.FontBold = False
cmdPreviewPO.FontBold = False
cmdClose.FontBold = False
End Sub

Private Sub cmdDeletePODetail_Click()
On Error GoTo ErrHandler

If MsgBox("The selected RECORD will be deleted, are you sure you want to continue?" & vbCr & "This action is irreversible.", vbOKCancel + vbQuestion, "Delete") = vbOK Then
    Me.MousePointer = vbHourglass
    adotblPODetails.Refresh
    adotblPODetails.Recordset.Find "PODetail_ID=" & adoPODetails.Recordset.Fields("PODetail_ID"), , adSearchForward, 0
    adotblPODetails.Recordset.Delete
    adotblPODetails.Refresh
    adoPODetails.Refresh
    Me.MousePointer = vbDefault
End If

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdEditPO_Click()
On Error GoTo ErrHandler

UpdateMode = True
cmdNewPO.Caption = "&Save"
cmdEditPO.Enabled = False
cmdDeletePO.Enabled = False
cmdCancelPO.Enabled = True
'txtDate.Locked = False
txtRemarks.Locked = False
mskDate.Enabled = True
fraPODetails.Enabled = False
cmdFirst.Enabled = False
cmdPrevious.Enabled = False
cmdNext.Enabled = False
cmdLast.Enabled = False
cmdPreviewPO.Enabled = False
cmdClose.Enabled = False

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdEditPO_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdNewPO.FontBold = False
cmdEditPO.FontBold = True
cmdDeletePO.FontBold = False
cmdCancelPO.FontBold = False
cmdPreviewPO.FontBold = False
cmdClose.FontBold = False
End Sub

Private Sub cmdEditPODetail_Click()
On Error GoTo ErrHandler

UpdateMode = True
cmdAddPODetail.Caption = "&Save"
cmdEditPODetail.Enabled = False
cmdDeletePODetail.Enabled = False
cmdCancelPODetail.Enabled = True
dcboClassification.Locked = False
dcboKind.Locked = False
txtQuantity.Locked = False

adoKind.CommandType = adCmdText
adoKind.RecordSource = "select * from tblstocktypecategory where StockType_ID=" & adoClassification.Recordset.Fields("StockType_ID")
adoKind.Refresh
Set dcboKind.RowSource = adoKind.Recordset
dcboKind.ListField = "Category"
dcboKind.ReFill
dcboKind.Refresh
dcboClassification.SetFocus

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdFirst_Click()
adoPO.Recordset.MoveFirst
Call qryPODetails(adoPO.Recordset.Fields("PO_ID"))
End Sub

Private Sub cmdLast_Click()
adoPO.Recordset.MoveLast
Call qryPODetails(adoPO.Recordset.Fields("PO_ID"))
End Sub

Private Sub cmdNewPO_Click()
On Error GoTo ErrHandler

If cmdNewPO.Caption = "&New" Then
    UpdateMode = True
    cmdNewPO.Caption = "&Save"
    cmdEditPO.Enabled = False
    cmdDeletePO.Enabled = False
    cmdCancelPO.Enabled = True
    fraPODetails.Enabled = False
    cmdFirst.Enabled = False
    cmdPrevious.Enabled = False
    cmdNext.Enabled = False
    cmdLast.Enabled = False
    cmdPreviewPO.Enabled = False
    cmdClose.Enabled = False
    
    adoPO.Recordset.AddNew
    'Call qryKinds(0)
    Call qryPODetails(0)
    'txtDate.Text = Date
    'txtDate.Locked = False
    mskDate.Text = Format(Date, "MM/dd/yyyy")
    mskDate.Enabled = True
    txtRemarks.Locked = False

ElseIf cmdNewPO.Caption = "&Save" Then
    If mskDate.Text = "" Or Not IsDate(mskDate.Text) Then
        MsgBox "Date is required."
        mskDate.SetFocus
        Exit Sub
    Else
        Me.MousePointer = vbHourglass
        cmdNewPO.Caption = "&New"
        adoPO.Recordset.Update
        
        cmdEditPO.Enabled = True
        cmdDeletePO.Enabled = True
        cmdCancelPO.Enabled = False
        'txtDate.Locked = True
        'txtRemarks.Locked = True
        mskDate.Enabled = False
        adoPO.Recordset.Resync
        adoPO.Refresh
        adoPO.Recordset.MoveLast
        fraPODetails.Enabled = True
        cmdFirst.Enabled = True
        cmdPrevious.Enabled = True
        cmdNext.Enabled = True
        cmdLast.Enabled = True
        cmdPreviewPO.Enabled = True
        cmdClose.Enabled = True
        Me.MousePointer = vbDefault
        MsgBox "Record successfully updated."
        UpdateMode = False
    End If
End If

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdNewPO_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdNewPO.FontBold = True
cmdEditPO.FontBold = False
cmdDeletePO.FontBold = False
cmdCancelPO.FontBold = False
cmdPreviewPO.FontBold = False
cmdClose.FontBold = False
End Sub

Private Sub cmdNext_Click()
cmdPrevious.Enabled = True
adoPO.Recordset.MoveNext
If adoPO.Recordset.EOF Then
    cmdNext.Enabled = False
    adoPO.Recordset.MoveLast
    Call qryPODetails(adoPO.Recordset.Fields("PO_ID"))
End If
Call qryPODetails(adoPO.Recordset.Fields("PO_ID"))
End Sub

Private Sub cmdPreviewPO_Click()
On Error GoTo ErrHandler

Me.MousePointer = vbHourglass
If envAmadeus.rscmdRptPO_Grouping.State = adStateOpen Then
    envAmadeus.rscmdRptPO_Grouping.Close
End If
Unload rptPO
envAmadeus.cmdRptPO_Grouping adoPO.Recordset.Fields("PO_ID")
rptPO.Show
rptPO.WindowState = vbMaximized
Me.MousePointer = vbDefault

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdPreviewPO_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdNewPO.FontBold = False
cmdEditPO.FontBold = False
cmdDeletePO.FontBold = False
cmdCancelPO.FontBold = False
cmdPreviewPO.FontBold = True
cmdClose.FontBold = False
End Sub

Private Sub cmdPrevious_Click()
cmdNext.Enabled = True
adoPO.Recordset.MovePrevious
If adoPO.Recordset.BOF Then
    cmdPrevious.Enabled = False
    adoPO.Recordset.MoveFirst
    Call qryPODetails(adoPO.Recordset.Fields("PO_ID"))
End If
Call qryPODetails(adoPO.Recordset.Fields("PO_ID"))
End Sub

Private Sub dcboClassification_Change()
On Error GoTo ErrHandler

If loadStatus = False Then Exit Sub

adoClassification.Recordset.MoveFirst
adoClassification.Recordset.Find "Type='" & dcboClassification.Text & "'", , adSearchForward, 0

adoKind.CommandType = adCmdText
adoKind.RecordSource = "select * from tblstocktypecategory where StockType_ID=" & adoClassification.Recordset.Fields("StockType_ID")
adoKind.Refresh

Set dcboKind.RowSource = adoKind.Recordset
dcboKind.ListField = "Category"
dcboKind.ReFill
dcboKind.Refresh
dcboKind.Text = ""

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub dcboClassification_Click(Area As Integer)
loadStatus = True
End Sub

Private Sub dcboClassification_LostFocus()
loadStatus = False
End Sub

Private Sub Form_Deactivate()
If Me.UpdateMode = True Then
    Me.ZOrder (vbBringToFront)
    Me.SetFocus
    Exit Sub
End If
End Sub

Private Sub Form_Load()
'On Error Resume Next

Call CenterForm(frmPO, MDIForm1)
loadStatus = False
Call ConnectDB(adoPO)
adoPO.CommandType = adCmdTable
adoPO.RecordSource = "tblPO"
adoPO.Refresh

Call ConnectDB(adoPODetails)
adoPODetails.CommandType = adCmdTable
adoPODetails.RecordSource = "qryPODetails"
adoPODetails.Refresh

Call ConnectDB(adotblPODetails)
adoPODetails.CommandType = adCmdTable
adoPODetails.RecordSource = "tblPODetails"
adoPODetails.Refresh

Call ConnectDB(adoClassification)
adoClassification.CommandType = adCmdTable
adoClassification.RecordSource = "tblstocktype"
adoClassification.Refresh

Call ConnectDB(adoKind)
adoKind.CommandType = adCmdTable
adoKind.RecordSource = "tblstocktypecategory"
adoKind.Refresh

qryPODetails (adoPO.Recordset.Fields("PO_ID"))
End Sub

Public Sub qryPODetails(criteria As Long)
adoPODetails.CommandType = adCmdText
adoPODetails.RecordSource = "select * from qryPODetails where PO_ID=" & criteria
adoPODetails.Refresh
End Sub

Private Sub grdPODetails_Error(ByVal DataError As Integer, Response As Integer)
Response = 0
End Sub

Private Sub txtQuantity_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 110 Or KeyCode = 190 Then SendKeys "{Backspace}"
End Sub

Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyDelete Then Exit Sub
If KeyAscii = vbKeyBack Then Exit Sub
If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
    KeyAscii = 0
    Exit Sub
End If
End Sub
