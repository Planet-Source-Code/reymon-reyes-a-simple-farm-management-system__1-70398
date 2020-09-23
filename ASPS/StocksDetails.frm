VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmStockDetails 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stocks Details"
   ClientHeight    =   8610
   ClientLeft      =   2385
   ClientTop       =   1590
   ClientWidth     =   11640
   ControlBox      =   0   'False
   FillColor       =   &H000080FF&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   11640
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
      DataSource      =   "adoBatch"
      Height          =   375
      Left            =   1200
      TabIndex        =   41
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
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
   Begin VB.TextBox txtBatchDate 
      DataField       =   "Date"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "M/d/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      DataSource      =   "adoStocks"
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
      Left            =   3360
      TabIndex        =   38
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo dcboSuppliers 
      Bindings        =   "StocksDetails.frx":0000
      DataField       =   "Supplier_ID"
      DataSource      =   "adoBatch"
      Height          =   360
      Left            =   1200
      TabIndex        =   1
      Top             =   1320
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   635
      _Version        =   393216
      Locked          =   -1  'True
      Style           =   2
      ListField       =   "Name"
      BoundColumn     =   "Supplier_ID"
      Text            =   "Suppliers"
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
   Begin MSAdodcLib.Adodc adoStocksGrid 
      Height          =   375
      Left            =   4560
      Top             =   1200
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
      Caption         =   "adoStocksGrid"
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
   Begin MSAdodcLib.Adodc adoStocks 
      Height          =   375
      Left            =   4560
      Top             =   840
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
      Caption         =   "adoStocks"
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
   Begin MSAdodcLib.Adodc adoSuppliers 
      Height          =   375
      Left            =   4560
      Top             =   480
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
      Caption         =   "adoSuppliers"
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
   Begin MSAdodcLib.Adodc adoUnitType 
      Height          =   375
      Left            =   6840
      Top             =   1200
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
      Caption         =   "adoUnitType"
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
   Begin MSAdodcLib.Adodc adoKind 
      Height          =   375
      Left            =   6840
      Top             =   840
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
   Begin MSAdodcLib.Adodc adoCategory 
      Height          =   375
      Left            =   6840
      Top             =   480
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
      Caption         =   "adoCategory"
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
   Begin VB.TextBox txtBatchID 
      DataField       =   "Batch_ID"
      DataSource      =   "adoStocks"
      Height          =   375
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   37
      Text            =   "txtBatchID"
      Top             =   7920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc adoBatch 
      Height          =   375
      Left            =   4560
      Top             =   120
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
      Caption         =   "batch"
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
      Caption         =   ">I"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   23
      ToolTipText     =   "Move to the last record"
      Top             =   8160
      Width           =   1335
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   22
      ToolTipText     =   "Move to the next record"
      Top             =   8160
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   21
      ToolTipText     =   "Move to the previous record"
      Top             =   8160
      Width           =   1335
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "I<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   20
      ToolTipText     =   "Move to the first record"
      Top             =   8160
      Width           =   1335
   End
   Begin VB.Frame fraReqOps 
      BackColor       =   &H0000C000&
      Height          =   5055
      Left            =   10200
      TabIndex        =   34
      Top             =   1800
      Width           =   1455
      Begin VB.CommandButton cmdShowInventory 
         Caption         =   "Inventory"
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
         TabIndex        =   39
         ToolTipText     =   "View inventory"
         Top             =   3840
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
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Close this form"
         Top             =   4440
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrintBatch 
         Caption         =   "&Print"
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
         ToolTipText     =   "Print the record directly"
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CommandButton cmdPreviewBatch 
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
         ToolTipText     =   "A print preview of the current record"
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelBatch 
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
         ToolTipText     =   "Cancel the current operation"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton cmdDeleteBatch 
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
         ToolTipText     =   "Delete the current record"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdEditBatch 
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
         ToolTipText     =   "Edit current record"
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdNewBatch 
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
         MaskColor       =   &H0080FF80&
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Add a new record"
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Particulars"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   120
      TabIndex        =   28
      Top             =   1800
      Width           =   9975
      Begin MSDataGridLib.DataGrid grdStocks 
         Bindings        =   "StocksDetails.frx":001B
         Height          =   3615
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   6376
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
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   8
         BeginProperty Column00 
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
         BeginProperty Column01 
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
         BeginProperty Column02 
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
         BeginProperty Column03 
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
         BeginProperty Column04 
            DataField       =   "Quantity"
            Caption         =   "Quantity"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
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
            DataField       =   "Subtotal"
            Caption         =   "Subtotal"
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
            Locked          =   -1  'True
            BeginProperty Column00 
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2039.811
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2204.788
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   615.118
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   2429.858
            EndProperty
         EndProperty
      End
   End
   Begin VB.TextBox txtReceiptNum 
      DataField       =   "Reciept_Number"
      DataSource      =   "adobatch"
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
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   0
      Top             =   840
      Width           =   1815
   End
   Begin VB.Frame fraStockDetail 
      BackColor       =   &H00C0FFC0&
      Height          =   2295
      Left            =   120
      TabIndex        =   29
      Top             =   5760
      Width           =   9975
      Begin VB.TextBox txtUOM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
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
         Left            =   6360
         TabIndex        =   19
         ToolTipText     =   "Cancel current operation"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
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
         Left            =   5040
         TabIndex        =   18
         ToolTipText     =   "Delete the selected item"
         Top             =   1800
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo dcboUOM 
         Bindings        =   "StocksDetails.frx":0037
         DataSource      =   "adoStocks"
         Height          =   360
         Left            =   360
         TabIndex        =   12
         Top             =   1800
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   635
         _Version        =   393216
         Locked          =   -1  'True
         Style           =   2
         ListField       =   "Unit"
         BoundColumn     =   "UnitType_ID"
         Text            =   "UOM"
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
      Begin MSDataListLib.DataCombo dcboKind 
         Bindings        =   "StocksDetails.frx":0051
         DataField       =   "StockTypeCategory_ID"
         DataSource      =   "adoStocks"
         Height          =   360
         Left            =   1680
         TabIndex        =   11
         Top             =   720
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   635
         _Version        =   393216
         Locked          =   -1  'True
         Style           =   2
         ListField       =   "Category"
         BoundColumn     =   "StockTypeCategory_ID"
         Text            =   "Kind"
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
         Bindings        =   "StocksDetails.frx":0067
         DataField       =   "StockType_ID"
         DataSource      =   "adoStocks"
         Height          =   360
         Left            =   1680
         TabIndex        =   10
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   635
         _Version        =   393216
         Locked          =   -1  'True
         Style           =   2
         ListField       =   "Type"
         BoundColumn     =   "StockType_ID"
         Text            =   "Classification"
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
      Begin VB.CommandButton cmdEditDetails 
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
         Left            =   3720
         TabIndex        =   17
         ToolTipText     =   "Edit the selected item"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdd 
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
         Left            =   2400
         TabIndex        =   16
         ToolTipText     =   "Add new stock item"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtRemarks 
         DataField       =   "Remarks"
         DataSource      =   "adoStocks"
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
         Left            =   5400
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   1200
         Width           =   4455
      End
      Begin VB.TextBox txtUnitCost 
         Alignment       =   1  'Right Justify
         DataField       =   "Unit_Cost"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         DataSource      =   "adoStocks"
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
         Left            =   5400
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   14
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtQuantity 
         Alignment       =   1  'Right Justify
         DataField       =   "Quantity"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         DataSource      =   "adoStocks"
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
         Left            =   5400
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   13
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks:"
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
         Left            =   4485
         TabIndex        =   36
         Top             =   1320
         Width           =   825
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Unit cost:"
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
         Left            =   4515
         TabIndex        =   35
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity:"
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
         Left            =   4560
         TabIndex        =   33
         Top             =   360
         Width           =   750
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Unit:"
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
         Left            =   1170
         TabIndex        =   32
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Kind:"
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
         Left            =   1125
         TabIndex        =   31
         Top             =   840
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Classification:"
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
         Left            =   345
         TabIndex        =   30
         Top             =   360
         Width           =   1200
      End
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1095
      Left            =   10200
      Top             =   720
      Width           =   1455
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   855
      Left            =   5280
      Shape           =   2  'Oval
      Top             =   8040
      Width           =   3135
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   495
      Left            =   4800
      Shape           =   4  'Rounded Rectangle
      Top             =   8400
      Width           =   5295
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      Height          =   855
      Index           =   1
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   8040
      Width           =   6615
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      Height          =   1335
      Left            =   7320
      Shape           =   4  'Rounded Rectangle
      Top             =   -600
      Width           =   4695
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   375
      Left            =   480
      Shape           =   3  'Circle
      Top             =   -120
      Width           =   615
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   375
      Left            =   840
      Top             =   -120
      Width           =   6975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier:"
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
      Left            =   270
      TabIndex        =   27
      Top             =   1440
      Width           =   870
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OR #:"
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
      Left            =   630
      TabIndex        =   26
      Top             =   960
      Width           =   510
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   25
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      DataSource      =   "Adodc1"
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
      Left            =   645
      TabIndex        =   24
      Top             =   480
      Width           =   495
   End
End
Attribute VB_Name = "frmStockDetails"
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

Dim loadStatus As Boolean 'status indicator to querying the adoKind
Dim kindStat As Boolean
Public UpdateMode As Boolean

Private Sub adoBatch_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub adoCategory_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub adoKind_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub adoStocks_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub adoStocksGrid_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub adoSuppliers_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub adoUnitType_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub cmdAdd_Click()
On Error GoTo ErrHandler

If cmdAdd.Caption = "&Add" Then
    UpdateMode = True
    cmdAdd.Caption = "&Save"
    cmdEditDetails.Enabled = False
    cmdDelete.Enabled = False
    cmdCancel.Enabled = True
    cmdFirst.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdLast.Enabled = False
    fraReqOps.Enabled = False
    
    'unlock the datacombo fields
    dcboClassification.Locked = False
    dcboKind.Locked = False
    'dcboUOM.Locked = False
    'unlock the text fields
    'txtStockName.Locked = False
    txtQuantity.Locked = False
    txtUnitCost.Locked = False
    txtRemarks.Locked = False
    'set adobatch to addnew
    adoStocks.Recordset.AddNew
    adoStocks.Recordset.Fields("Batch_ID") = adoBatch.Recordset.Fields("Batch_ID")
    dcboClassification.SetFocus
    
ElseIf cmdAdd.Caption = "&Save" Then
    'validate the fields
    If dcboClassification.Text = "" Then
        MsgBox "Please select a classification."
        dcboClassification.SetFocus
        Exit Sub
    ElseIf dcboKind.Text = "" Then
        MsgBox "Please select a kind."
        dcboKind.SetFocus
        Exit Sub
    ElseIf txtQuantity.Text = "" Then
        MsgBox "Please enter a quantity."
        txtQuantity.SetFocus
        Exit Sub
    ElseIf Not IsNumeric(txtQuantity.Text) Then
        MsgBox "Invalid Quantity value."
        txtQuantity.SetFocus
        Exit Sub
    ElseIf txtUnitCost.Text = "" Then
        MsgBox "Please select a unit cost."
        txtUnitCost.SetFocus
        Exit Sub
    ElseIf Not IsNumeric(txtUnitCost.Text) Then
        MsgBox "Invalid Unit Cost value."
        txtUnitCost.SetFocus
        Exit Sub
    Else
        Me.MousePointer = vbHourglass
        adoStocks.Recordset.Update
        cmdAdd.Caption = "&Add"
        cmdEditDetails.Enabled = True
        cmdDelete.Enabled = True
        cmdCancel.Enabled = False
        cmdFirst.Enabled = True
        cmdPrev.Enabled = True
        cmdNext.Enabled = True
        cmdLast.Enabled = True
        fraReqOps.Enabled = True
        'lock the datacombo fields
        dcboClassification.Locked = True
        dcboKind.Locked = True
        'dcboUOM.Locked = True
        'lock the text fields
        'txtStockName.Locked = True
        txtQuantity.Locked = True
        txtUnitCost.Locked = True
        txtRemarks.Locked = True
        
        'adoStocks.Recordset.Resync
        adoKind.CommandType = adCmdText
        adoKind.RecordSource = "select * from qryKinds" ' where StockType_ID=" & adoCategory.Recordset.Fields("StockType_ID")
        adoKind.Refresh
        Set dcboKind.RowSource = adoKind.Recordset
        dcboKind.ListField = "Category"
        dcboKind.ReFill

        adoStocks.Refresh
        Call qryStocksGrid(adoBatch.Recordset.Fields("Batch_ID"))
        adoStocksGrid.Refresh
        Me.MousePointer = vbDefault
        MsgBox "Record successfully updated."
        UpdateMode = False
    End If
End If

Exit Sub
ErrHandler:
    MsgBox Err.Description
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
On Error GoTo ErrHandler

adoKind.CommandType = adCmdText
adoKind.RecordSource = "select * from qryKinds" ' where StockType_ID=" & adoCategory.Recordset.Fields("StockType_ID")
adoKind.Refresh
Set dcboKind.RowSource = adoKind.Recordset
dcboKind.ListField = "Category"
dcboKind.ReFill

adoStocks.Recordset.Cancel
adoStocks.Refresh
adoStocksGrid.Refresh

cmdAdd.Caption = "&Add"
cmdDelete.Enabled = True
cmdEditDetails.Enabled = True
cmdCancel.Enabled = False
cmdFirst.Enabled = True
cmdPrev.Enabled = True
cmdNext.Enabled = True
cmdLast.Enabled = True
fraReqOps.Enabled = True
'lock the datacombo fields
dcboClassification.Locked = True
dcboKind.Locked = True
'dcboUOM.Locked = True
'lock the text fields
'txtStockName.Locked = True
txtQuantity.Locked = True
txtUnitCost.Locked = True
txtRemarks.Locked = True
UpdateMode = False

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdCancel2_Click()
adoBatch.Recordset.Cancel
adoBatch.Refresh
cmdNew.Caption = "&NEW"
cmdNew.Enabled = True
cmdDelete2.Enabled = True
cmdCancel2.Enabled = False
cmdEdit.Enabled = True

End Sub

Private Sub cmdCancelBatch_Click()
On Error GoTo ErrHandler

cmdNewBatch.Caption = "&New"
cmdNewBatch.Enabled = True
cmdDeleteBatch.Enabled = True
cmdEditBatch.Enabled = True
cmdCancelBatch.Enabled = False
fraStockDetail.Enabled = True
cmdPreviewBatch.Enabled = True
cmdPrintBatch.Enabled = True
cmdShowInventory.Enabled = True
cmdClose.Enabled = True
cmdFirst.Enabled = True
cmdPrev.Enabled = True
cmdNext.Enabled = True
cmdLast.Enabled = True
mskDate.Enabled = False
adoBatch.Recordset.Cancel
adoBatch.Refresh
Call qryStocksGrid(adoBatch.Recordset.Fields("Batch_ID"))
UpdateMode = False

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdCancelBatch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdCancelBatch.FontBold = True
cmdDeleteBatch.FontBold = False
cmdNewBatch.FontBold = False
cmdEditBatch.FontBold = False
cmdPreviewBatch.FontBold = False
cmdPrintBatch.FontBold = False
cmdShowInventory.FontBold = False
cmdClose.FontBold = False
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdClose.FontBold = True
cmdDeleteBatch.FontBold = False
cmdNewBatch.FontBold = False
cmdEditBatch.FontBold = False
cmdCancelBatch.FontBold = False
cmdPreviewBatch.FontBold = False
cmdPrintBatch.FontBold = False
cmdShowInventory.FontBold = False
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrHandler

If MsgBox("The selected RECORD will be deleted, are you sure you want to continue?" & vbCr & "This action is irreversible.", vbOKCancel + vbQuestion, "Delete") = vbOK Then
    Me.MousePointer = vbHourglass
    adoStocks.Recordset.Delete
    adoStocks.Recordset.Requery
    adoStocksGrid.Refresh
    Me.MousePointer = vbDefault
End If

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
On Error Resume Next
cmdNew.Caption = "&SAVE"
txtreceipt.Locked = False
supplier.Locked = False

If cmdNew.Caption = "&SAVE" Then
    cmdEdit.Enabled = False
    cmdDelete2.Enabled = False
    cmdCancel2.Enabled = True
    adoStocks.Recordset.Update
    Text1.SetFocus
    
    adoBatch.Recordset.Resync
    adoBatch.Refresh
    MsgBox "Record successfully updated."
    End If
End Sub

Private Sub cmdFind_Click()
On Error Resume Next
With adoStocks.Recordset
.MoveFirst
.Find "name ='" & InputBox("Enter STOCK NAME to Search:") & "'"
If .EOF = True Then
MsgBox "NO RECORD FOUND!!!", "Please try again"
.MoveFirst
End If
End With
End Sub

Private Sub cmdDeleteBatch_Click()
On Error GoTo ErrHandler

If MsgBox("The selected RECORD and its corresponding stocks will be deleted, are you sure you want to continue?" _
    & vbCr & "This action is irreversible.", vbOKCancel + vbExclamation, "Delete") = vbOK Then
    Me.MousePointer = vbHourglass
    adoBatch.Recordset.Delete
    adoBatch.Recordset.Requery
    Call qryStocksGrid(adoBatch.Recordset.Fields("Batch_ID"))
    Me.MousePointer = vbDefault
End If

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdDeleteBatch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdDeleteBatch.FontBold = True
cmdNewBatch.FontBold = False
cmdEditBatch.FontBold = False
cmdCancelBatch.FontBold = False
cmdPreviewBatch.FontBold = False
cmdPrintBatch.FontBold = False
cmdShowInventory.FontBold = False
cmdClose.FontBold = False
End Sub

Private Sub cmdEditBatch_Click()
On Error GoTo ErrHandler

UpdateMode = True
cmdNewBatch.Caption = "&Save"
cmdDeleteBatch.Enabled = False
cmdCancelBatch.Enabled = True
cmdEditBatch.Enabled = False
fraStockDetail.Enabled = False
cmdPreviewBatch.Enabled = False
cmdPrintBatch.Enabled = False
cmdShowInventory.Enabled = False
cmdClose.Enabled = False
cmdFirst.Enabled = False
cmdPrev.Enabled = False
cmdNext.Enabled = False
cmdLast.Enabled = False
txtReceiptNum.Locked = False
dcboSuppliers.Locked = False
txtReceiptNum.SetFocus
mskDate.Enabled = True

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdEditBatch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdEditBatch.FontBold = True
cmdNewBatch.FontBold = False
cmdDeleteBatch.FontBold = False
cmdCancelBatch.FontBold = False
cmdPreviewBatch.FontBold = False
cmdPrintBatch.FontBold = False
cmdShowInventory.FontBold = False
cmdClose.FontBold = False
End Sub

Private Sub cmdEditDetails_Click()
On Error GoTo ErrHandler

UpdateMode = True
cmdAdd.Caption = "&Save"
'MsgBox dcboClassification.Text
adoKind.CommandType = adCmdText
adoKind.RecordSource = "Select * from qryKinds where Type = '" & dcboClassification.Text & "'"
adoKind.Refresh

Set dcboKind.RowSource = adoKind.Recordset
dcboKind.ListField = "Category"
cmdEditDetails.Enabled = False
cmdDelete.Enabled = False
cmdCancel.Enabled = True
cmdDelete.Enabled = False
cmdCancel.Enabled = True
cmdFirst.Enabled = False
cmdPrev.Enabled = False
cmdNext.Enabled = False
cmdLast.Enabled = False
fraReqOps.Enabled = False

'unlock the datacombo fields
dcboClassification.Locked = False
dcboKind.Locked = False
'dcboUOM.Locked = False
'unlock the text fields
'txtStockName.Locked = False
txtQuantity.Locked = False
txtUnitCost.Locked = False
txtRemarks.Locked = False
dcboClassification.SetFocus

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdFirst_Click()
cmdNext.Enabled = True
adoBatch.Recordset.MoveFirst
Call qryStocksGrid(adoBatch.Recordset.Fields("Batch_ID"))
'MsgBox "This is the first record!"
End Sub

Private Sub cmdLast_Click()
cmdPrev.Enabled = True
adoBatch.Recordset.MoveLast
Call qryStocksGrid(adoBatch.Recordset.Fields("Batch_ID"))
'MsgBox "This is the last record!"
End Sub

Private Sub cmdNew_Click()
'On Error Resume Next
If cmdNew.Caption = "&NEW" Then
    cmdNew.Caption = "&SAVE"
    cmdEdit.Enabled = False
    cmdDelete2.Enabled = False
    cmdCancel2.Enabled = True
    
    adoBatch.Recordset.AddNew
    txtReceiptNum.Locked = False
    txtSupplier.Locked = False
    txtReceiptNum.SetFocus

ElseIf cmdNew.Caption = "&SAVE" Then
    If txtreceipt.Text = "" Then
    MsgBox "RECEIPT NUMBER IS REQUIRED!!!"
    txtreceipt.SetFocus
    If txtSupplier.Text = "" Then
    MsgBox "SUPPLIER'S NAME IS REQUIRED!!!"
   txtSupplier.SetFocus
    Exit Sub
    
Else
    cmdNew.Caption = "&NEW"
    adoStocks.Recordset.Update
    
    cmdDelete2.Enabled = True
    cmdEdit.Enabled = True
    cmdCancel2.Enabled = False
    txtreceipt.Locked = True
   txtSupplier.Locked = True
    
     adoBatch.Recordset.Resync
    adoBatch.Refresh
    MsgBox "Record successfully updated."
    End If
End If
End If
End Sub

Private Sub cmdNewBatch_Click()
'On Error GoTo ErrHandler

If cmdNewBatch.Caption = "&New" Then
    UpdateMode = True
    cmdNewBatch.Caption = "&Save"
    cmdEditBatch.Enabled = False
    cmdDeleteBatch.Enabled = False
    cmdCancelBatch.Enabled = True
    fraStockDetail.Enabled = False
    cmdPreviewBatch.Enabled = False
    cmdPrintBatch.Enabled = False
    cmdShowInventory.Enabled = False
    cmdClose.Enabled = False
    cmdFirst.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdLast.Enabled = False
    
    adoBatch.Recordset.AddNew
    Call qryStocksGrid(0)
    txtReceiptNum.Locked = False
    dcboSuppliers.Locked = False
    dcboSuppliers.Text = ""
    mskDate.Enabled = True
    mskDate.Text = Format(Date, "MM/dd/yyyy")
    txtReceiptNum.SetFocus

ElseIf cmdNewBatch.Caption = "&Save" Then
    If txtReceiptNum.Text = "" Then
        MsgBox "Receipt Number is required."
        txtReceiptNum.SetFocus
        Exit Sub
    ElseIf dcboSuppliers.Text = "" Then
        MsgBox "Please select a Supplier."
        dcboSuppliers.SetFocus
        Exit Sub
    ElseIf mskDate.Text = "" Or Not IsDate(mskDate.Text) Then
        MsgBox "Invalid date. Please provide appropriate date."
        mskDate.SetFocus
        Exit Sub
    Else
        Me.MousePointer = vbHourglass
        cmdNewBatch.Caption = "&New"
        adoBatch.Recordset.Update
        
        cmdDeleteBatch.Enabled = True
        cmdEditBatch.Enabled = True
        cmdCancelBatch.Enabled = False
        txtReceiptNum.Locked = True
        mskDate.Enabled = False
        dcboSuppliers.Locked = True
        fraStockDetail.Enabled = True
        cmdPreviewBatch.Enabled = True
        cmdPrintBatch.Enabled = True
        cmdShowInventory.Enabled = True
        cmdClose.Enabled = True
        cmdFirst.Enabled = True
        cmdPrev.Enabled = True
        cmdNext.Enabled = True
        cmdLast.Enabled = True
        adoBatch.Recordset.Resync
        Me.MousePointer = vbDefault
        MsgBox "Record successfully updated."
        UpdateMode = False
    End If
End If

'Exit Sub
'ErrHandler:
 '   MsgBox Err.Description
End Sub

Private Sub cmdNewBatch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdNewBatch.FontBold = True
cmdEditBatch.FontBold = False
cmdDeleteBatch.FontBold = False
cmdCancelBatch.FontBold = False
cmdPreviewBatch.FontBold = False
cmdPrintBatch.FontBold = False
cmdShowInventory.FontBold = False
cmdClose.FontBold = False
End Sub

Private Sub cmdNext_Click()
With adoBatch.Recordset
.MoveNext
If .EOF Then
   .MoveLast
   Call qryStocksGrid(.Fields("Batch_ID"))
    cmdPrev.Enabled = True
    cmdNext.Enabled = False
Else
    cmdPrev.Enabled = True
End If
Call qryStocksGrid(.Fields("Batch_ID"))
End With
End Sub

Private Sub cmdPrev_Click()
With adoBatch.Recordset
.MovePrevious
If .BOF Then
    .MoveFirst
    Call qryStocksGrid(adoBatch.Recordset.Fields("Batch_ID"))
    cmdPrev.Enabled = False
    cmdNext.Enabled = True
Else
    cmdNext.Enabled = True
End If
Call qryStocksGrid(.Fields("Batch_ID"))
End With
End Sub

Private Sub cmdPreviewBatch_Click()
On Error GoTo ErrHandler

Me.MousePointer = vbHourglass
envAmadeus.cmdStockDetails_Grouping adoBatch.Recordset.Fields("Batch_ID")
rptStockDetails.Show
rptStockDetails.WindowState = vbMaximized
Me.MousePointer = vbDefault

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdPreviewBatch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdPreviewBatch.FontBold = True
cmdDeleteBatch.FontBold = False
cmdNewBatch.FontBold = False
cmdEditBatch.FontBold = False
cmdCancelBatch.FontBold = False
cmdPrintBatch.FontBold = False
cmdShowInventory.FontBold = False
cmdClose.FontBold = False
End Sub

Private Sub cmdPrintBatch_Click()
On Error GoTo ErrHandler

Me.MousePointer = vbHourglass
envAmadeus.cmdStockDetails_Grouping adoBatch.Recordset.Fields("Batch_ID")
rptStockDetails.PrintReport
Unload rptStockDetails
Me.MousePointer = vbDefault

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdPrintBatch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdPrintBatch.FontBold = True
cmdDeleteBatch.FontBold = False
cmdNewBatch.FontBold = False
cmdEditBatch.FontBold = False
cmdCancelBatch.FontBold = False
cmdPreviewBatch.FontBold = False
cmdShowInventory.FontBold = False
cmdClose.FontBold = False
End Sub

Private Sub cmdShowInventory_Click()
frmInventory.Show
frmInventory.SetFocus
End Sub

Private Sub cmdShowInventory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdShowInventory.FontBold = True
cmdDeleteBatch.FontBold = False
cmdNewBatch.FontBold = False
cmdEditBatch.FontBold = False
cmdCancelBatch.FontBold = False
cmdPreviewBatch.FontBold = False
cmdPrintBatch.FontBold = False
cmdClose.FontBold = False
End Sub

Private Sub dcboKind_Change()
'test
If kindStat = False Then Exit Sub
If adoKind.Recordset.Bookmark = 0 Then Exit Sub
If dcboKind.Text = "" Then Exit Sub
adoKind.Recordset.Bookmark = dcboKind.SelectedItem
txtUOM.Text = adoKind.Recordset.Fields(3).Value
End Sub

Private Sub dcboKind_Click(Area As Integer)
kindStat = True
End Sub

Private Sub Form_Activate()
On Error Resume Next
If UpdateMode = True Then Exit Sub
Call CenterForm(frmStockDetails, MDIForm1)
adoSuppliers.Recordset.Requery
dcboSuppliers.ListField = "Name"
dcboSuppliers.ReFill
grdStocks.Refresh
End Sub

Private Sub Form_Deactivate()
If Me.UpdateMode = True Then
    Me.ZOrder (vbBringToFront)
    Me.SetFocus
    Exit Sub
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdNewBatch.FontBold = False
cmdEditBatch.FontBold = False
cmdDeleteBatch.FontBold = False
cmdCancelBatch.FontBold = False
cmdPreviewBatch.FontBold = False
cmdPrintBatch.FontBold = False
cmdShowInventory.FontBold = False
cmdClose.FontBold = False
End Sub

Private Sub grdStocks_Error(ByVal DataError As Integer, Response As Integer)
Response = 0
End Sub

Private Sub grdStocks_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim tempStock As Long
If adoStocksGrid.Recordset.EOF Then
    tempStock = 0
Else
    tempStock = adoStocksGrid.Recordset.Fields("Stock_ID")
End If
'LastRow = 0
'LastCol = 0
adoStocks.CommandType = adCmdText
adoStocks.RecordSource = "Select * from tblStock where Stock_ID=" & tempStock 'adoStocksGrid.Recordset.Fields("Stock_ID")
adoStocks.Refresh

txtUOM.Text = adoStocksGrid.Recordset.Fields("Unit")
'MsgBox "RowColChange"
End Sub

Private Sub dcboClassification_Change()
If loadStatus = False Then 'if loadStatus global variable is false then the next
    Exit Sub               'lines of code will not be executed
End If
'query adokind to display the corresponding Kind based on the selected classification
'in the classification datacombo
adoCategory.Recordset.Bookmark = dcboClassification.SelectedItem
adoKind.RecordSource = "Select * from qryKinds where StockType_ID = " & adoCategory.Recordset.Fields("StockType_ID") 'dcboClassification.SelectedItem
adoKind.Refresh

Set dcboKind.RowSource = adoKind.Recordset
dcboKind.ListField = "Category"

dcboKind.Text = ""
dcboKind.ReFill
dcboKind.Refresh
End Sub

Private Sub dcboClassification_Click(Area As Integer)
'set loadStatus global variable so that the query code in the dcboClassification_Change()
'will be executed
loadStatus = True
End Sub

Private Sub dcboClassification_LostFocus()
loadStatus = False
End Sub

Private Sub Form_Initialize()
loadStatus = False
kindStat = False
End Sub

Private Sub Form_Load()
'connect and set the connectionstring, commandtype, recordset properties of the adodcs
Call ConnectDB(adoBatch)
adoBatch.CommandType = adCmdTable
adoBatch.RecordSource = "tblBatch"
adoBatch.Refresh
Call ConnectDB(adoSuppliers)
adoSuppliers.CommandType = adCmdTable
adoSuppliers.RecordSource = "tblSupplier"
adoSuppliers.Refresh
Set dcboSuppliers.RowSource = adoSuppliers.Recordset
dcboSuppliers.ListField = "Name"
dcboSuppliers.BoundColumn = "Supplier_ID"
Set dcboSuppliers.DataSource = adoBatch.Recordset
dcboSuppliers.DataField = "Supplier_ID"
Call ConnectDB(adoStocks)
adoStocks.CommandType = adCmdTable
adoStocks.RecordSource = "tblStock"
adoStocks.Refresh
Call ConnectDB(adoCategory)
adoCategory.CommandType = adCmdTable
adoCategory.RecordSource = "tblstocktype"
adoCategory.Refresh
Call ConnectDB(adoKind)
adoKind.CommandType = adCmdText
adoKind.RecordSource = "select * from qryKinds" ' where StockType_ID=" & adoCategory.Recordset.Fields("StockType_ID")
adoKind.Refresh
Call ConnectDB(adoUnitType)
adoUnitType.CommandType = adCmdTable
adoUnitType.RecordSource = "tblunittype"
adoUnitType.Refresh
Call ConnectDB(adoStocksGrid)
adoStocksGrid.CommandType = adCmdText
Call qryStocksGrid(adoBatch.Recordset.Fields("Batch_ID"))
End Sub

Private Sub Label1_Click()
'Label2.Caption = Date
End Sub

Private Sub txtBatchDate_Change()
'txtBatchDate.Text = Date
End Sub

Private Function generateTempBatchID() As Long 'this code is not implemented so
Dim tempID As Long                             'just disregard

adoBatch.Refresh
adoBatch.Recordset.MoveLast
tempID = adoBatch.Recordset.Fields("Batch_ID") + 1
generateTempBatchID = tempID
adoBatch.Refresh
End Function

Private Sub qryStocksGrid(criteria As Long) 'code for querying the all the stocks
adoStocksGrid.CommandType = adCmdText       'of a certain batch record
adoStocksGrid.RecordSource = "Select * from qryStocksBatch where Batch_ID=" & criteria
adoStocksGrid.Refresh
End Sub

Private Sub txtQuantity_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 110 Or KeyCode = 190 Then SendKeys "{Backspace}"
End Sub

Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyDelete Then Exit Sub
If KeyAscii = vbKeyBack Then Exit Sub
If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
    KeyAscii = 0
    Exit Sub
End If
End Sub

Private Sub txtReceiptNum_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyDelete Then Exit Sub
If KeyAscii = vbKeyBack Then Exit Sub
If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
    KeyAscii = 0
    Exit Sub
End If
End Sub

Private Sub txtUnitCost_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyDelete Then Exit Sub
If KeyAscii = vbKeyBack Then Exit Sub
If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
    KeyAscii = 0
    Exit Sub
End If
End Sub
