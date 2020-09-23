VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRequest 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Request"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11265
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   11265
   Begin VB.TextBox txtQtyStockID 
      DataField       =   "Stock_ID"
      DataSource      =   "adoQtyRemain"
      Height          =   375
      Left            =   6000
      TabIndex        =   30
      Text            =   "txtQtyStockID"
      Top             =   8520
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSAdodcLib.Adodc adoQtyRemain 
      Height          =   375
      Left            =   6840
      Top             =   8520
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
      RecordSource    =   $"frmRequest.frx":0000
      Caption         =   "adoQtyRemain"
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
   Begin VB.TextBox txtRequestRemarks 
      DataField       =   "Remarks"
      DataSource      =   "adoRequest"
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
      TabIndex        =   4
      Top             =   1800
      Width           =   5655
   End
   Begin MSAdodcLib.Adodc adoRequestItems 
      Height          =   375
      Left            =   9000
      Top             =   8520
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "adoRequestItems"
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
      Caption         =   ">l"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   22
      ToolTipText     =   "Go to last record."
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   21
      ToolTipText     =   "Go to next record"
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   20
      ToolTipText     =   "Go to previous record."
      Top             =   8040
      Width           =   1215
   End
   Begin MSMask.MaskEdBox mskRequestDate 
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
      DataSource      =   "adoRequest"
      Height          =   350
      Left            =   5640
      TabIndex        =   5
      Top             =   840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      _Version        =   393216
      AllowPrompt     =   -1  'True
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
      Format          =   "mm/dd/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.Frame fraRequestItems 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Request Items"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   120
      TabIndex        =   27
      Top             =   2280
      Width           =   9615
      Begin VB.CommandButton cmdCancelRequestItem 
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
         Left            =   6600
         TabIndex        =   18
         ToolTipText     =   "Cancel current operation."
         Top             =   5160
         Width           =   1335
      End
      Begin VB.TextBox txtReqQuantity 
         Alignment       =   1  'Right Justify
         DataField       =   "Quantity"
         DataSource      =   "adoRequestDetail"
         Height          =   375
         Left            =   4920
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   14
         Top             =   4680
         Width           =   1575
      End
      Begin VB.CommandButton cmdDeleteRequestItem 
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
         Left            =   5280
         TabIndex        =   17
         ToolTipText     =   "Delete a request item."
         Top             =   5160
         Width           =   1335
      End
      Begin VB.CommandButton cmdEditRequestItem 
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
         Left            =   3960
         TabIndex        =   16
         ToolTipText     =   "Modify a request item."
         Top             =   5160
         Width           =   1335
      End
      Begin VB.CommandButton cmdRequestItem 
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
         Left            =   2640
         TabIndex        =   15
         ToolTipText     =   "Add a request item."
         Top             =   5160
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid grdRequestItems 
         Bindings        =   "frmRequest.frx":0051
         Height          =   4335
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   7646
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
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   7
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
         BeginProperty Column02 
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
         BeginProperty Column05 
            DataField       =   "SumOfQuantity"
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            BeginProperty Column00 
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2204.788
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1995.024
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   480.189
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
            EndProperty
         EndProperty
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity:"
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
         Left            =   3960
         TabIndex        =   29
         Top             =   4800
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0000C000&
      Height          =   4455
      Left            =   9960
      TabIndex        =   26
      Top             =   2280
      Width           =   1335
      Begin VB.CommandButton Command1 
         Caption         =   "Find"
         Height          =   615
         Left            =   120
         TabIndex        =   31
         Top             =   4800
         Width           =   1095
      End
      Begin VB.CommandButton cmdNewRequest 
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
         TabIndex        =   6
         ToolTipText     =   "Add a new request."
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdEditRequest 
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
         TabIndex        =   7
         ToolTipText     =   "Modify the current request."
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdDeleteRequest 
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
         TabIndex        =   8
         ToolTipText     =   "Delete the current request."
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelRequest 
         Caption         =   "Cancel"
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
         ToolTipText     =   "Cancel current operation."
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton cmdPreviewRequest 
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
         TabIndex        =   10
         ToolTipText     =   "Print preview of the request."
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton cmdPrintRequest 
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
         TabIndex        =   11
         ToolTipText     =   "Print the request."
         Top             =   3240
         Width           =   1095
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
         TabIndex        =   12
         ToolTipText     =   "Close this form."
         Top             =   3840
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "l<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   19
      ToolTipText     =   "Go to first record."
      Top             =   8040
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc adoGrower 
      Height          =   375
      Left            =   9000
      Top             =   8160
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "adoGrower"
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
   Begin MSDataListLib.DataCombo dcboGrower 
      Bindings        =   "frmRequest.frx":006F
      DataField       =   "Growers_ID"
      DataSource      =   "adoRequest"
      Height          =   360
      Left            =   1200
      TabIndex        =   2
      Top             =   1320
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   635
      _Version        =   393216
      Locked          =   -1  'True
      Style           =   2
      ListField       =   "Grower"
      BoundColumn     =   "Growers_ID"
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
   Begin VB.TextBox txtRequestID 
      DataField       =   "Request_ID"
      DataSource      =   "adoRequest"
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
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo dcboFieldman 
      Bindings        =   "frmRequest.frx":0087
      DataField       =   "Fieldman_ID"
      DataSource      =   "adoRequest"
      Height          =   360
      Left            =   1200
      TabIndex        =   1
      Top             =   840
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   635
      _Version        =   393216
      Locked          =   -1  'True
      Style           =   2
      ListField       =   "Name"
      BoundColumn     =   "Fieldman_ID"
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
   Begin MSAdodcLib.Adodc adoFieldman 
      Height          =   375
      Left            =   9000
      Top             =   7800
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "adoFieldman"
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
   Begin MSAdodcLib.Adodc adoRequest 
      Height          =   375
      Left            =   9000
      Top             =   7440
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "adoRequest"
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
   Begin MSAdodcLib.Adodc adoRequestDetail 
      Height          =   375
      Left            =   9000
      Top             =   7080
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "adoRequestDetail"
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
      Height          =   255
      Left            =   -120
      Shape           =   4  'Rounded Rectangle
      Top             =   8280
      Width           =   10935
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   2415
      Left            =   9960
      Top             =   -120
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   615
      Left            =   7200
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   230
      Left            =   1920
      Top             =   0
      Width           =   5895
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C000&
      BorderWidth     =   30
      X1              =   1200
      X2              =   1920
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   28
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Request ID:"
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
      Left            =   120
      TabIndex        =   25
      Top             =   480
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
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
      Left            =   5160
      TabIndex        =   24
      Top             =   960
      Width           =   435
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Leadman:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grower:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1440
      Width           =   735
   End
End
Attribute VB_Name = "frmRequest"
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
Public UpdateMode As Boolean

Private Sub adoFieldman_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub adoGrower_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub adoQtyRemain_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub adoRequest_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub adoRequestDetail_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub adoRequestItems_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub cmdCancelRequest_Click()
On Error GoTo ErrHandler

adoRequest.Recordset.Cancel
adoRequest.Refresh
cmdNewRequest.Caption = "&New"
cmdDeleteRequest.Enabled = True
cmdEditRequest.Enabled = True
cmdCancelRequest.Enabled = False
fraRequestItems.Enabled = True
cmdFirst.Enabled = True
cmdPrevious.Enabled = True
cmdNext.Enabled = True
cmdLast.Enabled = True
cmdPreviewRequest.Enabled = True
cmdPrintRequest.Enabled = True
cmdClose.Enabled = True
dcboFieldman.Locked = True
dcboGrower.Locked = True
mskRequestDate.Enabled = False
Call qryRequestItems(adoRequest.Recordset.Fields("Request_ID"))
UpdateMode = False

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdCancelRequest_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdNewRequest.FontBold = False
cmdEditRequest.FontBold = False
cmdDeleteRequest.FontBold = False
cmdCancelRequest.FontBold = True
cmdPreviewRequest.FontBold = False
cmdPrintRequest.FontBold = False
cmdClose.FontBold = False
End Sub

Private Sub cmdCancelRequestItem_Click()
On Error GoTo ErrHandler

adoRequestDetail.Recordset.Cancel
adoRequestDetail.Refresh
txtReqQuantity.Locked = True
cmdEditRequestItem.Caption = "&Edit"
cmdRequestItem.Enabled = True
cmdDeleteRequestItem.Enabled = True
cmdCancelRequestItem.Enabled = False
cmdFirst.Enabled = True
cmdPrevious.Enabled = True
cmdNext.Enabled = True
cmdLast.Enabled = True
UpdateMode = False

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdNewRequest.FontBold = False
cmdEditRequest.FontBold = False
cmdDeleteRequest.FontBold = False
cmdCancelRequest.FontBold = False
cmdPreviewRequest.FontBold = False
cmdPrintRequest.FontBold = False
cmdClose.FontBold = True
End Sub

Private Sub cmdDeleteRequest_Click()
On Error GoTo ErrHandler

If MsgBox("The selected RECORD will be deleted, are you sure you want to continue?" & vbCr & "This action is irreversible.", vbOKCancel + vbExclamation, "Delete") = vbOK Then
    Me.MousePointer = vbHourglass
    adoRequest.Recordset.Delete
    adoRequest.Recordset.Requery
    Call qryRequestItems(Me.adoRequest.Recordset.Fields("Request_ID"))
    Me.MousePointer = vbDefault
End If

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdDeleteRequest_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdNewRequest.FontBold = False
cmdEditRequest.FontBold = False
cmdDeleteRequest.FontBold = True
cmdCancelRequest.FontBold = False
cmdPreviewRequest.FontBold = False
cmdPrintRequest.FontBold = False
cmdClose.FontBold = False
End Sub

Private Sub cmdDeleteRequestItem_Click()
On Error GoTo ErrHandler

If MsgBox("The selected RECORD will be deleted, are you sure you want to continue?" & vbCr & "This action is irreversible.", vbOKCancel + vbExclamation, "Delete") = vbOK Then
    Me.MousePointer = vbHourglass
    adoRequestDetail.Recordset.Delete
    adoRequestDetail.Recordset.Requery
    adoRequestItems.Refresh
    Me.MousePointer = vbDefault
End If

Exit Sub
ErrHandler:
    MsgBox Err.Description
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdEditRequest_Click()
On Error GoTo ErrHandler

UpdateMode = True
cmdNewRequest.Caption = "&Save"
cmdEditRequest.Enabled = False
cmdDeleteRequest.Enabled = False
cmdCancelRequest.Enabled = True
fraRequestItems.Enabled = False
cmdFirst.Enabled = False
cmdPrevious.Enabled = False
cmdNext.Enabled = False
cmdLast.Enabled = False
cmdPreviewRequest.Enabled = False
cmdPrintRequest.Enabled = False
cmdClose.Enabled = False
mskRequestDate.Enabled = True
dcboFieldman.Locked = False
dcboGrower.Locked = False
dcboFieldman.SetFocus

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdEditRequest_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdNewRequest.FontBold = False
cmdEditRequest.FontBold = True
cmdDeleteRequest.FontBold = False
cmdCancelRequest.FontBold = False
cmdPreviewRequest.FontBold = False
cmdPrintRequest.FontBold = False
cmdClose.FontBold = False
End Sub

Private Sub cmdEditRequestItem_Click()
Dim tempTtlQty As Integer

On Error Resume Next

adoQtyRemain.Refresh
adoQtyRemain.Recordset.Find "Stock_ID=" & adoRequestItems.Recordset.Fields("Stock_ID")
tempTtlQty = adoQtyRemain.Recordset.Fields("QtyRemaining") + adoRequestItems.Recordset.Fields("SumOfQuantity")

If cmdEditRequestItem.Caption = "&Edit" Then
    UpdateMode = True
    cmdEditRequestItem.Caption = "&Save"
    cmdRequestItem.Enabled = False
    cmdCancelRequestItem.Enabled = True
    cmdDeleteRequestItem.Enabled = False
    cmdFirst.Enabled = False
    cmdPrevious.Enabled = False
    cmdNext.Enabled = False
    cmdLast.Enabled = False
    txtReqQuantity.Locked = False
    txtReqQuantity.SetFocus
ElseIf cmdEditRequestItem.Caption = "&Save" Then
    If txtReqQuantity.Text = "" Or Val(txtReqQuantity.Text) = 0 Or Not IsNumeric(txtReqQuantity.Text) Then
        MsgBox "Invalid Quantity."
        txtReqQuantity.SetFocus
        Exit Sub
    End If
    If Val(txtReqQuantity.Text) > tempTtlQty Then 'adoQtyRemain.Recordset.Fields("QtyRemaining") Then
        MsgBox "The requested quantity you entered is greater that the " & vbCr & _
          "remaining quantity of the item you selected." & vbCr & "You can either enter a " & _
          "value greater than or equal to the remaining quantity of " & tempTtlQty & " .", vbInformation
        txtReqQuantity.SetFocus
        Exit Sub
    End If
    Me.MousePointer = vbHourglass
    cmdEditRequestItem.Caption = "&Edit"
    cmdCancelRequestItem.Enabled = False
    cmdDeleteRequestItem.Enabled = True
    cmdRequestItem.Enabled = True
    cmdFirst.Enabled = True
    cmdPrevious.Enabled = True
    cmdNext.Enabled = True
    cmdLast.Enabled = True
    adoRequestDetail.Recordset.Update
    adoRequestDetail.Refresh
    adoRequestItems.Refresh
    txtReqQuantity.Locked = True
    Me.MousePointer = vbDefault
    MsgBox "Record successfully updated."
    UpdateMode = False
End If
End Sub

Private Sub cmdFirst_Click()
adoRequest.Recordset.MoveFirst
Call qryRequestItems(Me.adoRequest.Recordset.Fields("Request_ID"))
End Sub

Private Sub cmdLast_Click()
adoRequest.Recordset.MoveLast
Call qryRequestItems(adoRequest.Recordset.Fields("Request_ID"))
End Sub

Private Sub cmdNewRequest_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdNewRequest.FontBold = True
cmdEditRequest.FontBold = False
cmdDeleteRequest.FontBold = False
cmdCancelRequest.FontBold = False
cmdPreviewRequest.FontBold = False
cmdPrintRequest.FontBold = False
cmdClose.FontBold = False
End Sub

Private Sub cmdNext_Click()
cmdPrevious.Enabled = True
adoRequest.Recordset.MoveNext
If adoRequest.Recordset.EOF Then
    cmdNext.Enabled = False
    adoRequest.Recordset.MoveLast
    Call qryRequestItems(adoRequest.Recordset.Fields("Request_ID"))
End If
Call qryRequestItems(adoRequest.Recordset.Fields("Request_ID"))
End Sub

Private Sub cmdPreviewRequest_Click()
On Error GoTo ErrHandler

Me.MousePointer = vbHourglass
If envAmadeus.rscmdRequest.State = adStateOpen Then
    envAmadeus.rscmdRequest.Close
End If
Unload rptRequest
envAmadeus.cmdRequest adoRequest.Recordset.Fields("Request_ID")
rptRequest.Show
rptRequest.WindowState = vbMaximized
Me.MousePointer = vbDefault

Exit Sub
ErrHandler:
    MsgBox Err.Description
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdPreviewRequest_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdNewRequest.FontBold = False
cmdEditRequest.FontBold = False
cmdDeleteRequest.FontBold = False
cmdCancelRequest.FontBold = False
cmdPreviewRequest.FontBold = True
cmdPrintRequest.FontBold = False
cmdClose.FontBold = False
End Sub

Private Sub cmdPrevious_Click()
cmdNext.Enabled = True
adoRequest.Recordset.MovePrevious
If adoRequest.Recordset.BOF Then
    cmdPrevious.Enabled = False
    adoRequest.Recordset.MoveFirst
    Call qryRequestItems(adoRequest.Recordset.Fields("Request_ID"))
End If
Call qryRequestItems(adoRequest.Recordset.Fields("Request_ID"))
End Sub

Private Sub cmdPrintRequest_Click()
On Error GoTo ErrHandler

Me.MousePointer = vbHourglass
envAmadeus.cmdRequest adoRequest.Recordset.Fields("Request_ID")
rptRequest.PrintReport
Unload rptRequest
Me.MousePointer = vbDefault

Exit Sub
ErrHandler:
    MsgBox Err.Description
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdPrintRequest_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdNewRequest.FontBold = False
cmdEditRequest.FontBold = False
cmdDeleteRequest.FontBold = False
cmdCancelRequest.FontBold = False
cmdPreviewRequest.FontBold = False
cmdPrintRequest.FontBold = True
cmdClose.FontBold = False
End Sub

Private Sub cmdRequestItem_Click()
frmRequestDetails.Show
End Sub

Private Sub Form_Activate()
If UpdateMode = True Then Exit Sub
adoFieldman.Recordset.Requery
adoGrower.Recordset.Requery
dcboFieldman.ListField = "Name"
dcboFieldman.ReFill
dcboGrower.ListField = "Grower"
dcboGrower.ReFill
grdRequestItems.Refresh
End Sub

Private Sub Form_Deactivate()
If Me.UpdateMode = True Then
    Me.ZOrder (vbBringToFront)
'    MsgBox "Please save or cancel current operation to continue."
    Me.SetFocus
    Exit Sub
End If
End Sub

Private Sub Form_Load()
Call CenterForm(frmRequest, MDIForm1)
'setup the adodcs
Call ConnectDB(adoRequest)
adoRequest.CommandType = adCmdTable
adoRequest.RecordSource = "tblRequest"
adoRequest.Refresh
Call ConnectDB(adoFieldman)
adoFieldman.CommandType = adCmdTable
adoFieldman.RecordSource = "FieldManNames"
adoFieldman.Refresh
'set the rowsource,listfield,datasource,datafield, and boundcolumn
'properties of the bound column
Set dcboFieldman.RowSource = adoFieldman.Recordset
dcboFieldman.ListField = "Name"
Set dcboFieldman.DataSource = adoRequest
dcboFieldman.DataField = "Fieldman_ID"
dcboFieldman.BoundColumn = "Fieldman_ID"
Call ConnectDB(adoGrower)
adoGrower.CommandType = adCmdTable
adoGrower.RecordSource = "tblGrower"
adoGrower.Refresh
'set the rowsource,listfield,datasource,datafield, and boundcolumn
'properties of the bound column
Set dcboGrower.RowSource = adoGrower.Recordset
dcboGrower.ListField = "Grower"
Set dcboFieldman.DataSource = adoRequest
dcboGrower.DataField = "Growers_ID"
dcboGrower.BoundColumn = "Growers_ID"
Call ConnectDB(adoRequestItems)
adoRequestItems.CommandType = adCmdTable
adoRequestItems.RecordSource = "qryRequestItems"
adoRequestItems.Refresh
Call ConnectDB(adoRequestDetail)
adoRequestDetail.CommandType = adCmdTable
adoRequestDetail.RecordSource = "tblRequestDetail"
adoRequestDetail.Refresh
Call ConnectDB(adoQtyRemain)
adoQtyRemain.CommandType = adCmdText
adoQtyRemain.RecordSource = "SELECT qryInventory.Stock_ID, qryInventory.QtyRemaining FROM qryInventory"
adoQtyRemain.Refresh

Call qryRequestItems(adoRequest.Recordset.Fields("Request_ID"))
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdNewRequest_Click()
On Error GoTo ErrHandler

If cmdNewRequest.Caption = "&New" Then
    UpdateMode = True
    cmdNewRequest.Caption = "&Save"
    cmdEditRequest.Enabled = False
    cmdDeleteRequest.Enabled = False
    cmdCancelRequest.Enabled = True
    fraRequestItems.Enabled = False
    cmdFirst.Enabled = False
    cmdPrevious.Enabled = False
    cmdNext.Enabled = False
    cmdLast.Enabled = False
    cmdPreviewRequest.Enabled = False
    cmdPrintRequest.Enabled = False
    cmdClose.Enabled = False
    
    adoRequest.Recordset.AddNew
    Call qryRequestItems(0)
    mskRequestDate.Enabled = True
    mskRequestDate.Text = Format(Date, "MM/dd/yyyy")
    dcboFieldman.Locked = False
    dcboFieldman.Text = ""
    dcboGrower.Locked = False
    dcboGrower.Text = ""
    dcboFieldman.SetFocus

ElseIf cmdNewRequest.Caption = "&Save" Then
    If mskRequestDate.Text = "" Or Not IsDate(mskRequestDate.Text) Then
        MsgBox "Invalid Date. Please enter the appropriate date."
        mskRequestDate.SetFocus
        Exit Sub
    ElseIf dcboFieldman.Text = "" Then
        MsgBox "Please select a Fieldman."
        dcboFieldman.SetFocus
        Exit Sub
    ElseIf dcboGrower.Text = "" Then
        MsgBox "Please select a Grower."
        dcboGrower.SetFocus
    Else
        Me.MousePointer = vbHourglass
        cmdNewRequest.Caption = "&New"
        adoRequest.Recordset.Update
        
        cmdDeleteRequest.Enabled = True
        cmdEditRequest.Enabled = True
        cmdCancelRequest.Enabled = False
        fraRequestItems.Enabled = True
        dcboFieldman.Locked = True
        dcboGrower.Locked = True
        mskRequestDate.Enabled = False
        cmdFirst.Enabled = True
        cmdPrevious.Enabled = True
        cmdNext.Enabled = True
        cmdLast.Enabled = True
        cmdPreviewRequest.Enabled = True
        cmdPrintRequest.Enabled = True
        cmdClose.Enabled = True
        
        adoRequest.Recordset.Resync
        'adoRequest.Refresh
        'adoRequest.Recordset.MoveLast
        Call qryRequestItems(Me.adoRequest.Recordset.Fields("Request_ID"))
        Me.MousePointer = vbDefault
        MsgBox "Record successfully updated."
        UpdateMode = False
    End If
End If

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub qryRequestItems(criteriaID As Long)
On Error Resume Next
adoRequestItems.CommandType = adCmdText
adoRequestItems.RecordSource = "select * from qryRequestItems where Rrequest_ID=" & criteriaID
adoRequestItems.Refresh
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdNewRequest.FontBold = False
cmdEditRequest.FontBold = False
cmdDeleteRequest.FontBold = False
cmdCancelRequest.FontBold = False
cmdPreviewRequest.FontBold = False
cmdPrintRequest.FontBold = False
cmdClose.FontBold = False
End Sub

Private Sub grdRequestItems_Error(ByVal DataError As Integer, Response As Integer)
Response = 0
End Sub

Private Sub grdRequestItems_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
Dim tempStock As Long
'MsgBox adoRequestItems.Recordset.Fields("Rrequest_ID")
If adoRequestItems.Recordset.EOF Then
    tempStock = 0
Else
    tempStock = adoRequestItems.Recordset.Fields("RequestDetail_ID")
End If
adoRequestDetail.CommandType = adCmdText
adoRequestDetail.RecordSource = "Select * from tblRequestDetail where RequestDetail_ID=" & tempStock
adoRequestDetail.Refresh
End Sub

Private Sub txtReqQuantity_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 110 Or KeyCode = 190 Then SendKeys "{Backspace}"
End Sub

Private Sub txtReqQuantity_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyDelete Then Exit Sub
If KeyAscii = vbKeyBack Then Exit Sub
If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
    KeyAscii = 0
    Exit Sub
End If
End Sub
