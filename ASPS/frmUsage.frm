VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUsage 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Usage"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10530
   ControlBox      =   0   'False
   Icon            =   "frmUsage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   10530
   Begin MSMask.MaskEdBox mskDate 
      DataField       =   "UDate"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "MM/dd/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      DataSource      =   "adoUsage"
      Height          =   375
      Left            =   1560
      TabIndex        =   21
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
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
   Begin VB.Frame frmCommands 
      BackColor       =   &H0000C000&
      Height          =   3255
      Left            =   9120
      TabIndex        =   20
      Top             =   2280
      Width           =   1455
      Begin VB.CommandButton cmdAddUsage 
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
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdEditUsage 
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
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdDeleteUsage 
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
         TabIndex        =   9
         Top             =   1440
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
         TabIndex        =   11
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
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
         TabIndex        =   10
         Top             =   2040
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid grdUsage 
      Bindings        =   "frmUsage.frx":000C
      Height          =   5655
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9975
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
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "UDate"
         Caption         =   "Date"
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
         DataField       =   "Grower"
         Caption         =   "Grower"
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
         DataField       =   "Name"
         Caption         =   "Leadman"
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
      BeginProperty Column04 
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
      BeginProperty Column05 
         DataField       =   "Qty"
         Caption         =   "Qty"
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
      BeginProperty Column06 
         DataField       =   "QtyUsed"
         Caption         =   "QtyUsed"
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
         DataField       =   "QtyBalance"
         Caption         =   "QtyBalance"
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
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   915.024
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoUsageItems 
      Height          =   375
      Left            =   9360
      Top             =   5760
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
      Caption         =   "adoUsageItems"
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
      Left            =   9360
      Top             =   7200
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
      Left            =   9360
      Top             =   6840
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
   Begin MSAdodcLib.Adodc adoLeadman 
      Height          =   375
      Left            =   9360
      Top             =   6480
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
      Caption         =   "adoLeadman"
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
   Begin MSAdodcLib.Adodc adoGrower 
      Height          =   375
      Left            =   9360
      Top             =   6120
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
   Begin MSAdodcLib.Adodc adoUsage 
      Height          =   375
      Left            =   9360
      Top             =   5400
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
      Caption         =   "adoUsage"
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
   Begin VB.TextBox txtQtyUsed 
      Alignment       =   1  'Right Justify
      DataField       =   "QtyUsed"
      DataSource      =   "adoUsage"
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
      Left            =   5640
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   6
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      DataField       =   "Qty"
      DataSource      =   "adoUsage"
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
      Left            =   5640
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   5
      Top             =   1320
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo dcboKind 
      Bindings        =   "frmUsage.frx":0028
      DataField       =   "StockTypeCategory_ID"
      DataSource      =   "adoUsage"
      Height          =   360
      Left            =   5640
      TabIndex        =   4
      Top             =   840
      Width           =   2775
      _ExtentX        =   4895
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
      Bindings        =   "frmUsage.frx":003E
      DataField       =   "StockType_ID"
      DataSource      =   "adoUsage"
      Height          =   360
      Left            =   1560
      TabIndex        =   3
      Top             =   1800
      Width           =   2775
      _ExtentX        =   4895
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
   Begin MSDataListLib.DataCombo dcboLeadman 
      Bindings        =   "frmUsage.frx":005E
      DataField       =   "Fieldman_ID"
      DataSource      =   "adoUsage"
      Height          =   360
      Left            =   1560
      TabIndex        =   2
      Top             =   1320
      Width           =   2775
      _ExtentX        =   4895
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
   Begin MSDataListLib.DataCombo dcboGrower 
      Bindings        =   "frmUsage.frx":0077
      DataField       =   "Grower_ID"
      DataSource      =   "adoUsage"
      Height          =   360
      Left            =   1560
      TabIndex        =   1
      Top             =   840
      Width           =   2775
      _ExtentX        =   4895
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
   Begin VB.TextBox txtDateUsage 
      DataField       =   "UDate"
      DataSource      =   "adoUsage"
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
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   615
      Index           =   1
      Left            =   -120
      Shape           =   4  'Rounded Rectangle
      Top             =   7680
      Width           =   9375
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   615
      Index           =   0
      Left            =   840
      Shape           =   4  'Rounded Rectangle
      Top             =   -360
      Width           =   8415
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   2295
      Left            =   9120
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Used:"
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
      Left            =   5010
      TabIndex        =   19
      Top             =   1920
      Width           =   525
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
      Left            =   4680
      TabIndex        =   18
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kind:"
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
      Left            =   5040
      TabIndex        =   17
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label4 
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
      Left            =   120
      TabIndex        =   16
      Top             =   1920
      Width           =   1320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Leadman:"
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
      TabIndex        =   15
      Top             =   1440
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grower:"
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
      Left            =   675
      TabIndex        =   14
      Top             =   960
      Width           =   765
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
      Left            =   945
      TabIndex        =   13
      Top             =   480
      Width           =   495
   End
End
Attribute VB_Name = "frmUsage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public UpdateMode As Boolean
Dim LoadKind As Boolean

Private Sub adoClassification_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub adoGrower_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub adoKind_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub adoleadman_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub adoUsage_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub adoUsageItems_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub cmdAddUsage_Click()
On Error GoTo ErrHandler

If cmdAddUsage.Caption = "&Add" Then
    UpdateMode = True
    cmdAddUsage.Caption = "&Save"
    grdUsage.Enabled = False
    adoUsage.Recordset.AddNew
    cmdEditUsage.Enabled = False
    cmdDeleteUsage.Enabled = False
    cmdClose.Enabled = False
    cmdCancel.Enabled = True
    'txtDateUsage.Locked = False
    dcboGrower.Locked = False
    dcboLeadman.Locked = False
    dcboClassification.Locked = False
    dcboKind.Locked = False
    dcboGrower.Text = ""
    dcboLeadman.Text = ""
    dcboClassification.Text = ""
    dcboKind.Text = ""
    txtQty.Locked = False
    txtQtyUsed.Locked = False
    'txtDateUsage.Text = Date
    mskDate.Text = Format(Date, "MM/dd/yyyy")
    dcboGrower.SetFocus
ElseIf cmdAddUsage.Caption = "&Save" Then
    If Not IsDate(mskDate.Text) Then
        MsgBox "Invalid Date."
        mskDate.SetFocus
        Exit Sub
    ElseIf dcboGrower.Text = "" Then
        MsgBox "Please select a Grower."
        dcboGrower.SetFocus
        Exit Sub
    ElseIf dcboLeadman.Text = "" Then
        MsgBox "Please select a Leadman."
        dcboLeadman.SetFocus
        Exit Sub
    ElseIf dcboClassification.Text = "" Then
        MsgBox "Please select a Classification."
        dcboClassification.SetFocus
        Exit Sub
    ElseIf dcboKind.Text = "" Then
        MsgBox "Please select a Kind."
        dcboKind.SetFocus
        Exit Sub
    ElseIf Not IsNumeric(txtQty.Text) Or txtQty.Text = "" Or IsNull(Val(txtQty.Text)) Then
        MsgBox "Invalid Quantity value."
        txtQty.SetFocus
        Exit Sub
    ElseIf Not IsNumeric(txtQtyUsed.Text) Or txtQtyUsed.Text = "" Or IsNull(Val(txtQtyUsed.Text)) Then
        MsgBox "Invalid Quantity Used value."
        txtQtyUsed.SetFocus
        Exit Sub
    ElseIf Val(txtQtyUsed.Text) > Val(txtQty.Text) Then
        MsgBox "Quantity used cannot be larger than the quantity."
        txtQtyUsed.SetFocus
        Exit Sub
    Else
        cmdAddUsage.Caption = "&Add"
        adoUsage.Recordset.Update
        adoUsage.Recordset.Resync
        cmdEditUsage.Enabled = True
        cmdDeleteUsage.Enabled = True
        cmdClose.Enabled = True
        cmdCancel.Enabled = False
        'txtDateUsage.Locked = True
        dcboGrower.Locked = True
        dcboLeadman.Locked = True
        dcboClassification.Locked = True
        dcboKind.Locked = True
        txtQty.Locked = True
        txtQtyUsed.Locked = True
        grdUsage.Enabled = True
        
        adoUsageItems.Refresh
        UpdateMode = False
        MsgBox "Record successfully updated."
        adoKind.CommandType = adCmdTable
        adoKind.RecordSource = "tblstocktypecategory"
        adoKind.Refresh
        Set dcboKind.RowSource = adoKind.Recordset
        dcboKind.ListField = "Category"
        dcboKind.BoundColumn = "StockTypeCategory_ID"
        Set dcboKind.DataSource = adoUsage.Recordset
        dcboKind.DataField = "StockTypeCategory_ID"
        dcboKind.Refresh
        adoUsageItems.Refresh
    End If
End If
    
Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdAddUsage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAddUsage.FontBold = True
cmdEditUsage.FontBold = False
cmdDeleteUsage.FontBold = False
cmdCancel.FontBold = False
cmdClose.FontBold = False
End Sub

Private Sub cmdCancel_Click()
On Error GoTo ErrHandler

UpdateMode = False
cmdAddUsage.Caption = "&Add"
adoUsage.Recordset.Cancel
adoUsage.Refresh
adoUsageItems.Refresh

adoGrower.CommandType = adCmdTable
adoGrower.RecordSource = "tblGrower"
adoGrower.Refresh
Set dcboGrower.RowSource = adoGrower.Recordset
dcboGrower.ListField = "Grower"
dcboGrower.BoundColumn = "Growers_ID"
Set dcboGrower.DataSource = adoUsage.Recordset
dcboGrower.DataField = "Grower_ID"
dcboGrower.Refresh

adoLeadman.CommandType = adCmdTable
adoLeadman.RecordSource = "FieldManNames"
adoLeadman.Refresh
Set dcboLeadman.RowSource = adoLeadman.Recordset
dcboLeadman.ListField = "Name"
dcboLeadman.BoundColumn = "Fieldman_ID"
Set dcboLeadman.DataSource = adoUsage.Recordset
dcboLeadman.DataField = "Fieldman_ID"
dcboLeadman.Refresh

adoClassification.CommandType = adCmdTable
adoClassification.RecordSource = "tblstocktype"
adoClassification.Refresh
Set dcboClassification.RowSource = adoClassification.Recordset
dcboClassification.ListField = "Type"
dcboClassification.BoundColumn = "StockType_ID"
Set dcboClassification.DataSource = adoUsage.Recordset
dcboClassification.DataField = "StockType_ID"
dcboClassification.Refresh

adoKind.CommandType = adCmdTable
adoKind.RecordSource = "tblstocktypecategory"
adoKind.Refresh
Set dcboKind.RowSource = adoKind.Recordset
dcboKind.ListField = "Category"
dcboKind.BoundColumn = "StockTypeCategory_ID"
Set dcboKind.DataSource = adoUsage.Recordset
dcboKind.DataField = "StockTypeCategory_ID"
dcboKind.Refresh

'txtDateUsage.Locked = True
dcboGrower.Locked = True
dcboLeadman.Locked = True
dcboClassification.Locked = True
dcboKind.Locked = True
txtQty.Locked = True
txtQtyUsed.Locked = True
cmdEditUsage.Enabled = True
cmdDeleteUsage.Enabled = True
cmdClose.Enabled = True
cmdCancel.Enabled = False
grdUsage.Enabled = True

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAddUsage.FontBold = False
cmdEditUsage.FontBold = False
cmdDeleteUsage.FontBold = False
cmdCancel.FontBold = True
cmdClose.FontBold = False
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAddUsage.FontBold = False
cmdEditUsage.FontBold = False
cmdDeleteUsage.FontBold = False
cmdCancel.FontBold = False
cmdClose.FontBold = True
End Sub

Private Sub cmdDeleteUsage_Click()
On Error GoTo ErrHandler

If MsgBox("The selected RECORD will be deleted, are you sure you want to continue?" & vbCr & "This action is irreversible.", vbOKCancel + vbExclamation, "Delete") = vbOK Then
    adoUsage.Recordset.Delete
    adoUsage.Recordset.Requery
    adoUsage.Refresh
    adoUsageItems.Refresh
End If

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdSaveUsage_Click()
adoUsage.Recordset.Update
adoUsage.Recordset.Resync
End Sub

Private Sub cmdDeleteUsage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAddUsage.FontBold = False
cmdEditUsage.FontBold = False
cmdDeleteUsage.FontBold = True
cmdCancel.FontBold = False
cmdClose.FontBold = False
End Sub

Private Sub cmdEditUsage_Click()
On Error GoTo ErrHandler
UpdateMode = True
grdUsage.Enabled = False
cmdAddUsage.Caption = "&Save"
'txtDateUsage.Locked = False
dcboGrower.Locked = False
dcboLeadman.Locked = False
dcboClassification.Locked = False
dcboKind.Locked = False
txtQty.Locked = False
txtQtyUsed.Locked = False
cmdEditUsage.Enabled = False
cmdDeleteUsage.Enabled = False
cmdCancel.Enabled = True
cmdClose.Enabled = False

adoKind.CommandType = adCmdText
adoKind.RecordSource = "select * from tblstocktypecategory where StockType_ID=" & _
    adoUsage.Recordset.Fields("StockType_ID")
adoKind.Refresh
Set dcboKind.RowSource = adoKind.Recordset
dcboKind.ListField = "Category"
dcboKind.BoundColumn = "StockTypeCategory_ID"
dcboKind.ReFill

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdEditUsage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAddUsage.FontBold = False
cmdEditUsage.FontBold = True
cmdDeleteUsage.FontBold = False
cmdCancel.FontBold = False
cmdClose.FontBold = False
End Sub

Private Sub dcboClassification_Change()
On Error Resume Next
If LoadKind = False Then Exit Sub
If dcboClassification.Text = "" Then Exit Sub
adoClassification.Recordset.Bookmark = dcboClassification.SelectedItem
adoKind.CommandType = adCmdText
adoKind.RecordSource = "select * from tblstocktypecategory where StockType_ID=" & _
    adoClassification.Recordset.Fields("StockType_ID")
adoKind.Refresh
Set dcboKind.RowSource = adoKind.Recordset
dcboKind.ListField = "Category"
dcboKind.Text = ""
dcboKind.Refresh
End Sub

Private Sub dcboClassification_GotFocus()
LoadKind = True
End Sub

Private Sub dcboClassification_LostFocus()
LoadKind = False
End Sub

Private Sub Form_Deactivate()
If Me.UpdateMode = True Then
    Me.ZOrder (vbBringToFront)
    Me.SetFocus
    Exit Sub
End If
End Sub

Private Sub Form_Load()
Call CenterForm(frmUsage, MDIForm1)
LoadKind = False

Call ConnectDB(adoUsage)
adoUsage.CommandType = adCmdTable
adoUsage.RecordSource = "tblItemUsage"
adoUsage.Refresh

Call ConnectDB(adoUsageItems)
adoUsageItems.CommandType = adCmdTable
adoUsageItems.RecordSource = "qryItemUsage"
adoUsageItems.Refresh

Call ConnectDB(adoGrower)
adoGrower.CommandType = adCmdTable
adoGrower.RecordSource = "tblGrower"
adoGrower.Refresh
Set dcboGrower.RowSource = adoGrower.Recordset
dcboGrower.ListField = "Grower"
dcboGrower.BoundColumn = "Growers_ID"
Set dcboGrower.DataSource = adoUsage.Recordset
dcboGrower.DataField = "Grower_ID"
dcboGrower.Refresh

Call ConnectDB(adoLeadman)
adoLeadman.CommandType = adCmdTable
adoLeadman.RecordSource = "FieldManNames"
adoLeadman.Refresh
Set dcboLeadman.RowSource = adoLeadman.Recordset
dcboLeadman.ListField = "Name"
dcboLeadman.BoundColumn = "Fieldman_ID"
Set dcboLeadman.DataSource = adoUsage.Recordset
dcboLeadman.DataField = "Fieldman_ID"
dcboLeadman.Refresh

Call ConnectDB(adoClassification)
adoClassification.CommandType = adCmdTable
adoClassification.RecordSource = "tblstocktype"
adoClassification.Refresh
Set dcboClassification.RowSource = adoClassification.Recordset
dcboClassification.ListField = "Type"
dcboClassification.BoundColumn = "StockType_ID"
Set dcboClassification.DataSource = adoUsage.Recordset
dcboClassification.DataField = "StockType_ID"
dcboClassification.Refresh

Call ConnectDB(adoKind)
adoKind.CommandType = adCmdTable
adoKind.RecordSource = "tblstocktypecategory"
adoKind.Refresh
Set dcboKind.RowSource = adoKind.Recordset
dcboKind.ListField = "Category"
dcboKind.BoundColumn = "StockTypeCategory_ID"
Set dcboKind.DataSource = adoUsage.Recordset
dcboKind.DataField = "StockTypeCategory_ID"
dcboKind.Refresh

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAddUsage.FontBold = False
cmdEditUsage.FontBold = False
cmdDeleteUsage.FontBold = False
cmdCancel.FontBold = False
cmdClose.FontBold = False
End Sub

Private Sub grdUsage_Error(ByVal DataError As Integer, Response As Integer)
Response = 0
End Sub

Private Sub grdUsage_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
adoUsage.Recordset.Bookmark = adoUsageItems.Recordset.Bookmark
End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 110 Or KeyCode = 190 Then SendKeys "{Backspace}"
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyDelete Then Exit Sub
If KeyAscii = vbKeyBack Then Exit Sub
If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
    KeyAscii = 0
    Exit Sub
End If
End Sub

Private Sub txtQtyUsed_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 110 Or KeyCode = 190 Then SendKeys "{Backspace}"
End Sub

Private Sub txtQtyUsed_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyDelete Then Exit Sub
If KeyAscii = vbKeyBack Then Exit Sub
If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
    KeyAscii = 0
    Exit Sub
End If
End Sub
