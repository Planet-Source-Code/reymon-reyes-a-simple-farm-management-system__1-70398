VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmItemUsage 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Usage"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10665
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   10665
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   720
      TabIndex        =   21
      Top             =   6600
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc adoUsageDel 
      Height          =   375
      Left            =   5160
      Top             =   0
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
      CommandType     =   2
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
      Connect         =   $"frmItemUsage.frx":0000
      OLEDBString     =   $"frmItemUsage.frx":0096
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "qryItemUsage"
      Caption         =   "adoUsageDel"
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
      Left            =   7440
      Top             =   1440
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
   Begin MSAdodcLib.Adodc adoClassification 
      Height          =   375
      Left            =   7440
      Top             =   1080
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
      Left            =   7440
      Top             =   720
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
      Left            =   7440
      Top             =   360
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
      Left            =   7440
      Top             =   0
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
      CommandType     =   2
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
      Connect         =   $"frmItemUsage.frx":012C
      OLEDBString     =   $"frmItemUsage.frx":01C2
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tblItemUsage"
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
      Left            =   7485
      TabIndex        =   13
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelUsage 
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
      Left            =   6165
      TabIndex        =   12
      Top             =   6600
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
      Left            =   4800
      TabIndex        =   11
      Top             =   6600
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
      Left            =   3525
      TabIndex        =   10
      Top             =   6600
      Width           =   1215
   End
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
      Left            =   2205
      TabIndex        =   9
      Top             =   6600
      Width           =   1215
   End
   Begin VB.TextBox txtUsed 
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
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtQuantity 
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
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1080
      Width           =   1335
   End
   Begin MSDataListLib.DataCombo dcboKind 
      Bindings        =   "frmItemUsage.frx":0258
      DataField       =   "StockTypeCategory_ID"
      DataSource      =   "adoUsage"
      Height          =   360
      Left            =   5160
      TabIndex        =   5
      Top             =   600
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
      Bindings        =   "frmItemUsage.frx":026E
      DataField       =   "StockType_ID"
      DataSource      =   "adoUsage"
      Height          =   360
      Left            =   1560
      TabIndex        =   4
      Top             =   1560
      Width           =   2415
      _ExtentX        =   4260
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
      Bindings        =   "frmItemUsage.frx":028E
      DataField       =   "Fieldman_ID"
      DataSource      =   "adoUsage"
      Height          =   360
      Left            =   1560
      TabIndex        =   3
      Top             =   1080
      Width           =   2415
      _ExtentX        =   4260
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
      Bindings        =   "frmItemUsage.frx":02A7
      DataField       =   "Grower_ID"
      DataSource      =   "adoUsage"
      Height          =   360
      Left            =   1560
      TabIndex        =   2
      Top             =   600
      Width           =   2415
      _ExtentX        =   4260
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
      Height          =   350
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
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
   Begin MSDataGridLib.DataGrid grdUsage 
      Bindings        =   "frmItemUsage.frx":02BF
      Height          =   4335
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   7646
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
         RecordSelectors =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   1170.142
         EndProperty
      EndProperty
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      DataField       =   "UsageID"
      DataSource      =   "adoUsageDel"
      Height          =   375
      Left            =   4200
      TabIndex        =   20
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Used:"
      DataSource      =   "adoUsage"
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
      Left            =   4530
      TabIndex        =   19
      Top             =   1680
      Width           =   525
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity:"
      DataSource      =   "adoUsage"
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
      Left            =   4200
      TabIndex        =   18
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label5 
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
      Height          =   360
      Left            =   480
      TabIndex        =   17
      Top             =   1200
      Width           =   960
   End
   Begin VB.Label Label4 
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
      Height          =   360
      Left            =   675
      TabIndex        =   16
      Top             =   720
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kind:"
      DataSource      =   "adoUsage"
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
      Left            =   4560
      TabIndex        =   15
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label2 
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
      Height          =   360
      Left            =   120
      TabIndex        =   14
      Top             =   1680
      Width           =   1320
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
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmItemUsage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LoadKind As Boolean
Public UpdateMode As Boolean

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

Private Sub cmdAddUsage_Click()
If cmdAddUsage.Caption = "&Add" Then
    UpdateMode = True
    cmdAddUsage.Caption = "&Save"
    adoUsage.Recordset.AddNew
    grdUsage.Enabled = False
    cmdEditUsage.Enabled = False
    cmdDeleteUsage.Enabled = False
    cmdCancelUsage.Enabled = True
    cmdClose.Enabled = False
    mskDate.Enabled = True
    mskDate.Text = Format(Date, "MM/dd/yyyy")
    dcboGrower.Locked = False
    dcboLeadman.Locked = False
    dcboClassification.Locked = False
    dcboKind.Locked = False
    txtQuantity.Locked = False
    txtUsed.Locked = False
    dcboGrower.Text = ""
    dcboLeadman.Text = ""
    dcboClassification.Text = ""
    dcboKind.Text = ""
    dcboGrower.SetFocus
    
ElseIf cmdAddUsage.Caption = "&Save" Then
    If mskDate.Text = "" Or Not IsDate(mskDate.Text) Then
        MsgBox "Invalid date. Please enter appropriate date."
        mskDate.SetFocus
        Exit Sub
    ElseIf dcboGrower.Text = "" Then
        MsgBox "Please select a grower."
        dcboGrower.SetFocus
        Exit Sub
    ElseIf dcboLeadman.Text = "" Then
        MsgBox "Please select a lead man."
        dcboLeadman.SetFocus
        Exit Sub
    ElseIf dcboClassification.Text = "" Then
        MsgBox "Please select a classification."
        dcboClassification.SetFocus
        Exit Sub
    ElseIf dcboKind.Text = "" Then
        MsgBox "Please select a kind."
        dcboKind.SetFocus
        Exit Sub
    ElseIf txtQuantity.Text = "" Or Not IsNumeric(txtQuantity.Text) Or IsNull(txtQuantity.Text) Then
        MsgBox "Invalid quantity value."
        txtQuantity.SetFocus
        Exit Sub
    ElseIf txtUsed.Text = "" Or Not IsNumeric(txtUsed.Text) Or IsNull(txtUsed.Text) Then
        MsgBox "Invalid quantity used value."
        txtUsed.SetFocus
        Exit Sub
    Else
        Me.MousePointer = vbHourglass
        cmdAddUsage.Caption = "&Add"
        adoUsage.Recordset.Update
        cmdEditUsage.Enabled = True
        cmdDeleteUsage.Enabled = True
        cmdCancelUsage.Enabled = False
        cmdClose.Enabled = True
        mskDate.Enabled = False
        dcboGrower.Locked = True
        dcboLeadman.Locked = True
        dcboClassification.Locked = True
        dcboKind.Locked = True
        txtQuantity.Locked = True
        txtUsed.Locked = True
        grdUsage.Enabled = True
        adoUsage.Recordset.Requery
        adoUsageDel.Refresh
        'adoUsage.Recordset.MoveLast
        'If grdUsage.Bookmark <> 0 Then
            
        'grdUsage.SelBookmarks.Add adoUsage.Recordset.Bookmark
        Me.MousePointer = vbDefault
        
        MsgBox "Record successfully updated."
        UpdateMode = False
    End If
End If
End Sub

Private Sub cmdCancelUsage_Click()
UpdateMode = False
adoUsage.Recordset.CancelUpdate
adoUsage.Refresh
cmdAddUsage.Caption = "&Add"
cmdEditUsage.Enabled = True
cmdDeleteUsage.Enabled = True
cmdCancelUsage.Enabled = False
cmdClose.Enabled = True
cmdEditUsage.Enabled = True
mskDate.Enabled = False
dcboGrower.Locked = True
dcboLeadman.Locked = True
dcboClassification.Locked = True
dcboKind.Locked = True
txtQuantity.Locked = True
txtUsed.Locked = True
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDeleteUsage_Click()
'On Error Resume Next

If MsgBox("The selected RECORD will be deleted, are you sure you want to continue?" & vbCr & "This action is irreversible.", vbOKCancel + vbExclamation, "Delete") = vbOK Then
  '  Me.MousePointer = vbHourglass
    'adoUsageDel.Recordset.Delete
    adoUsage.Recordset.Delete
    'adoUsageDel.Recordset.Delete
    adoUsage.Refresh
 '   Me.MousePointer = vbDefault
End If

Exit Sub
errorHandler:
        MsgBox Err.Description
End Sub

Private Sub cmdEditUsage_Click()
UpdateMode = True
cmdAddUsage.Caption = "&Save"
cmdEditUsage.Enabled = False
cmdDeleteUsage.Enabled = False
cmdCancelUsage.Enabled = True
cmdClose.Enabled = False
mskDate.Enabled = True
dcboGrower.Locked = False
dcboLeadman.Locked = False
dcboClassification.Locked = False
dcboKind.Locked = False
txtQuantity.Locked = False
txtUsed.Locked = False
cmdEditUsage.Enabled = False
End Sub

Private Sub Command1_Click()
adoUsageDel.Recordset.Bookmark = adoUsage.Recordset.Bookmark
End Sub

Private Sub dcboClassification_Change()

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
On Error GoTo ErrHandler

Call CenterForm(Me, MDIForm1)
LoadKind = False
Call ConnectDB(adoUsage)
adoUsage.CommandType = adCmdTable
adoUsage.RecordSource = "qryItemUsage"
adoUsage.Refresh

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

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub grdUsage_Error(ByVal DataError As Integer, Response As Integer)
'Response = 0
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

Private Sub txtUsed_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 110 Or KeyCode = 190 Then SendKeys "{Backspace}"
End Sub

Private Sub txtUsed_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyDelete Then Exit Sub
If KeyAscii = vbKeyBack Then Exit Sub
If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
    KeyAscii = 0
    Exit Sub
End If
End Sub
