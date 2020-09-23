VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Form3 
   Caption         =   "Inventory"
   ClientHeight    =   5805
   ClientLeft      =   3210
   ClientTop       =   1380
   ClientWidth     =   8505
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   5805
   ScaleWidth      =   8505
   Begin VB.CommandButton cmdClose 
      Caption         =   "&CLOSE"
      Height          =   495
      Left            =   7080
      TabIndex        =   9
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdStocks 
      Caption         =   "&STOCKS"
      Height          =   495
      Left            =   5760
      TabIndex        =   8
      Top             =   5040
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3375
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   5953
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&SEARCH"
      Height          =   375
      Left            =   7200
      TabIndex        =   6
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   720
      Width           =   2055
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   315
      Left            =   2520
      TabIndex        =   3
      Top             =   720
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin VB.Label Label3 
      Caption         =   "Enter Text"
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Kind"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Classification"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
