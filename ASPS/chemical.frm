VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Form2 
   Caption         =   "Stocks Details"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10305
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8085
   ScaleWidth      =   10305
   Begin VB.Frame Frame3 
      Height          =   4815
      Left            =   8760
      TabIndex        =   22
      Top             =   1920
      Width           =   1335
      Begin VB.CommandButton cmdClose 
         Caption         =   "&CLOSE"
         Height          =   495
         Left            =   120
         TabIndex        =   29
         Top             =   4080
         Width           =   1095
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&PRINT"
         Height          =   495
         Left            =   120
         TabIndex        =   28
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "PREVIE&W"
         Height          =   495
         Left            =   120
         TabIndex        =   27
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel2 
         Caption         =   "CANCE&L"
         Height          =   495
         Left            =   120
         TabIndex        =   26
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete2 
         Caption         =   "DELE&TE"
         Height          =   495
         Left            =   120
         TabIndex        =   25
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&EDIT"
         Height          =   495
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&NEW"
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2295
      Left            =   120
      TabIndex        =   8
      Top             =   4440
      Width           =   8415
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&CANCEL"
         Height          =   375
         Left            =   4320
         TabIndex        =   21
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&DELETE"
         Height          =   375
         Left            =   2880
         TabIndex        =   20
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&ADD"
         Height          =   375
         Left            =   1440
         TabIndex        =   19
         Top             =   1800
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo DataCombo4 
         Height          =   315
         Left            =   5280
         TabIndex        =   18
         Top             =   1200
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Text            =   "DataCombo4"
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Height          =   315
         Left            =   1320
         TabIndex        =   16
         Top             =   1200
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Text            =   "DataCombo3"
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1320
         TabIndex        =   14
         Top             =   720
         Width           =   6735
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   315
         Left            =   5280
         TabIndex        =   12
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   315
         Left            =   1320
         TabIndex        =   10
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label9 
         Caption         =   "Quantity:"
         Height          =   375
         Left            =   4440
         TabIndex        =   17
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Unit:"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Name:"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Kind:"
         Height          =   375
         Left            =   4440
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Classification:"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Details"
      Height          =   2295
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   8415
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1695
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   2990
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
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1200
      TabIndex        =   5
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Suppliers:"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Receipt No."
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Date:"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
