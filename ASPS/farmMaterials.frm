VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9660
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   7170
   ScaleWidth      =   9660
   Begin VB.CommandButton cmdLast 
      Caption         =   ">l"
      Height          =   375
      Left            =   5280
      TabIndex        =   22
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">>"
      Height          =   375
      Left            =   3960
      TabIndex        =   21
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "<<"
      Height          =   375
      Left            =   2640
      TabIndex        =   20
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "l<"
      Height          =   375
      Left            =   1320
      TabIndex        =   19
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   4695
      Left            =   7920
      TabIndex        =   11
      Top             =   1560
      Width           =   1575
      Begin VB.CommandButton cmdClose 
         Caption         =   "CLO&SE"
         Height          =   495
         Left            =   240
         TabIndex        =   18
         Top             =   3960
         Width           =   1095
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&PRINT"
         Height          =   495
         Left            =   240
         TabIndex        =   17
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "PREVIE&W"
         Height          =   495
         Left            =   240
         TabIndex        =   16
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel4 
         Caption         =   "C&ANCEL"
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete3 
         Caption         =   "DELE&TE"
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&EDIT"
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&NEW"
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Request Details"
      Height          =   4695
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   7575
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&DELETE"
         Height          =   495
         Left            =   4440
         TabIndex        =   10
         Top             =   3960
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&CANCEL"
         Height          =   495
         Left            =   3000
         TabIndex        =   9
         Top             =   3960
         Width           =   1335
      End
      Begin VB.CommandButton cmdADD 
         Caption         =   "&ADD"
         Height          =   495
         Left            =   1560
         TabIndex        =   8
         Top             =   3960
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3375
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   7095
         _ExtentX        =   12515
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
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   1440
      TabIndex        =   5
      Top             =   840
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   240
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Leadman:"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Date:"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Request ID:"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
