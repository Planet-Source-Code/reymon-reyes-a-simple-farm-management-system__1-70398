VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9615
   LinkTopic       =   "Form2"
   ScaleHeight     =   6465
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   1335
      Left            =   120
      TabIndex        =   34
      Top             =   5160
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   2355
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Height          =   1695
      Left            =   0
      TabIndex        =   23
      Top             =   3360
      Width           =   8895
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New Chemical"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   7320
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&NEXT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   960
         Width           =   2415
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "&PREVIOUS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   960
         Width           =   2655
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&FIND"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&DELETE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&CANCEL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&SAVE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&ADD"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Chemicals 
      BackColor       =   &H00C0C000&
      Caption         =   "Chemicals"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      Begin VB.TextBox txtItemDescription 
         BackColor       =   &H00E0E0E0&
         DataField       =   "Chemical_name"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox txtunitmeasure 
         BackColor       =   &H00E0E0E0&
         DataField       =   "Unit_measure"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2760
         Width           =   2415
      End
      Begin VB.TextBox txtunitcost 
         BackColor       =   &H00E0E0E0&
         DataField       =   "Unit_cost"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox txtquantity 
         BackColor       =   &H00E0E0E0&
         DataField       =   "Quantity"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox txtamount 
         BackColor       =   &H00E0E0E0&
         DataField       =   "Amount"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox Text21 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "Date"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox Text22 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "Time"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox Text23 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
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
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   600
         Width           =   2415
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E0E0E0&
         FillColor       =   &H00FFC0C0&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   8355
         TabIndex        =   3
         Top             =   1080
         Width           =   8415
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00E0E0E0&
         DataField       =   "Remarks"
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "chemicalform.frx":0000
         Left            =   6120
         List            =   "chemicalform.frx":000A
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   2760
         Width           =   2415
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00E0E0E0&
         DataField       =   "Transaction_number"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   1320
         Width           =   2415
      End
      Begin MSAdodcLib.Adodc Adodc4 
         Height          =   330
         Left            =   6360
         Top             =   4680
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
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
         Connect         =   $"chemicalform.frx":002B
         OLEDBString     =   $"chemicalform.frx":00E5
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select*qrychem"
         Caption         =   "Adodc4"
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
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   330
         Left            =   3960
         Top             =   4560
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
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
         Connect         =   $"chemicalform.frx":019F
         OLEDBString     =   $"chemicalform.frx":0236
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select* from tblItemChemicals"
         Caption         =   "Adodc3"
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
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   375
         Left            =   2040
         Top             =   4560
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
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
         Connect         =   $"chemicalform.frx":02CD
         OLEDBString     =   $"chemicalform.frx":0387
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "tblItemChemicals"
         Caption         =   "Adodc2"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "chemicalform.frx":0441
         DataField       =   "Item_chemical_code"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1800
         TabIndex        =   8
         Top             =   1800
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         BackColor       =   14737632
         ListField       =   "Item_chemical_code"
         BoundColumn     =   "Item_code"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   120
         Top             =   4560
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
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
         Connect         =   $"chemicalform.frx":0456
         OLEDBString     =   $"chemicalform.frx":0510
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "tblChemical"
         Caption         =   "Adodc1"
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
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   135
         Left            =   3000
         TabIndex        =   22
         Top             =   4320
         Width           =   15
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Item Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Item Descriiption"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Unit of Measure"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Unit Cost"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4440
         TabIndex        =   18
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   " Quantity"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   4440
         TabIndex        =   17
         Top             =   1320
         Width           =   1605
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4440
         TabIndex        =   16
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4440
         TabIndex        =   15
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Transacton No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   1575
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "chemicalform.frx":05CA
      Height          =   1815
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   3201
      _Version        =   393216
      BackColor       =   16777152
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
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
