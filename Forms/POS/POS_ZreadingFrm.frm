VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form POS_ZreadingFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "POS_ZreadingFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCancel 
      Caption         =   "ESC:Cancel"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   240
      Picture         =   "POS_ZreadingFrm.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Width           =   3975
   End
   Begin VB.CommandButton btnPrint 
      Caption         =   "F1: Print Z-Reading"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Picture         =   "POS_ZreadingFrm.frx":239B
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1920
      Width           =   3975
   End
   Begin MSComCtl2.DTPicker dtDate 
      Height          =   450
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   794
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   95485953
      CurrentDate     =   42297
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATE"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   525
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Z Reading Report"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   840
      TabIndex        =   3
      Top             =   360
      Width           =   1950
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   240
      X2              =   4200
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "POS_ZreadingFrm.frx":45C3
      Top             =   240
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   3735
      Left            =   120
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "POS_ZreadingFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnPrint_Click()
    Dim crxApp As New CRAXDRT.Application
    Dim crxRpt As New CRAXDRT.Report
    
    Screen.MousePointer = vbHourglass
    Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\POS_ZReading.rpt")
    crxRpt.EnableParameterPrompting = False
    crxRpt.DiscardSavedData
    Call ResetRptDB(crxRpt)
    
    'crxRpt.ParameterFields.GetItemByName("ReportTitle").AddCurrentValue txtTitle.text
    'crxRpt.ParameterFields.GetItemByName("DateFrom").AddCurrentValue DateFrom.value & " " & TimeFrom.value
    'crxRpt.ParameterFields.GetItemByName("DateTo").AddCurrentValue DateTo.value & " " & TimeTo.value
    
    crxRpt.ParameterFields.GetItemByName("@Date").AddCurrentValue dtDate.value
    crxRpt.ParameterFields.GetItemByName("@UserId").AddCurrentValue UserId
    crxRpt.ParameterFields.GetItemByName("@WorkstationId").AddCurrentValue WorkstationId
    
    crxRpt.PrintOut False
    Screen.MousePointer = vbDefault
    
    'POS Audit Trail
    SavePOSAuditTrail UserId, WorkstationId, 0, "Generate X-Reading Report"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            btnPrint_Click
        Case vbKeyEscape
            btnCancel_Click
    End Select
End Sub

Private Sub Form_Load()
    dtDate.value = Format(Now, "MM/DD/YY")
End Sub

