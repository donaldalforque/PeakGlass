VERSION 5.00
Begin VB.Form BASE_SetSelectionFrm 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5295
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbSet 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   600
      Width           =   5055
   End
   Begin VB.CommandButton btnAccept 
      Caption         =   "ENTER: Accept"
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
      Left            =   1920
      Picture         =   "BASE_SetSelectionFrm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "ESC: Cancel"
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
      Left            =   3600
      Picture         =   "BASE_SetSelectionFrm.frx":23D4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblCaption_Title 
      AutoSize        =   -1  'True
      Caption         =   "SELECT SET:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1500
   End
End
Attribute VB_Name = "BASE_SetSelectionFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAccept_Click()
    CurrentSetid = cmbSet.ItemData(cmbSet.ListIndex)
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            btnAccept_Click
        Case vbKeyEscape
            btnCancel_Click
    End Select
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Set rec = Global_Data("Sets")
    
    If Not rec.EOF Then
        cmbSet.Clear
        Do Until rec.EOF
            cmbSet.AddItem rec!SetName
            cmbSet.ItemData(cmbSet.NewIndex) = rec!SetId
            rec.MoveNext
        Loop
        cmbSet.ListIndex = 0
    End If
End Sub
