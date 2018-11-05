VERSION 5.00
Begin VB.Form POS_CutSizeFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7710
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbMetric 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtWidth 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   4440
      TabIndex        =   1
      Text            =   "1"
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox txtLength 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   720
      TabIndex        =   0
      Text            =   "1"
      Top             =   600
      Width           =   1575
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
      Left            =   4200
      Picture         =   "POS_CutSizeFrm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
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
      Left            =   5880
      Picture         =   "POS_CutSizeFrm.frx":23D4
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   2760
      TabIndex        =   8
      Top             =   600
      Width           =   315
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "W:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   3480
      TabIndex        =   7
      Top             =   600
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "L:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CUT SIZE"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   120
      Picture         =   "POS_CutSizeFrm.frx":4763
      Stretch         =   -1  'True
      Top             =   120
      Width           =   7485
   End
End
Attribute VB_Name = "POS_CutSizeFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SizePrice, DeductQuantity As Double

Private Sub btnAccept_Click()
    SizePrice = 0
    DeductQuantity = 1
    
    'GET PRICE
    Dim con As New ADODB.Connection
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "POS_CutSizePrice_Get"
    cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , Val(POS_ItemSearchFrm.lvItemSearch.SelectedItem.SubItems(5)))
    cmd.Parameters.Append cmd.CreateParameter("@Length", adDecimal, adParamInput, , NVAL(txtLength.text))
                          cmd.Parameters("@Length").Precision = 18
                          cmd.Parameters("@Length").NumericScale = 2
    cmd.Parameters.Append cmd.CreateParameter("@Width", adDecimal, adParamInput, , NVAL(txtWidth.text))
                          cmd.Parameters("@Width").Precision = 18
                          cmd.Parameters("@Width").NumericScale = 2
    cmd.Parameters.Append cmd.CreateParameter("@MetricId", adInteger, adParamInput, , cmbMetric.ItemData(cmbMetric.ListIndex))
    Set rec = cmd.Execute
    
    If Not rec.EOF Then
        SizePrice = rec!price
        DeductQuantity = rec!DeductQuantity
    Else
        SizePrice = 0
        DeductQuantity = 1
    End If
    con.Close
    
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    SizePrice = 0
    TrackClipping = False
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
    Set rec = Global_Data("MetricUnit")
    
    cmbMetric.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            cmbMetric.AddItem rec!Abb
            cmbMetric.ItemData(cmbMetric.NewIndex) = rec!MetricId
            rec.MoveNext
        Loop
    End If
    cmbMetric.text = "in."
End Sub

Private Sub txtLength_Change()
    If IsNumeric(txtLength.text) = False Then
        txtLength.text = "1"
        selectText txtLength
    Else
        If NVAL(txtLength.text) <= 0 Then
            txtLength.text = "1"
            selectText txtLength
        End If
    End If
End Sub

Private Sub txtLength_GotFocus()
    selectText txtLength
End Sub

Private Sub txtWidth_Change()
    If IsNumeric(txtWidth.text) = False Then
        txtWidth.text = "1"
        selectText txtWidth
    Else
        If NVAL(txtWidth.text) <= 0 Then
            txtWidth.text = "1"
            selectText txtWidth
        End If
    End If
End Sub

Private Sub txtWidth_GotFocus()
        selectText txtWidth
End Sub
