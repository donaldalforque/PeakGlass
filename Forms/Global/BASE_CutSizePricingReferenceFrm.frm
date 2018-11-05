VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form BASE_CutSizePricingReferenceFrm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtWidth 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3480
      MaxLength       =   50
      TabIndex        =   4
      Top             =   5880
      Width           =   1575
   End
   Begin VB.ComboBox cmbName 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1800
      Width           =   4215
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   840
      MaxLength       =   50
      TabIndex        =   3
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CheckBox chkShow 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Show All"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4050
      TabIndex        =   0
      Top             =   2160
      Width           =   1000
   End
   Begin MSComctlLib.ListView lvUoms 
      Height          =   3255
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Id"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Length"
         Object.Width           =   6253
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Width"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.Toolbar tb_Standard 
      Height          =   330
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   582
      ButtonWidth     =   1588
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4680
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BASE_CutSizePricingReferenceFrm.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BASE_CutSizePricingReferenceFrm.frx":6862
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BASE_CutSizePricingReferenceFrm.frx":D0C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BASE_CutSizePricingReferenceFrm.frx":13926
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Width:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   2880
      TabIndex        =   10
      Top             =   5900
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Length:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   240
      TabIndex        =   9
      Top             =   5900
      Width           =   570
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "BASE_CutSizePricingReferenceFrm.frx":1A188
      Top             =   680
      Width           =   480
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cut size Pricing Reference"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Left            =   840
      TabIndex        =   7
      Top             =   720
      Width           =   2940
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "You can set up standard cut sizes pricing for products purchased with uneven and indefinite sizes."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   450
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   4815
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   5895
      Left            =   120
      Top             =   480
      Width           =   5055
   End
End
Attribute VB_Name = "BASE_CutSizePricingReferenceFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CutSizePricingReferenceId As Integer
Dim identifier As Integer
Public Sub Initialize()
    txtName.text = ""
    CutSizePricingReferenceId = 0
    txtName.SetFocus
End Sub
Public Sub Populate(ByVal data As String)
    Dim item As MSComctlLib.ListItem
    Select Case data
        Case "CutSizePricingReference"
            Dim con As New ADODB.Connection
            Set cmd = New ADODB.Command
            Set rec = New ADODB.Recordset
            
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "BASE_CutSizePricingReference_Get"
            cmd.Parameters.Append cmd.CreateParameter("@CutSizePricingId", adInteger, adParamInput, , cmbname.ItemData(cmbname.ListIndex))
            Set rec = cmd.Execute
            lvUoms.ListItems.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        Set item = lvUoms.ListItems.add(, , "")
                            item.SubItems(1) = rec!CutSizePricingReferenceId
                            item.SubItems(2) = rec!value
                            item.SubItems(3) = rec!width
                            item.Checked = True
                    End If
                    rec.MoveNext
                Loop
            End If
            con.Close
        Case "CutSizePricing"
            identifier = 0
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("CutSizePricing")
            cmbname.Clear
            cmbname.AddItem ""
            cmbname.ItemData(cmbname.NewIndex) = 0
            cmbname.ListIndex = 0
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        cmbname.AddItem rec!Name
                        cmbname.ItemData(cmbname.NewIndex) = rec!CutSizePricingId
                    End If
                    rec.MoveNext
                Loop
            End If
            identifier = 1
    End Select
End Sub

Private Sub chkShow_Click()
    Dim item As MSComctlLib.ListItem
    
    Dim con As New ADODB.Connection
    con.ConnectionString = ConnString
    con.Open
    
    Set rec = New ADODB.Recordset
    'Set rec = Global_Data("CutSizePricingReference")
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_CutSizePricingReference_Get"
    cmd.Parameters.Append cmd.CreateParameter("@CutSizePricingId", adInteger, adParamInput, , cmbname.ItemData(cmbname.ListIndex))
    cmd.ActiveConnection = con
    Set rec = cmd.Execute
    lvUoms.ListItems.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            If chkShow.value = 1 Then
                Set item = lvUoms.ListItems.add(, , "")
                    item.SubItems(1) = rec!CutSizePricingReferenceId
                    item.SubItems(2) = rec!value
                    item.SubItems(3) = rec!width
                If rec!isActive = "True" Then item.Checked = True
                lvUoms.ColumnHeaders(1).width = lvUoms.width * 0.06
                lvUoms.ColumnHeaders(3).width = lvUoms.width * 0.44
                lvUoms.ColumnHeaders(4).width = lvUoms.width * 0.44
            Else
                If rec!isActive = "True" Then
                    Set item = lvUoms.ListItems.add(, , "")
                        item.SubItems(1) = rec!CutSizePricingReferenceId
                        item.SubItems(2) = rec!value
                        item.SubItems(3) = rec!width
                    If rec!isActive = "True" Then item.Checked = True
                    lvUoms.ColumnHeaders(1).width = lvUoms.width * 0
                    lvUoms.ColumnHeaders(3).width = lvUoms.width * 0.47
                    lvUoms.ColumnHeaders(4).width = lvUoms.width * 0.47
                End If
            End If
            rec.MoveNext
        Loop
    End If
    con.Close
End Sub

Private Sub cmbName_Click()
    If identifier <> 0 Then
        Populate "CutSizePricingReference"
    End If
End Sub

Private Sub Form_Load()
    lvUoms.ColumnHeaders(1).width = lvUoms.width * 0
    lvUoms.ColumnHeaders(3).width = lvUoms.width * 0.47
    lvUoms.ColumnHeaders(4).width = lvUoms.width * 0.47
    'Populate "CutSizePricingReference"
    Populate "CutSizePricing"
End Sub


Private Sub lvUoms_ItemClick(ByVal item As MSComctlLib.ListItem)
    CutSizePricingReferenceId = item.SubItems(1)
    txtName.text = item.SubItems(2)
    txtWidth.text = item.SubItems(3)
    txtName.SetFocus
End Sub

Private Sub tb_Standard_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ErrorHandler:
    Select Case Button.Index
        Case 1 'NEW
            Initialize
        Case 2 'Save
            If cmbname.ListIndex = 0 Then
                MsgBox "Please select a cut size name.", vbExclamation
                Exit Sub
            End If
            
            Dim CutSizePricingId As Integer
            CutSizePricingId = cmbname.ItemData(cmbname.ListIndex)
            
            Dim item As MSComctlLib.ListItem
            Set con = New ADODB.Connection
            Set rec = New ADODB.Recordset
            con.ConnectionString = ConnString
            con.Open
            
            'Check for Deactivate/Activated Lists
            For Each item In lvUoms.ListItems
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "BASE_CutSizePricingReference_Update"
                cmd.Parameters.Append cmd.CreateParameter("@CutSizePricingReferenceId", adInteger, adParamInputOutput, , item.SubItems(1))
                cmd.Parameters.Append cmd.CreateParameter("@CutSizePricingId", adInteger, adParamInput, , CutSizePricingId)
                cmd.Parameters.Append cmd.CreateParameter("@Value", adVarChar, adParamInput, 250, item.SubItems(2))
                cmd.Parameters.Append cmd.CreateParameter("@Width", adVarChar, adParamInput, 250, item.SubItems(3))
                cmd.Parameters.Append cmd.CreateParameter("@isActive", adBoolean, adParamInput, , item.Checked)
                cmd.Execute
            Next
            
            If Trim(txtName.text) = "" Then
                Exit Sub
            End If
        
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.Parameters.Append cmd.CreateParameter("@CutSizePricingReferenceId", adInteger, adParamInputOutput, , CutSizePricingReferenceId)
            cmd.Parameters.Append cmd.CreateParameter("@CutSizePricingId", adInteger, adParamInput, , CutSizePricingId)
            cmd.Parameters.Append cmd.CreateParameter("@Value", adVarChar, adParamInput, 250, txtName.text)
            cmd.Parameters.Append cmd.CreateParameter("@Width", adVarChar, adParamInput, 250, txtWidth.text)
            If lvUoms.ListItems.Count > 0 Then
                cmd.Parameters.Append cmd.CreateParameter("@isActive", adBoolean, adParamInput, , lvUoms.SelectedItem.Checked)
            Else
                cmd.Parameters.Append cmd.CreateParameter("@isActive", adBoolean, adParamInput, , "True")
            End If
            
            If CutSizePricingReferenceId = 0 Then
                cmd.CommandText = "BASE_CutSizePricingReference_Insert"
                cmd.Execute
                CutSizePricingReferenceId = cmd.Parameters("@CutSizePricingReferenceId")
                Set item = lvUoms.ListItems.add(, , "")
                    item.SubItems(1) = CutSizePricingReferenceId
                    item.SubItems(2) = txtName.text
                    item.Checked = True
                    item.Selected = True
                    item.EnsureVisible
            Else
                cmd.CommandText = "BASE_CutSizePricingReference_Update"
                cmd.Execute
                For Each item In lvUoms.ListItems
                    If item.SubItems(1) = CutSizePricingReferenceId Then
                        item.SubItems(2) = txtName.text
                        item.SubItems(3) = txtWidth.text
                        item.Selected = True
                        item.EnsureVisible
                    End If
                Next
            End If
            con.Close
        Case 4 'DELETE
            Set con = New ADODB.Connection
            con.ConnectionString = ConnString
            con.Open
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.Parameters.Append cmd.CreateParameter("@CutSizePricingReferenceId", adInteger, adParamInput, , CutSizePricingReferenceId)
            cmd.CommandText = "BASE_CutSizePricingReference_Delete"
            cmd.Execute
            con.Close
            lvUoms.ListItems.Remove (lvUoms.SelectedItem.Index)
            Initialize
    End Select
    Exit Sub
'ErrorHandler:
'    If IsNumeric(Err.Description) = True Then
'        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(Val(Err.Description))
'    Else
'        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & Err.Description
'    End If
'    GLOBAL_MessageFrm.Show (1)
End Sub

Private Sub Text1_Change()
    
End Sub

Private Sub txtName_Change()
    If IsNumeric(txtName.text) = False Then
        txtName.text = ""
    End If
End Sub

Private Sub txtName_GotFocus()
    selectText txtName
End Sub
Private Sub txtWidth_Change()
    If IsNumeric(txtWidth.text) = False Then
        txtWidth.text = ""
    End If
End Sub
