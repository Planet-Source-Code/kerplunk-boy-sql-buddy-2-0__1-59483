VERSION 5.00
Begin VB.Form FrmLingua 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Língua"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbLinguas 
      Height          =   315
      Left            =   1380
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3540
      TabIndex        =   1
      Top             =   540
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3540
      TabIndex        =   0
      Top             =   60
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Língua:"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   300
      Width           =   1155
   End
End
Attribute VB_Name = "FrmLingua"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
WriteIni "Lenguage", "Default", cmbLinguas.Text, App.Path & "\SQL Tester.ini"
MsgBox GetMessage(8, cmbLinguas.Text), vbInformation
Unload Me
End Sub

Private Sub Form_Load()
TranslateForm Me
Dim tmpcn As New ADODB.Connection
Dim tmprs As New ADODB.Recordset
Set tmpcn = New ADODB.Connection
Set tmprs = New ADODB.Recordset

tmpcn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MultiLang.mdb"

tmprs.Open "select * from Controls", tmpcn
For Each campo In tmprs.Fields
    If campo.Name <> "Form" And campo.Name <> "Control" And campo.Name <> "PropertyName" Then
        cmbLinguas.AddItem campo.Name
    End If
Next campo
cmbLinguas.Text = DefaultLenguage
tmprs.Close
Set Rs = Nothing
tmpcn.Close
Set tmprs = Nothing
End Sub
