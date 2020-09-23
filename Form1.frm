VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Query builder & Tester"
   ClientHeight    =   8865
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   12435
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12435
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvwItens 
      Height          =   4575
      Left            =   60
      TabIndex        =   31
      Top             =   3900
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   8070
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdParar 
      Caption         =   "Parar"
      Height          =   495
      Left            =   6900
      Picture         =   "Form1.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   1260
      Width           =   615
   End
   Begin VB.Frame Frame2 
      Caption         =   "Query"
      Height          =   2055
      Left            =   0
      TabIndex        =   27
      Top             =   1680
      Width           =   7575
      Begin VB.TextBox TxtQuery 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   180
         Width           =   7455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Estrutura do banco"
      Height          =   3855
      Left            =   7620
      TabIndex        =   25
      Top             =   0
      Width           =   4695
      Begin MSComctlLib.TreeView trvEstruturaMysql 
         Height          =   3495
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   6165
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   6
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.TreeView trvEstruturaAccess 
         Height          =   3495
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   6165
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   6
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.TreeView trvEstruturaMSDE 
         Height          =   3495
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   6165
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   6
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   180
      Top             =   540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton cmdConectarMSDE 
      Height          =   315
      Left            =   6960
      Picture         =   "Form1.frx":0694
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Conectar ao MSDE(Tecla de atalho - F7)"
      Top             =   840
      Width           =   495
   End
   Begin VB.ListBox lstSenha 
      Height          =   255
      Left            =   9000
      TabIndex        =   21
      Top             =   6360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox lstUsuario 
      Height          =   255
      Left            =   9000
      TabIndex        =   20
      Top             =   6360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.CommandButton cmdConectarMySql 
      Height          =   315
      Left            =   6960
      Picture         =   "Form1.frx":0A1E
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Conectar ao MySql (Tecla de atalho - F6)"
      Top             =   540
      Width           =   495
   End
   Begin VB.CommandButton cmdConectarAccess 
      Height          =   315
      Left            =   6960
      Picture         =   "Form1.frx":0DA8
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Conectar com banco Access(Tecla de atalho - F5)"
      Top             =   240
      Width           =   495
   End
   Begin VB.Frame S 
      Caption         =   "Status"
      Height          =   1215
      Left            =   4860
      TabIndex        =   13
      Top             =   0
      Width           =   2655
      Begin VB.Label lblEstadoMSDE 
         Caption         =   "Desconectado"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   840
         TabIndex        =   23
         Top             =   900
         Width           =   1275
      End
      Begin VB.Label Label8 
         Caption         =   "MSDE:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   900
         Width           =   555
      End
      Begin VB.Label lblEstadoMySql 
         Caption         =   "Desconectado"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   840
         TabIndex        =   17
         Top             =   600
         Width           =   1275
      End
      Begin VB.Label lblEstadoAcess 
         Caption         =   "Desconectado"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   840
         TabIndex        =   16
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "MySql:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   795
      End
      Begin VB.Label Label6 
         Caption         =   "Access:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.TextBox txtSchema 
      Height          =   285
      Left            =   2160
      TabIndex        =   12
      Top             =   1320
      Width           =   2595
   End
   Begin VB.TextBox txtSenha 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   11
      Top             =   1020
      Width           =   2595
   End
   Begin VB.TextBox txtUsuario 
      Height          =   285
      Left            =   2160
      TabIndex        =   10
      Top             =   720
      Width           =   2595
   End
   Begin VB.ListBox lstServers 
      Height          =   255
      ItemData        =   "Form1.frx":1332
      Left            =   9060
      List            =   "Form1.frx":1334
      TabIndex        =   5
      Top             =   6360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Top             =   420
      Width           =   2595
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":1336
      Left            =   2160
      List            =   "Form1.frx":1343
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   60
      Width           =   2595
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   8490
      Width           =   12435
      _ExtentX        =   21934
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Schemata(somente MySql):"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   1995
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Senha:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1020
      Width           =   1995
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Usuário:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   1995
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Servidor/Arquivo:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   420
      Width           =   1995
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Executar query no banco:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1995
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "Arquivo"
      Begin VB.Menu mnuArquivoNova 
         Caption         =   "Nova Query"
      End
      Begin VB.Menu mnuArquivoAbrir 
         Caption         =   "Abrir Query"
      End
      Begin VB.Menu mnuArquivoSeparador1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArquivoSalvar 
         Caption         =   "Salvar"
      End
      Begin VB.Menu mnuArquivoSeparador2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArquivoSair 
         Caption         =   "Sair"
      End
   End
   Begin VB.Menu mnuOpcoes 
      Caption         =   "Opções"
      Begin VB.Menu mnuOpcoesConexoes 
         Caption         =   "Conexões"
      End
      Begin VB.Menu mnuOpcoesLingua 
         Caption         =   "Língua"
      End
   End
   Begin VB.Menu mnuSobre 
      Caption         =   "Sobre"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdConectarAccess_Click()
On Error GoTo erro
If Access.State = adStateOpen Then
   Access.Close
   lblEstadoAcess.Caption = GetMessage(2)
   lblEstadoAcess.ForeColor = vbRed
   Exit Sub
End If
frmAguarde.Show
Access.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & lstServers.List(0)
If Access.State = adStateOpen Then
   lblEstadoAcess.Caption = GetMessage(5)
   lblEstadoAcess.ForeColor = vbGreen
   trvEstruturaAccess.Nodes.Clear
   Set Rs = Access.OpenSchema(adSchemaTables)
   
   Dim tiposcampos(1 To 205)
   For i% = 1 To 205
       tiposcampos(i%) = GetMessage(6)
   Next i%
   tiposcampos(16) = "Integer"
   tiposcampos(2) = "Integer"
   tiposcampos(3) = "Integer"
   tiposcampos(20) = "Integer"
   tiposcampos(17) = "Integer"
   tiposcampos(18) = "Integer"
   tiposcampos(19) = "Integer"
   tiposcampos(21) = "Integer"
   tiposcampos(4) = "Single"
   tiposcampos(5) = "Double"
   tiposcampos(6) = "Currency"
   tiposcampos(14) = "Decimal"
   tiposcampos(131) = "Numeric"
   tiposcampos(11) = "Boolean"
   tiposcampos(10) = "Error"
   tiposcampos(72) = "GUID"
   tiposcampos(7) = "Date"
   tiposcampos(133) = "DBDate"
   tiposcampos(134) = "DBTime"
   tiposcampos(135) = "DBTimeStamp"
   tiposcampos(8) = "BSTR"
   tiposcampos(129) = "Char"
   tiposcampos(200) = "VarChar"
   tiposcampos(201) = "LongVarChar"
   tiposcampos(130) = "WChar"
   tiposcampos(202) = "VarWChar"
   tiposcampos(203) = "LongVarWChar"
   tiposcampos(128) = "Binary"
   tiposcampos(204) = "VarBinary"
   tiposcampos(205) = "LongVarBinary"

   
   
   
   
   Do Until Rs.EOF
      If Mid(Rs!table_name, 1, 4) <> "MSys" Then
           trvEstruturaAccess.Nodes.Add , , Rs!table_name, Rs!table_name
           Set Rs1 = New ADODB.Recordset
           Rs1.Open "select * from " & Rs!table_name, Access
           For Each campo In Rs1.Fields
              trvEstruturaAccess.Nodes.Add CStr(Rs!table_name), tvwChild, CStr(Rs!table_name) & "-" & CStr(campo.Name), CStr(campo.Name) & "(" & GetMessage(7) & ": " & tiposcampos(campo.Type) & "(" & campo.DefinedSize & "))"
           Next campo
           Rs1.Close
           Set Rs1 = Nothing
      End If
      Rs.MoveNext
    
   Loop
   Rs.Close
   Set Rs = Nothing

   
End If
Unload frmAguarde
Exit Sub
erro:
If Err.Number = -2147467259 Then
   If MsgBox("O seu banco de dados não pode ser encontrado no caminho especificado(" & lstServers.List(0) & "). Certifique-se que o arquivo encontra-se disponível e o computador de destino esteja ligado." + vbNewLine + "Deseja tentar conectar-se novamente?", vbYesNo + vbCritical, "Banco de dados não encontrado") = vbYes Then
      cmdConectarAccess_Click
   End If
Else
   MsgBox "Erro: " & Err.Number & vbNewLine & "Descrição: " & Err.Description
   Resume Next
End If
End Sub

Private Sub cmdConectarMSDE_Click()
On Error GoTo erro
If MSDE.State = adStateOpen Then
   MSDE.Close
   lblEstadoMSDE.Caption = "Desconectado"
   lblEstadoMSDE.ForeColor = vbRed
   Exit Sub
End If
MSDE.Open "Provider=SQLOLEDB;Data Source=" & lstServers.List(2) & ";User ID=" & lstUsuario.List(2) & ";Password=" & lstSenha.List(2) & ";Initial Catalog=" & txtSchema.Text
frmAguarde.Show 1
If MSDE.State = adStateOpen Then
   lblEstadoMSDE.Caption = "Conectado"
   lblEstadoMSDE.ForeColor = vbGreen
   trvEstruturaMSDE.Nodes.Clear
      Dim tiposcampos(1 To 205)
   For i% = 1 To 205
       tiposcampos(i%) = "Desconhecido"
   Next i%
   tiposcampos(16) = "Integer"
   tiposcampos(2) = "Integer"
   tiposcampos(3) = "Integer"
   tiposcampos(20) = "Integer"
   tiposcampos(17) = "Integer"
   tiposcampos(18) = "Integer"
   tiposcampos(19) = "Integer"
   tiposcampos(21) = "Integer"
   tiposcampos(4) = "Single"
   tiposcampos(5) = "Double"
   tiposcampos(6) = "Currency"
   tiposcampos(14) = "Decimal"
   tiposcampos(131) = "Numeric"
   tiposcampos(11) = "Boolean"
   tiposcampos(10) = "Error"
   tiposcampos(72) = "GUID"
   tiposcampos(7) = "Date"
   tiposcampos(133) = "DBDate"
   tiposcampos(134) = "DBTime"
   tiposcampos(135) = "DBTimeStamp"
   tiposcampos(8) = "BSTR"
   tiposcampos(129) = "Char"
   tiposcampos(200) = "VarChar"
   tiposcampos(201) = "LongVarChar"
   tiposcampos(130) = "WChar"
   tiposcampos(202) = "VarWChar"
   tiposcampos(203) = "LongVarWChar"
   tiposcampos(128) = "Binary"
   tiposcampos(204) = "VarBinary"
   tiposcampos(205) = "LongVarBinary"
   
   Set Rs = MSDE.OpenSchema(adSchemaTables)
   Do Until Rs.EOF
      If Mid(Rs!table_name, 1, 4) <> "MSys" Then
           trvEstruturaMSDE.Nodes.Add , , Rs!table_name, Rs!table_name
           Set Rs1 = New ADODB.Recordset
           On Error Resume Next
           Rs1.Open "select * from " & Rs!table_name, MSDE
           For Each campo In Rs1.Fields
              trvEstruturaMSDE.Nodes.Add CStr(Rs!table_name), tvwChild, CStr(Rs!table_name) & "-" & CStr(campo.Name), CStr(campo.Name) & "(Tipo: " & tiposcampos(campo.Type) & "(" & campo.DefinedSize & "))"
           Next campo
           Rs1.Close
           On Error GoTo 0
           Set Rs1 = Nothing
      End If
      Rs.MoveNext
    
   Loop
   Rs.Close
   Set Rs = Nothing

   
End If
Unload frmAguarde
Exit Sub
erro:
If Err.Number = -2147467259 Then
   MsgBox "Usuário/senha, servidor ou esquema inválido ou não encontrado.", vbOKOnly + vbCritical, "Erro"
Else
   MsgBox "Descrição: " & Err.Description & vbNewLine & "Número: " & Err.Number
End If

End Sub

Private Sub cmdConectarMySql_Click()
On Error GoTo erro
If MySql.State = adStateOpen Then
   MySql.Close
   lblEstadoMySql.Caption = GetMessage(2)
   lblEstadoMySql.ForeColor = vbRed
   Exit Sub
End If
frmAguarde.Show
MySql.Open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & lstServers.List(1) & ";PORT=3306;DATABASE=" & txtSchema.Text & ";USER=" & lstUsuario.List(1) & ";PASSWORD=" & lstSenha.List(1) & ";OPTION=3"
If MySql.State = adStateOpen Then
   lblEstadoMySql.Caption = GetMessage(5)
   lblEstadoMySql.ForeColor = vbGreen
   trvEstruturaMysql.Nodes.Clear
   Set Rs = MySql.OpenSchema(adSchemaTables)
   Do Until Rs.EOF
      If Mid(Rs!table_name, 1, 4) <> "MSys" Then
           trvEstruturaMysql.Nodes.Add , , Rs!table_name, Rs!table_name
           Set Rs1 = New ADODB.Recordset
           Rs1.Open "describe " & Rs!table_name, MySql
           Do While Not Rs1.EOF
              trvEstruturaMysql.Nodes.Add CStr(Rs!table_name), tvwChild, CStr(Rs!table_name) & "-" & CStr(Rs1!Field), CStr(Rs1!Field) & "(Tipo: " & CStr(Rs1!Type) & ")"
              Rs1.MoveNext
           Loop
           Rs1.Close
           Set Rs1 = Nothing
      End If
      Rs.MoveNext
    
   Loop
   Rs.Close
   Set Rs = Nothing
End If
Unload frmAguarde
Exit Sub
erro:
If Err.Number = -2147467259 Then
   MsgBox "Usuário/senha, servidor ou esquema inválido ou não encontrado.", vbOKOnly + vbCritical, "Erro"
Else
   MsgBox "Descrição: " & Err.Description & vbNewLine & "Número: " & Err.Number
End If
End Sub

Private Sub cmdParar_Click()
Parar = True
End Sub

Private Sub Combo1_Click()
txtServer.Text = lstServers.List(Combo1.ListIndex)
txtUsuario.Text = lstUsuario.List(Combo1.ListIndex)
txtSenha.Text = lstSenha.List(Combo1.ListIndex)
If Combo1.ListIndex = 0 Then
   trvEstruturaMysql.Visible = False
   trvEstruturaAccess.Visible = True
   trvEstruturaMSDE.Visible = False
ElseIf Combo1.ListIndex = 1 Then
   trvEstruturaMysql.Visible = True
   trvEstruturaAccess.Visible = False
   trvEstruturaMSDE.Visible = False
Else
   trvEstruturaMysql.Visible = False
   trvEstruturaAccess.Visible = False
   trvEstruturaMSDE.Visible = True

End If
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 116 Then
   trvEstruturaMysql.Visible = False
   trvEstruturaMSDE.Visible = False
   trvEstruturaAccess.Visible = False
   cmdConectarAccess_Click
   trvEstruturaAccess.Visible = True
ElseIf KeyCode = 117 Then
   trvEstruturaMysql.Visible = False
   trvEstruturaMSDE.Visible = False
   trvEstruturaAccess.Visible = False
   cmdConectarMySql_Click
   trvEstruturaMysql.Visible = True
ElseIf KeyCode = 118 Then
   trvEstruturaMysql.Visible = False
   trvEstruturaMSDE.Visible = False
   trvEstruturaAccess.Visible = False
   cmdConectarMSDE_Click
   trvEstruturaMysql.Visible = True
End If
End Sub

Private Sub Form_Load()
DefaultLenguage = GetIni("Lenguage", "Default", App.Path & "\SQL Tester.ini")
Set TranslateFile = New ADODB.Connection
Set MessageRs = New ADODB.Connection

TranslateForm Me ', "Portuguese"

Set MySql = New ADODB.Connection
Set Access = New ADODB.Connection
lstServers.AddItem GetIni("AMBIENTE", "CAMINHO_BANCO", FSO.GetSpecialFolder(1) & "\Newsystem.ini") & "\Nsys001.mdb"
lstServers.AddItem "localhost"
lstServers.AddItem "(Sem suporte)"
lstServers.AddItem "localhost"
txtSchema.Text = "mysql"
lstUsuario.AddItem "admin"
lstUsuario.AddItem "root"
lstUsuario.AddItem "admin"
lstSenha.AddItem ""
lstSenha.AddItem "root"
lstSenha.AddItem "sa"
Combo1.ListIndex = 0
FormatarForm Me
lblEstadoAcess.ForeColor = vbRed
lblEstadoMSDE.ForeColor = vbRed
lblEstadoMySql.ForeColor = vbRed

End Sub


Private Sub mnuArquivoAbrir_Click()
On Error GoTo erro
CommonDialog.ShowOpen
Dim texto As String
Open CommonDialog.FileName For Input As #1
Input #1, texto
Close #1
TxtQuery.Text = texto
Exit Sub
erro:
If Err.Number = 32755 Then
   Exit Sub
End If

End Sub

Private Sub mnuArquivoNova_Click()
lvwItens.ColumnHeaders.Clear
lvwItens.ListItems.Clear
TxtQuery.Text = ""

End Sub

Private Sub mnuArquivoSair_Click()
End
End Sub

Private Sub mnuArquivoSalvar_Click()
On Error GoTo erro
CommonDialog.ShowSave
Open CommonDialog.FileName For Output As #1
Print #1, TxtQuery.Text
Close #1
Exit Sub
erro:
If Err.Number = 32755 Then
   Exit Sub
End If
End Sub

Private Sub mnuOpcoesLingua_Click()
FrmLingua.Show
End Sub

Private Sub mnuSobre_Click()
frmAbout.Show 1
End Sub


Private Sub trvEstruturaAccess_DblClick()
If InStr(1, trvEstruturaAccess.SelectedItem.Text, "(") = 0 Then
   TxtQuery.Text = TxtQuery.Text & trvEstruturaAccess.SelectedItem.Text
Else
   TxtQuery.Text = TxtQuery.Text & ExtractArgument(1, trvEstruturaAccess.SelectedItem.Text, "(")
End If
End Sub

Private Sub trvEstruturaMSDE_DblClick()
If InStr(1, trvEstruturaMSDE.SelectedItem.Text, "(") = 0 Then
   TxtQuery.Text = TxtQuery.Text & trvEstruturaMSDE.SelectedItem.Text
Else
   TxtQuery.Text = TxtQuery.Text & ExtractArgument(1, trvEstruturaMSDE.SelectedItem.Text, "(")
End If
End Sub

Private Sub trvEstruturaMysql_DblClick()
If InStr(1, trvEstruturaMysql.SelectedItem.Text, "(") = 0 Then
   TxtQuery.Text = TxtQuery.Text & trvEstruturaMysql.SelectedItem.Text
Else
   TxtQuery.Text = TxtQuery.Text & ExtractArgument(1, trvEstruturaMysql.SelectedItem.Text, "(")
End If
End Sub

Private Sub TxtQuery_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
   Parar = False
   cmdParar.Enabled = True
   On Error GoTo sqlerro
   If Combo1.Text = "Access" Then
      Rs.Open TxtQuery.Text, Access
   ElseIf Combo1.Text = "MySql" Then
      Rs.Open TxtQuery.Text, MySql
   Else
      Rs.Open TxtQuery.Text, MSDE
   End If
   If Rs.State = 0 Then
      StatusBar1.SimpleText = GetMessage(3)
      Exit Sub
   End If
   cmdParar.Enabled = True
   lvwItens.Visible = False
   frmAguarde.Show
   
   lvwItens.ColumnHeaders.Clear
   lvwItens.ListItems.Clear
   For Each campo In Rs.Fields
       lvwItens.ColumnHeaders.Add , , campo.Name
   Next campo
   Do While Not Rs.EOF
      lvwItens.ListItems.Add , , Rs(0)
      For i% = 1 To Rs.Fields.Count - 1
          lvwItens.ListItems(lvwItens.ListItems.Count).ListSubItems.Add , , IIf(IsNull(Rs(i%)), "Null", Rs(i%))
      Next i%
      conta = conta + 1
      DoEvents
      If Parar Then Exit Do
      Rs.MoveNext
   Loop
   
   Parar = False
   cmdParar.Enabled = False
   lvwItens.Visible = True
   Rs.Close
   Unload frmAguarde
   StatusBar1.SimpleText = conta & " registro(s)."
   SendKeys "{BS}"
ElseIf KeyCode = vbKeyF4 Then
   If Combo1.ListIndex = 0 Then
      Combo1.ListIndex = 1
   ElseIf Combo1.ListIndex = 1 Then
      Combo1.ListIndex = 2
   Else
      Combo1.ListIndex = 0
   End If
   txtServer.Text = lstServers.List(Combo1.ListIndex)
   txtUsuario.Text = lstUsuario.List(Combo1.ListIndex)
   txtSenha.Text = lstSenha.List(Combo1.ListIndex)
   If Combo1.ListIndex = 0 Then
      trvEstruturaMysql.Visible = False
      trvEstruturaAccess.Visible = True
      trvEstruturaMSDE.Visible = False
   ElseIf Combo1.ListIndex = 1 Then
      trvEstruturaMysql.Visible = True
      trvEstruturaAccess.Visible = False
      trvEstruturaMSDE.Visible = False
   ElseIf Combo1.ListIndex = 2 Then
      trvEstruturaMysql.Visible = False
      trvEstruturaAccess.Visible = False
      trvEstruturaMSDE.Visible = True
   
   End If
End If

Exit Sub

sqlerro:
If Err.Number = 3709 Then
   MsgBox GetMessage(1), vbOKOnly + vbCritical, GetMessage(2)
   cmdParar.Enabled = False
ElseIf Err.Number = -2147217900 Then
   MsgBox "A instrução SQL que vc digitou contém um erro: " & vbNewLine & Err.Description, vbOKOnly + vbCritical, "Não conectado"
   cmdParar.Enabled = False
Else
   MsgBox "Descrição: " & Err.Description & vbNewLine & "Número: " & Err.Number, vbOKOnly + vbCritical, "Erro inesperado"
   cmdParar.Enabled = False
End If
Resume Next
End Sub

Private Sub txtSenha_Change()
lstSenha.List(Combo1.ListIndex) = txtSenha.Text
End Sub


Private Sub txtServer_Change()
lstServers.List(Combo1.ListIndex) = txtServer.Text
End Sub

Private Sub txtUsuario_Change()
lstUsuario.List(Combo1.ListIndex) = txtUsuario.Text
End Sub
