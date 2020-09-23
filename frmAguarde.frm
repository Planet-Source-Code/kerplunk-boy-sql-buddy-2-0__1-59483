VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAguarde 
   BorderStyle     =   0  'None
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3975
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00400000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   60
      Top             =   900
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483643
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAguarde.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAguarde.frx":059A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAguarde.frx":0B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAguarde.frx":0ECE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   40
      Left            =   420
      Top             =   1200
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      FontTransparent =   0   'False
      Height          =   720
      Left            =   3120
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   1
      Top             =   300
      Width           =   720
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   300
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   0
      Top             =   300
      Width           =   720
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   1740
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   3
      Top             =   240
      Width           =   795
   End
   Begin VB.Line Line4 
      X1              =   60
      X2              =   3900
      Y1              =   1620
      Y2              =   1620
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   60
      X2              =   3900
      Y1              =   60
      Y2              =   60
   End
   Begin VB.Line Line2 
      X1              =   3900
      X2              =   3900
      Y1              =   60
      Y2              =   1620
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   60
      X2              =   60
      Y1              =   60
      Y2              =   1620
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Aguarde enquanto o comando é executado"
      Height          =   495
      Left            =   1140
      TabIndex        =   2
      Top             =   1140
      Width           =   1755
   End
End
Attribute VB_Name = "frmAguarde"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vai As Boolean

Private Sub Form_Load()
vai = False
Picture1.Picture = ImageList1.ListImages(4).ExtractIcon
Picture2.Picture = ImageList1.ListImages(3).ExtractIcon
End Sub

Private Sub Timer1_Timer()
If vai = True Then
   Picture3.Left = Picture3.Left + 90
   Picture3.Picture = ImageList1.ListImages(1).ExtractIcon
   If Picture3.Left >= 3000 Then
      vai = False
   End If
Else
   Picture3.Left = Picture3.Left - 90
   Picture3.Picture = ImageList1.ListImages(2).ExtractIcon
   If Picture3.Left <= 540 Then
      vai = True
   End If
End If
End Sub
