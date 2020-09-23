Attribute VB_Name = "Module1"
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpFileName As String) As Integer
Public FSO As New FileSystemObject
Public MySql  As New ADODB.Connection
Public Access As New ADODB.Connection
Public MSDE   As New ADODB.Connection
Public Rs     As New ADODB.Recordset
Public Rs1    As New ADODB.Recordset

Public TranslateFile As New ADODB.Connection
Public TranslateRs   As New ADODB.Recordset
Public MessagesRs    As New ADODB.Recordset
Public DefaultLenguage As String

Public Parar  As Boolean
Function GetIni(section, key, Arq) As String
'section = É o que está entre []
'key = É o nome que se encontra antes do sinal de igual (=)
'arq = É o nome do arquivo INI
Dim val As String
Dim Valor As Integer
val = String$(255, 0)
Valor = GetPrivateProfileString(section, key, "", val, Len(val), Arq)
'If worked = 0 Then
'GetIni = ""
'Else
GetIni = Left(val, CStr(Valor))
'End If
End Function
Sub WriteIni(section, key, dado, Arq)
'section = É o que está entre []
'key = É o nome que se encontra antes do sinal de igual (=)
'dado = É o valor que vai depois do sinal de igual (=)
'arq = É o nome do arquivo INI
Dim val As String
Dim Valor As Integer
val = String$(255, 0)
Valor = WritePrivateProfileString(section, key, dado, Arq)
End Sub


Public Function ExtractArgument(ArgNum As Integer, srchstr As String, Delim As String) As String
    
    'Extract an argument or token from a str
    '     ing based on its position
    'and a delimiter.
    On Error GoTo Err_ExtractArgument
    Dim ArgCount As Integer
    Dim LastPos As Integer
    Dim Pos As Integer
    Dim Arg As String
    Arg = ""
    LastPos = 1
    If ArgNum = 1 Then Arg = srchstr


    Do While InStr(srchstr, Delim) > 0
        Pos = InStr(LastPos, srchstr, Delim)


        If Pos = 0 Then
            'No More Args found
            If ArgCount = ArgNum - 1 Then Arg = Mid(srchstr, LastPos)
            Exit Do
        Else
            ArgCount = ArgCount + 1


            If ArgCount = ArgNum Then
                Arg = Mid(srchstr, LastPos, Pos - LastPos)
                Exit Do
            End If
        End If
        LastPos = Pos + 1
    Loop
    '---------
    ExtractArgument = Arg
    Exit Function
Err_ExtractArgument:
    MsgBox "Error " & Err & ": " & Error
    Resume Next
End Function
Public Function FormatarForm(TheForm As Form)

For i% = 0 To TheForm.Controls.Count - 1
    If TypeOf TheForm.Controls(i%) Is Label Then
       TheForm.Controls(i%).Font = "Arial"
       TheForm.Controls(i%).FontBold = True
       TheForm.Controls(i%).ForeColor = vbBlack
'    ElseIf TypeOf TheForm.Controls(i%) Is CampoMoeda.Moeda Then
'       TheForm.Controls(i%).Font = "Arial"
'       TheForm.Controls(i%).Font.Bold = True
'       TheForm.Controls(i%).ForeColor = &H80000002
'    ElseIf TypeOf TheForm.Controls(i%) Is CampoData.Data Then
'       TheForm.Controls(i%).Font = "Arial"
'       TheForm.Controls(i%).Font.Bold = True
'       TheForm.Controls(i%).ForeColor = &H80000002
    ElseIf TypeOf TheForm.Controls(i%) Is TextBox Then
       TheForm.Controls(i%).Font = "Arial"
       TheForm.Controls(i%).Font.Bold = True
       TheForm.Controls(i%).ForeColor = &H80000002
    ElseIf TypeOf TheForm.Controls(i%) Is OptionButton Then
       TheForm.Controls(i%).FontName = "Arial"
       TheForm.Controls(i%).FontBold = True
       TheForm.Controls(i%).ForeColor = vbBlack
    ElseIf TypeOf TheForm.Controls(i%) Is CheckBox Then
       TheForm.Controls(i%).FontName = "Arial"
       TheForm.Controls(i%).FontBold = True
    ElseIf TypeOf TheForm.Controls(i%) Is CommandButton Then
       TheForm.Controls(i%).FontName = "Arial"
       TheForm.Controls(i%).FontBold = True
    ElseIf TypeOf TheForm.Controls(i%) Is Frame Then
       TheForm.Controls(i%).FontName = "Arial"
       TheForm.Controls(i%).FontBold = True
       TheForm.Controls(i%).ForeColor = &H80000002
    ElseIf TypeOf TheForm.Controls(i%) Is MSComctlLib.ListView Then
       TheForm.Controls(i%).ForeColor = &H80000002
       TheForm.Controls(i%).Font.Name = "Arial"
       TheForm.Controls(i%).Font.Bold = True
    ElseIf TypeOf TheForm.Controls(i%) Is ComboBox Then
       TheForm.Controls(i%).ForeColor = &H80000002
       TheForm.Controls(i%).FontName = "Arial"
       TheForm.Controls(i%).FontBold = True
    End If
Next i%
End Function

Public Function TranslateForm(TheForm As Form, Optional Lenguage As String)
If Lenguage = "" Then Lenguage = DefaultLenguage
Set TranslateRs = New ADODB.Recordset
If TranslateFile.State = adStateClosed Then
   TranslateFile.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MultiLang.mdb"
End If
TranslateRs.Open "select * from Controls where Form='" & TheForm.Name & "' and Control='(Itself)'", TranslateFile
TheForm.Caption = TranslateRs(CStr(Lenguage))
TranslateRs.Close

For Each thecontrol In TheForm.Controls
    TranslateRs.Open "select * from Controls where Form='" & TheForm.Name & "' and Control='" & thecontrol.Name & "'", TranslateFile
    If Not TranslateRs.EOF Then
       Do While Not TranslateRs.EOF
          CallByName thecontrol, TranslateRs("PropertyName"), VbLet, TranslateRs(CStr(Lenguage))
          TranslateRs.MoveNext
       Loop
    End If
    TranslateRs.Close
    
Next thecontrol
'TranslateFile.Close
End Function

Public Function GetMessage(MessageID As Double, Optional Lenguage As String) As String
If Lenguage = "" Then Lenguage = DefaultLenguage

MessagesRs.Open "select " & Lenguage & " from Messages where MessageID=" & MessageID, TranslateFile
GetMessage = MessagesRs(CStr(Lenguage))
MessagesRs.Close
End Function

