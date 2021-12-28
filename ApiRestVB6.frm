VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSenha 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Text            =   "System@0671"
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox txtLogin 
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Text            =   "davi.ricardo"
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton Login 
      Caption         =   "Login"
      Height          =   615
      Left            =   1680
      TabIndex        =   0
      Top             =   2520
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'DECLARANDO VARIAVEIS
Private Type ResponseSuccess
    IdRegistro As String
    IdUsuario As String
    Email As String
    Status As String
    StatusCode As String
    Message As String
    IdUsuario_AD As String
End Type

Private Type ResponseError
    Code As String
    Message As String
End Type

Private Sub Form_Load()

End Sub


Private Sub Login_Click()
' CONEXAO
Dim strResponsse As String, url As String
Dim responseArray() As String

Dim http As Object
    Set http = CreateObject("WinHttp.WinHttprequest.5.1")
    url = "https://localhost:44358/api/v1/Authenticate/auth?login=" + txtLogin.Text + "&senha=" + txtSenha.Text
    http.Open "Get", url, False
    http.setRequestHeader "token", "b6a8210005dd868798395257d39f0b9001659bfa" 'APiVB6
    http.Send
    strResponsse = http.responseText
  ' CONEXÃO FIM
    
   ' CRIANDO OBJECTO
   
    Dim success As ResponseSuccess
    Dim error As ResponseError
    If InStr(1, strResponsse, ":200") > 0 Then
        responseArray = Split(strResponsse, ",")
        
        success.Email = Split(responseArray(2), ":")(1)
        success.IdRegistro = Split(responseArray(0), ":")(1)
        success.Usuario = Split(responseArray(1), ":")(1)
        success.Message = Split(responseArray(5), ":")(1)
        success.Status = Split(responseArray(3), ":")(1)
        success.StatusCode = Split(responseArray(4), ":")(1)
    Else
        responseArray = Split(strResponsse, ",")
        error.Code = Split(responseArray(0), ":")(1)
        error.Message = Split(responseArray(1), ":")(1)
    End If
    

End Sub


