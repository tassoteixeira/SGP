VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form cerradoBrowser 
   Caption         =   "Web Cerrado"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4740
      Top             =   180
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7140
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cerradoBrowser.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cerradoBrowser.frx":0452
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cerradoBrowser.frx":08A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cerradoBrowser.frx":1892
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cerradoBrowser.frx":1CE4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4635
      Left            =   60
      TabIndex        =   1
      Top             =   900
      Width           =   8955
      ExtentX         =   15796
      ExtentY         =   8176
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   1535
      ButtonWidth     =   1376
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Voltar"
            Object.ToolTipText     =   "Volta para a página anterior"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Proximo"
            Object.ToolTipText     =   "Avança para próxima página"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Parar"
            Object.ToolTipText     =   "Parar de carregar página"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Inicial"
            Object.ToolTipText     =   "Página Inicial (Cerrado Informática)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "GIC"
            Key             =   "GIC"
            Description     =   "Gerenciador de Integração Corporativa"
            Object.ToolTipText     =   "Gerenciador de Integração Corporativa"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "TeleCheq"
            Key             =   "TeleCheq"
            Description     =   "Consulta de Cheque TeleCheque"
            Object.ToolTipText     =   "Consulta de Cheque TeleCheque"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "cerradoBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lWebGic As String

Private Sub Form_Load()
    Dim xSite As String
    
    Screen.MousePointer = 1
    xSite = g_string
    g_string = ""
    If ReadINI("GRUPO DE EMPRESAS", "Nome do Grupo", gArquivoIni) = "TEIXEIRA E PINHEIRO LTDA" Then
        lWebGic = "http://192.168.1.6:8080/GIC"
    Else
        lWebGic = "http://tasso.myvnc.com:8080/GIC"
    End If

    WebBrowser1.Navigate xSite
End Sub
Private Sub Form_Resize()
    'se estado da tela for minimizado
    If Me.WindowState = 1 Then
    Else
        WebBrowser1.Top = 900
        WebBrowser1.Left = 60
        WebBrowser1.Width = Me.Width - 120
        WebBrowser1.Height = Me.Height - 720
    End If
End Sub
Private Sub Timer1_Timer()
   If WebBrowser1.Busy = False Then
      Timer1.Enabled = False
      Me.Caption = WebBrowser1.LocationName
   Else
      Me.Caption = "Abrindo a página ..."
   End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error GoTo FileError
    
    Timer1.Enabled = True
    Select Case Button.Key
      Case "Voltar"
         WebBrowser1.GoBack
      Case "Avancar"
         WebBrowser1.GoForward
      Case "Parar"
         Timer1.Enabled = False
         WebBrowser1.Stop
         Me.Caption = WebBrowser1.LocationName
'      Case "Atualizar"
'         WebBrowser1.Refresh
'      Case "Pesquisar"
'         WebBrowser1.GoSearch
      Case "Inicial"
         'WebBrowser1.GoHome
        WebBrowser1.Navigate "http://www.cerradoinformatica.com/cerradoinformatica"
      Case "GIC"
         'WebBrowser1.GoHome
        WebBrowser1.Navigate lWebGic
      Case "TeleCheq"
         'WebBrowser1.GoHome
        WebBrowser1.Navigate "http://www.telecheque.com.br"
    End Select
    Exit Sub
    
FileError:
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
  On Error Resume Next

  Me.Caption = WebBrowser1.LocationName

End Sub
