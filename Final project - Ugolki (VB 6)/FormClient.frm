VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormClient 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connect to server"
   ClientHeight    =   3900
   ClientLeft      =   10860
   ClientTop       =   6585
   ClientWidth     =   6555
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   6555
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   3525
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "Waiting for IP..."
            TextSave        =   "Waiting for IP..."
         EndProperty
      EndProperty
      MousePointer    =   1
   End
   Begin VB.CommandButton BtnConnect 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Connect"
      Height          =   615
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Enter IP-address"
      Height          =   195
      Left            =   1560
      TabIndex        =   4
      Top             =   600
      Width           =   1170
   End
End
Attribute VB_Name = "FormClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public counter As Integer

Private Sub BtnConnect_Click()
    
   Form1.Client.Close
   
   Form1.Client.Connect Text1.Text & "." & Text2.Text & "." & Text3.Text & "." & Text4.Text, 1111
 
   BtnConnect.Enabled = False
   
End Sub

Private Sub Form_Unload(Cancel As Integer)    ' closing the client window is considered as a change of mind, resulting in closing the connection
Form1.Client.Close
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)     ' when "." pressed, next window for IP-address is selected (to make input more comfortable)

If KeyAscii = Asc(".") Then

KeyAscii = 0

Text2.SetFocus

Text1.Text = Mid(Text1.Text, 1, Len(Text1.Text))

End If


End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)

If KeyAscii = Asc(".") Then

KeyAscii = 0

Text3.SetFocus

Text2.Text = Mid(Text2.Text, 1, Len(Text2.Text))

End If


End Sub


Private Sub Text3_KeyPress(KeyAscii As Integer)

If KeyAscii = Asc(".") Then

KeyAscii = 0

Text4.SetFocus

Text3.Text = Mid(Text3.Text, 1, Len(Text3.Text))

End If


End Sub


