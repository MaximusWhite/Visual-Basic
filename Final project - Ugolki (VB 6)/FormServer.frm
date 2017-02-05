VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormServer 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create a server"
   ClientHeight    =   3915
   ClientLeft      =   11040
   ClientTop       =   1830
   ClientWidth     =   6075
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   6075
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Play!"
      Enabled         =   0   'False
      Height          =   735
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   3540
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6175
            MinWidth        =   6175
            Text            =   "Waiting for a client..."
            TextSave        =   "Waiting for a client..."
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   480
   End
End
Attribute VB_Name = "FormServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  
    Call Form1.CreateField
    Call Form1.CreatePieces
    
    
  Form1.ApplicationState = "server"
  
  Randomize
  
  Dim i As Integer                     'firstly, server randomly calculates the color for itself
  
  i = Int((10 - 1 + 1) * Rnd() + 1)

  If i <= 5 Then
  
    Form1.Player = "white"
  
  Else
  
    Form1.Player = "black"
   
     
  End If



  FormServer.Hide
    
  Call Form1.HideButtons

If Form1.Player = "white" Then       ' then client gets a signal to start with the given color, opposite to what server calculated for itself

Form1.Server.SendData "system|start|black"

Form1.Label1.Caption = "Your move."

Else

Form1.Server.SendData "system|start|white"

Form1.Label1.Caption = "Opponent's move."

End If


Command1.Enabled = False

Status.Panels(1).Text = "Waiting for a client..."

End Sub

Private Sub Form_Load()

Form1.Server.Listen

Label1.Caption = "Your Ip-adress is: " + Form1.Server.LocalIP


End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Server.Close
End Sub
