VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lucy's Cuban Cafe"
   ClientHeight    =   5490
   ClientLeft      =   7380
   ClientTop       =   3840
   ClientWidth     =   7500
   BeginProperty Font 
      Name            =   "Arial Black"
      Size            =   11.25
      Charset         =   0
      Weight          =   900
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   7500
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   9240
      Top             =   600
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4680
      Top             =   3840
   End
   Begin VB.Frame Frame3 
      Caption         =   "Making Order"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7560
      TabIndex        =   26
      Top             =   960
      Width           =   2055
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   28
         Top             =   720
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Please wait..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   27
         Top             =   480
         Width           =   945
      End
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   3
      Left            =   4680
      Top             =   3000
   End
   Begin VB.Timer Timer3 
      Interval        =   3
      Left            =   5160
      Top             =   960
   End
   Begin VB.Frame Frame2 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5160
      TabIndex        =   21
      Top             =   0
      Width           =   2055
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "0 $"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   22
         Top             =   360
         Width           =   435
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   600
      Top             =   120
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2400
      ItemData        =   "Form1.frx":0000
      Left            =   120
      List            =   "Form1.frx":0002
      TabIndex        =   11
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton cmdOrd 
      Caption         =   "      "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   3360
      TabIndex        =   10
      ToolTipText     =   "Bread pudding"
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton cmdOrd 
      Caption         =   "      "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   1800
      TabIndex        =   9
      ToolTipText     =   "Baked custard"
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton cmdOrd 
      Caption         =   "      "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   8
      ToolTipText     =   "Coffee with milk"
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton cmdOrd 
      Caption         =   "      "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   3360
      TabIndex        =   7
      ToolTipText     =   "Cassava"
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton cmdOrd 
      Caption         =   "      "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   1800
      TabIndex        =   6
      ToolTipText     =   "Black beans and rice"
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton cmdOrd 
      Caption         =   "      "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   5
      ToolTipText     =   "Cuban bread with meat and cheese"
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton cmdOrd 
      Caption         =   "      "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   3360
      TabIndex        =   4
      ToolTipText     =   "Pork served with platains"
      Top             =   2760
      Width           =   375
   End
   Begin VB.CommandButton cmdOrd 
      Caption         =   "      "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   1800
      TabIndex        =   3
      ToolTipText     =   "Shredded beef"
      Top             =   2760
      Width           =   375
   End
   Begin VB.Timer An_Btn 
      Interval        =   1
      Left            =   240
      Top             =   4680
   End
   Begin VB.CommandButton cmdOrd 
      Caption         =   "      "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Chicken and yellow rice"
      Top             =   2760
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   10080
      Top             =   2040
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   7560
      TabIndex        =   0
      Top             =   1930
      Width           =   2415
      Begin VB.Label Lbl_Price 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   20
         Top             =   3000
         Width           =   480
      End
      Begin VB.Label Lbl_Price 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   19
         Top             =   2640
         Width           =   480
      End
      Begin VB.Label Lbl_Price 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   18
         Top             =   2280
         Width           =   480
      End
      Begin VB.Label Lbl_Price 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   17
         Top             =   1920
         Width           =   480
      End
      Begin VB.Label Lbl_Price 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   16
         Top             =   1560
         Width           =   480
      End
      Begin VB.Label Lbl_Price 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   15
         Top             =   1200
         Width           =   480
      End
      Begin VB.Label Lbl_Price 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   480
      End
      Begin VB.Label Lbl_Price 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Lbl_Price 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.CommandButton Ord_new 
      Caption         =   "New Order"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   5160
      TabIndex        =   24
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton Quit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   5160
      TabIndex        =   23
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton Ord_Fake 
      Caption         =   "Confirm Order"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   5160
      TabIndex        =   25
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Prices <<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6840
      TabIndex        =   1
      Top             =   2445
      Width           =   660
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim anim_open As Boolean, anim_close As Boolean, vt As Integer, vs As Integer, names(0 To 8) As String, prices(0 To 8) As Double
Dim total As Double, von As Integer, in_process As Boolean, ordering As Boolean, load As Integer, counter As Integer, vfc As Integer

''''''''''''''''''''''''''''''''WORKING WITH ARRAYS'''''''''''''''''''''''''


Private Sub create_names(arr() As String)       ' creating  names for items
   Dim i As Integer
    arr(0) = "Arroz con Pollo"
    arr(1) = "Ropa Vieja"
    arr(2) = "Masitas"
    arr(3) = "Cuban Sandwich"
    arr(4) = "Moros"
    arr(5) = "Yuca"
    arr(6) = "Cafe con Leche"
    arr(7) = "Flan"
    arr(8) = "Pudin de Pan"
     
   For i = 0 To 8
      
      cmdOrd(i).Caption = arr(i)
   
   Next i
     
End Sub
Private Sub create_prices(arr() As Double)      ' creating prices for items
    Randomize
    Dim i As Integer
    
    For i = 0 To 8
    
    arr(i) = FormatNumber((25 - 5 + 1) * Rnd() + 5, 2)
    
    Next i
            
End Sub



Private Sub fill_prices(arr() As String)                   ' filling the price list
Dim i As Integer
    For i = 0 To 8
    
    Lbl_Price(i).Caption = arr(i) + " --- " + Str(prices(i)) + "$"
    
    Next i
   
End Sub



'''''''''''''''''''''''''''''''''''''''''''FORM LOAD'''''''''''''''''''''''''''''''''


Private Sub Form_Load()

Frame2.Top = Frame2.Height * -1

Call create_names(names)

Call create_prices(prices)

Call fill_prices(names)

total = 0

anim_close = False
anim_open = True
  vt = 65
  vs = 100
  von = -65
  vfc = 70
  
End Sub


''''''''''''''''''''''''''''''''''PROCESSING BUTTONS'''''''''''''''''''''''''''''''''''''''

Private Sub cmdOrd_Click(Index As Integer)                 'CASE statement required for this exercise
          
         Select Case Index
          
          Case 0
          
          total = total + prices(0)
          List1.AddItem (names(0) + " --- " + Str(prices(0)) + " $")
          Label2.Caption = Str(total) + " $"
          
          Case 1
          total = total + prices(1)
          List1.AddItem (names(1) + " --- " + Str(prices(1)) + " $")
          Label2.Caption = Str(total) + " $"
          
          Case 2
          total = total + prices(2)
          List1.AddItem (names(2) + " --- " + Str(prices(2)) + " $")
          Label2.Caption = Str(total) + " $"
           
          Case 3
          total = total + prices(3)
          List1.AddItem (names(3) + " --- " + Str(prices(3)) + " $")
          Label2.Caption = Str(total) + " $"
           
          Case 4
          total = total + prices(4)
          List1.AddItem (names(4) + " --- " + Str(prices(4)) + " $")
          Label2.Caption = Str(total) + " $"
          
          Case 5
          total = total + prices(5)
          List1.AddItem (names(5) + " --- " + Str(prices(5)) + " $")
          Label2.Caption = Str(total) + " $"
          
          Case 6
          total = total + prices(6)
          List1.AddItem (names(6) + " --- " + Str(prices(6)) + " $")
          Label2.Caption = Str(total) + " $"
          
          Case 7
          total = total + prices(7)
          List1.AddItem (names(7) + " --- " + Str(prices(7)) + " $")
          Label2.Caption = Str(total) + " $"
            
          Case 8
          total = total + prices(8)
          List1.AddItem (names(8) + " --- " + Str(prices(8)) + " $")
          Label2.Caption = Str(total) + " $"
           
          
          
         End Select
          
          
End Sub


Private Sub Ord_Fake_Click()           'confirm order button
 Timer5.Enabled = True
End Sub

Private Sub Ord_new_Click()                   ' new order button
   
   If MsgBox("Do you really want to make a new order? Really, really?", vbYesNo, "Really?") = vbYes Then
    total = 0
    List1.Clear
    Label2.Caption = "0 $"
   End If
End Sub

Private Sub Quit_Click()               ' sweet and holy unload me

Unload Me

End Sub




'''''''''''''''''''''''''''''''''''''''''''ANIMATIONS''''''''''''''''''''''''''''''''''




Private Sub pretending()                        'animation of "ordering" (pretending)
       Label4.Visible = True
       counter = counter + 1
       If counter = 15 Then Label3.Caption = "Connecting server..."
       
       If counter = 25 Then Label3.Caption = "Authentication..."
       
       If counter = 32 Then
       
       MsgBox ("Authentication failed! Please try again later.")
       Timer6.Enabled = False
       counter = 0
       ordering = False
       vfc = -70
       Label3.Caption = "Please wait..."
       Timer5.Enabled = True
       
       
       
       
       End If
       
       
       
End Sub



Private Sub anim_confirm()                                 'make order animation
   Frame3.Left = Frame3.Left - vfc
   If Frame3.Left <= Frame2.Left Then
   Frame3.Left = Frame2.Left
   vfc = 0
   Timer5.Enabled = False
   ordering = True
   Timer6.Enabled = True
   End If
   
   If Frame3.Left >= Form1.Width Then
   Timer5.Enabled = False
   Frame3.Left = Form1.Width
   vfc = 70
   
   End If
   

End Sub

Private Sub loading()                                  'loading animation for making order
    

    If load = 0 Then Label4.Caption = "- - - - -"
    load = load + 1
    If load = 1 Then Label4.Caption = "- - = - -"
    If load = 2 Then Label4.Caption = "- = - = -"
    If load = 3 Then Label4.Caption = "= - - - ="
    If load = 4 Then
    Label4.Caption = "- - - - -"
    load = 0
    End If
    
End Sub



Private Sub anim_ordnew()                                    'new order button animation
      
      Ord_new.Top = Ord_new.Top + von
      If Ord_new.Top <= 1140 Then
        
        Timer4.Enabled = False
      
         von = von * -1
      End If

      
    
      If Ord_new.Top > 2760 Then
      
        Timer4.Enabled = False
         von = von * -1
      End If
   
End Sub


Private Sub anim_total()                                    'Total price label animation
     Dim vf As Integer
     vf = 30
     Frame2.Top = Frame2.Top + vf
     
     If Frame2.Top >= 0 Then
     
     Frame2.Top = 0
     Timer3.Enabled = False
     
     End If
     
End Sub


Private Sub anim_list()                                       ' listbox animation
 Dim vlw As Integer
    

    vlw = 150
    
    List1.Width = List1.Width + vlw
    
    If List1.Width >= 4815 Then
    
    List1.Width = 4815
    vlw = 0
      
    Timer2.Enabled = False
    
    End If
    
End Sub
Private Sub anim_buttons(ind As Integer)                       ' buttons animation
     
     Dim vbh As Integer, vbw As Integer
       vbh = 20
       vbw = 40
       
       cmdOrd(ind).Height = cmdOrd(ind).Height + vbh
       cmdOrd(ind).Width = cmdOrd(ind).Width + vbw
       If cmdOrd(ind).Height >= 795 Then
          cmdOrd(ind).Height = 795
          vbh = 0
       End If
       If cmdOrd(ind).Width >= 1455 Then
          cmdOrd(ind).Width = 1455
          vbw = 0
       End If
       
       If (vbh = 0) And (vbw = 0) Then
       
       An_Btn.Enabled = False
      
       End If
       
       
       
End Sub

Private Sub anim_op()                                        ' animation opening price list
   
      
      Label1.Top = Label1.Top - vt
      Frame1.Left = Frame1.Left - vs
  
  
  
  If Label1.Top <= 850 Then
  
  vt = 0
  Label1.Top = 850
  
  End If
  
  If Frame1.Left <= Form1.Width - Frame1.Width - 300 Then
  vs = 0
  Frame1.Left = Form1.Width - Frame1.Width - 300
  End If
  
  If (vt = 0) And (vs = 0) Then
     
     Timer1.Enabled = False
     Label1.Caption = "Prices >>"
     anim_open = False
     anim_close = True
     vt = 65
     vs = 90
     in_process = False
End If
End Sub
Private Sub anim_cl()                                         'animation closing price list

 
      Label1.Top = Label1.Top + vt
      Frame1.Left = Frame1.Left + vs
  
  
  
  If Label1.Top >= 2450 Then
  
  vt = 0
  Label1.Top = 2450
  
  End If
  
  If Frame1.Left >= Form1.Width Then
  vs = 0
  Frame1.Left = Form1.Width
  End If
  
  If (vt = 0) And (vs = 0) Then
     
     Timer1.Enabled = False
     Label1.Caption = "Prices <<"
     anim_open = True
     anim_close = False
     vt = 65
     vs = 90
     in_process = False
End If



 


End Sub


Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)   ' turning price list animation on
If ordering = False Then
Timer1.Enabled = True
If in_process = False Then Timer4.Enabled = True
in_process = True
End If

End Sub


''''''''''''''''''''''''''''''' TIMERS'''''''''''''''''''''''''''''''''



Private Sub Timer1_Timer()                              ' price list animation timer
 
 If anim_open = True Then Call anim_op
 If anim_close = True Then Call anim_cl

End Sub

Private Sub An_Btn_Timer()                              'buttons animation timer
Dim i As Integer
For i = 0 To 8
Call anim_buttons(i)
Next i
End Sub

Private Sub Timer2_Timer()                              ' listbox animation timer
 Call anim_list
End Sub

Private Sub Timer3_Timer()                              'Total price label animation timer
Call anim_total
End Sub

Private Sub Timer4_Timer()                                'new order button animation timer
Call anim_ordnew
End Sub

Private Sub Timer5_Timer()                                'making order order frame animation timer
If ordering = False Then Call anim_confirm
End Sub

Private Sub Timer6_Timer()                                'loading bar and weird connection phrases animation timer

Call loading
Call pretending
End Sub
