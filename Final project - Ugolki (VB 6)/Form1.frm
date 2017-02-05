VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UGOLKI"
   ClientHeight    =   9885
   ClientLeft      =   975
   ClientTop       =   960
   ClientWidth     =   12915
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   204
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9885
   ScaleWidth      =   12915
   Begin MSWinsockLib.Winsock Client 
      Left            =   1080
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Server 
      Left            =   360
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton BtnServer 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Create server"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3840
      Width           =   2535
   End
   Begin VB.CommandButton BtnClient 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Connect to server"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4920
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   360
      Top             =   360
   End
   Begin VB.CommandButton BtnPlay 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Play on one computer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   375
      Left            =   9600
      TabIndex        =   13
      Top             =   7920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   375
      Left            =   9600
      TabIndex        =   12
      Top             =   7320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "MultiMove indicator"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9000
      TabIndex        =   11
      Top             =   6600
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   8880
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   4920
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Press ESC to return to the main menu (nothing will be saved)"
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
      Left            =   480
      TabIndex        =   10
      Top             =   8880
      Visible         =   0   'False
      Width           =   8535
   End
   Begin VB.Label LblSystem 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
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
      Left            =   240
      TabIndex        =   9
      Top             =   9480
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Your color: "
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
      Left            =   9120
      TabIndex        =   8
      Top             =   3000
      Width           =   810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label3"
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
      Left            =   9120
      TabIndex        =   5
      Top             =   2280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
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
      Left            =   9120
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label LabelMultiMove 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "MultiMove = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9480
      TabIndex        =   3
      Top             =   8880
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label LabelMakingMove 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "MakingMove = "
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
      Left            =   9480
      TabIndex        =   2
      Top             =   8520
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ArrayOfCells(-1 To 10, -1 To 10) As TCell 'array of field cells is bigger in size, so that during the move checks there won't be any callings out of range

Private WhitePiece(1 To 9) As TGamePiece  'two separate arrays for white and black game pieces, which makes the win check be easier and lighter for hardware
Private BlackPiece(1 To 9) As TGamePiece

Public MakingMove As Boolean  ' this boolean tells if the player selected any game piece and is making move (used in several conditions to determine if called procedure should proceed or not)

Public MovingX As Integer, MovingY As Integer   ' these values represent the coordinates of the game piece which is selected and is going to be moved
Public TempMovingX As Integer, TempMovingY As Integer ' temp coordinates used for the sequence of cells during MultiMove

Public MovingWhiteIndex As Integer, MovingBlackIndex As Integer 'the indexes for the white and black pieces called during move

Public WhiteMoves As Boolean   ' indicates if the white player moves (otherwise it's black), which helps to establish proper checkings if playing through LAN

Public Player As String   ' indicates whose turn it is during current move

Public MultiMove As Boolean ' indicator, telling whether the current move is simple or MultiMove; makes moving easier as well as the whole work of the program lighter for the hardware

Public MultiRemove As Boolean  ' indicator showing if the MultiMove had been made, so it helps to avoid an extra work loading computer

Private MultiIndexX() As Integer, MultiIndexY() As Integer ' very important dynamic arrays necessary for redrawing selected cells after move was made

Public MultiNumber As Integer ' number of cells selected during MultiMove, used for redrawing after the move was made

Public BlackCounter As Integer, WhiteCounter As Integer ' counter of moves made
 
Public NetPlayer As String    ' used for network connection as an indicator of turn
                              ' (so if player presses on any piece, the program firstly checks if the piece's color is the same as the color given to player,
                              ' and then checks if this color is equal to NetPlayer indicator)

Public ApplicationState As String ' indicator of how the program should act as a whole ("server" or "client" ; "self" means that no connection exists and it's possible to play only on one computer)

Public FirstTime As Boolean    ' indicates if field and pieces were created before or not (used for refreshing the posisions of pieces and needed variables)


Public Sub ReturnToMain() ' hiding the field and refreshing the pieces positions
  
 Dim i As Integer, j As Integer, TempX, TempY As Integer, TempRow As Integer, TempColumn As Integer

If FirstTime = False Then
    For i = 1 To 8
    
       For j = 1 To 8
       
           
         ArrayOfCells(i, j).Cell.Visible = False
           
         ArrayOfCells(i, j).Taken = False
           
       Next j
       
        
    Next i
    
''''''''''''''' hiding white pieces ''''''''''
    
 
 TempY = 625
 TempX = 625
 TempRow = 1
 TempColumn = 1
 
 
 For i = 1 To 9
               
 With WhitePiece(i).Piece
 
     .Left = TempX
     .Top = TempY
     .Visible = False
 
 End With
 
 TempX = TempX + 1000
 
 With WhitePiece(i)
   
   .CurrentX = TempColumn
   .CurrentY = TempRow
 
 End With
   
   TempColumn = TempColumn + 1
   
 If (i = 3) Or (i = 6) Then
   
   TempX = 625
   TempY = TempY + 1000
   TempColumn = 1
   TempRow = TempRow + 1
   
 End If
 
 Next i
 
 
  For i = 1 To 3
 
    For j = 1 To 3
    
       ArrayOfCells(i, j).Taken = True
       ArrayOfCells(i, j).PossibleToMoveHere = False
      
    Next j
    
 Next i
 
 
    
 ''''''''''''''' hiding black pieces '''''''''''''

 TempY = ArrayOfCells(6, 6).Cell.Top + 25
 TempX = ArrayOfCells(6, 6).Cell.Left + 25
 TempRow = 6
 TempColumn = 6
 
 For i = 1 To 9
 
 With BlackPiece(i).Piece
 
     .Left = TempX
     .Top = TempY
     .Visible = False
 
 End With
 
 TempX = TempX + 1000
 
 With BlackPiece(i)
   
   .CurrentX = TempColumn
   .CurrentY = TempRow
 
 End With
   
   TempColumn = TempColumn + 1
   
 If (i = 3) Or (i = 6) Then
   
   TempX = ArrayOfCells(6, 6).Cell.Left + 25
   TempY = TempY + 1000
   TempColumn = 6
   TempRow = TempRow + 1
   
 End If
 
 Next i

 For i = 6 To 8
 
    For j = 6 To 8
    
       ArrayOfCells(i, j).Taken = True
       ArrayOfCells(i, j).PossibleToMoveHere = False
      
    Next j
    
 Next i


''''''''' showing menu buttons, hiding side-labels, setting everything by default'''''''''''''''

BtnPlay.Visible = True

BtnServer.Visible = True

BtnClient.Visible = True

Label1.Visible = False

Label2.Visible = False

Label3.Visible = False

Label4.Visible = False

BlackCounter = 0

WhiteCounter = 0

Label5.Visible = False

Form1.Width = 9405

If ApplicationState = "client" And Player <> "" Then Client.SendData "disconnect|"

If ApplicationState = "server" And Player <> "" Then Server.SendData "disconnect|"

Image1.Visible = False

Label6.Visible = False

ApplicationState = "self"

NetPlayer = "white"

End If

End Sub

Public Sub HideButtons()    'hiding buttons of the main menu
 
    BtnPlay.Visible = False
    BtnServer.Visible = False
    BtnClient.Visible = False
    
    Form1.SetFocus


End Sub

Public Sub CheckWin()   ' with each move this check is happening
                        'it doesn't load the program, which is useful
                        'it goes through all game pieces and checks if their positions are in "the winning zone"
                        'if the game piece is in its "winning zone", then counter for this type of pieces incriminates
                        'in the end, if the  counter for any of players is 9 (so that every piece is in its winning zone), the player wins
  
  Dim i As Integer, CountWhite As Integer, CountBlack As Integer


   For i = 1 To 9

      If ((BlackPiece(i).CurrentX >= 1) And (BlackPiece(i).CurrentX <= 3)) And ((BlackPiece(i).CurrentY >= 1) And (BlackPiece(i).CurrentY <= 3)) Then CountBlack = CountBlack + 1
       
   Next i
   
   
   
   For i = 1 To 9

      If ((WhitePiece(i).CurrentX >= 6) And (WhitePiece(i).CurrentX <= 8)) And ((WhitePiece(i).CurrentY >= 6) And (WhitePiece(i).CurrentY <= 8)) Then CountWhite = CountWhite + 1
       
   Next i
   
   If CountWhite = 9 Then
   
   MsgBox "WHITE WINS!!!"
     
    Player = "" ' after one of the players won, this statement "friezes" the game, doesn't allow it to complete selecting and moving game pieces
    
    Label1.Caption = " White player won!"
    
   End If
   If CountBlack = 9 Then
   
   MsgBox "BLACK WINS!!!"
   
    Player = ""   ' here same as  ^^^^^^^^^^^^^
    
    Label1.Caption = " Black player won!"
   
   End If

End Sub
Public Sub SelectPiece(X As Integer, Y As Integer)   'this sub changes the picture of selected piece

   ArrayOfCells(X, Y).Cell.Picture = LoadPicture("pictures/selected.gif")

End Sub
Public Sub ChangePlayer()   ' one of the most important procedures, changes the player who can move his (her) game pieces
                            ' so, basically it is a turn switcher

  If Player = "white" Then
   
     Player = "black"
     
     Label1.Caption = " Black player moves."
   
   Else
   
     Player = "white"
   
     Label1.Caption = "White player moves."
   
   End If
   
End Sub

Public Sub ChangeNetPlayer()   ' changes NetPlayer for network play


   If NetPlayer = "white" Then
   
      NetPlayer = "black"
     
     If Player = NetPlayer Then
        
        Label1.Caption = "Your move."
     
     Else
        
        Label1.Caption = "Opponents move."
     
     
     End If
   
   Else
   
     NetPlayer = "white"
   
     If Player = NetPlayer Then
        
        Label1.Caption = "Your move."
     
     Else
        
        Label1.Caption = "Opponents move."
     
     
     End If
   
   End If


End Sub


Public Sub GenerateFirstPlayer()   ' randomly picks a color making move first for games on one computer
  Randomize
  Dim i As Integer
  
  i = Int((10 - 1 + 1) * Rnd() + 1)

  If i <= 5 Then
  
    Player = "white"
    
    Label1.Caption = "White player moves."
  
  Else
  
    Player = "black"
   
    Label1.Caption = "Black player moves."
     
  End If
   
   
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub CreatePieces()    ' Creation of game pieces

 Dim i As Integer, j As Integer, TempX, TempY As Integer
 Dim TempName As String, TempRow As Integer, TempColumn As Integer
 
 If FirstTime = True Then
   
 ''''''''''''''   CREATING WHITE PIECES  ''''''''''''''''''''
 TempY = 625
 TempX = 625
 TempRow = 1
 TempColumn = 1
 
 For i = 1 To 9
 
 TempName = "WhitePiece" & i
 
 Set WhitePiece(i) = New TGamePiece   ' since I use self-made classes, it's necessary to initialize them
 
 Set WhitePiece(i).Piece = Form1.Controls.Add("vb.image", TempName)  ' one of the fields of the class TGamePiece is an Image object
                                                                     ' which is actually the only visible part of the whole class
                                                                     ' and which will be clicked
                                                                     ' to make it "real" it's necessary to add it as a control of Form1
               
 With WhitePiece(i).Piece
 
     .Left = TempX
     .Top = TempY
     .Width = 970
     .Height = 970
     .Stretch = True
     .Picture = LoadPicture("pictures\checker_white.gif")
     .ZOrder (0)
     .Visible = True
 
 End With
 
 TempX = TempX + 1000
 
 With WhitePiece(i)
   
   .Color = "white"
   .CurrentX = TempColumn
   .CurrentY = TempRow
   .Index = i
 
 End With
   
   TempColumn = TempColumn + 1
   
 If (i = 3) Or (i = 6) Then
   
   TempX = 625
   TempY = TempY + 1000
   TempColumn = 1
   TempRow = TempRow + 1
   
 End If
 
 Next i

 For i = 1 To 3
 
    For j = 1 To 3
    
       ArrayOfCells(i, j).Taken = True
       ArrayOfCells(i, j).PossibleToMoveHere = False
      
    Next j
    
 Next i
    
  ''''''''''''''''' CREATING BLACK PIECES '''''''''''''''''''''''''''''''
  
  
 TempY = ArrayOfCells(6, 6).Cell.Top + 25
 TempX = ArrayOfCells(6, 6).Cell.Left + 25
 TempRow = 6
 TempColumn = 6
 
 For i = 1 To 9
 
 TempName = "BlackPiece" & i
 
 Set BlackPiece(i) = New TGamePiece
 
 Set BlackPiece(i).Piece = Form1.Controls.Add("vb.image", TempName)
 
 With BlackPiece(i).Piece
 
     .Left = TempX
     .Top = TempY
     .Width = 970
     .Height = 970
     .Stretch = True
     .Picture = LoadPicture("pictures\checker_black.gif")
     .ZOrder (0)
     .Visible = True
 
 End With
 
 TempX = TempX + 1000
 
 With BlackPiece(i)
   
   .Color = "black"
   .CurrentX = TempColumn
   .CurrentY = TempRow
   .Index = i
 
 End With
   
   TempColumn = TempColumn + 1
   
 If (i = 3) Or (i = 6) Then
   
   TempX = ArrayOfCells(6, 6).Cell.Left + 25
   TempY = TempY + 1000
   TempColumn = 6
   TempRow = TempRow + 1
   
 End If
 
 Next i

 For i = 6 To 8
 
    For j = 6 To 8
    
       ArrayOfCells(i, j).Taken = True
       ArrayOfCells(i, j).PossibleToMoveHere = False
      
    Next j
    
 Next i
      
FirstTime = False

Else
     
    For i = 1 To 9
     
       WhitePiece(i).Piece.Visible = True
       BlackPiece(i).Piece.Visible = True
     
     Next i
   

End If

End Sub

Public Sub CreateField()     'this sub creates a game field consisting of the cells
                              'which are also objects of self-made class TFieldCell
  

Dim i As Integer, j As Integer, StupidName As String, TempXPlacement As Integer, TempYPlacement As Integer


Label1.Visible = True

Label2.Visible = True

Label3.Visible = True

Label5.Visible = True

Form1.Width = 12225

Image1.Visible = True

Label6.Visible = True

If FirstTime = False Then

   For i = 1 To 8
   
   
      For j = 1 To 8
      
         ArrayOfCells(i, j).Cell.Visible = True
      
      Next j
      
      
   Next i


   Exit Sub
   

End If
     
     
    TempYPlacement = 600
    For i = 1 To 8
       
    TempXPlacement = 600
   
       
       For j = 1 To 8
            
            StupidName = "FieldCell" & i & j
       
            Set ArrayOfCells(j, i) = New TCell
            
            Set ArrayOfCells(j, i).Cell = Me.Controls.Add("vb.image", StupidName)
              
            With ArrayOfCells(j, i).Cell
            
                   .Height = 1000
                   .Width = 1000
                   .Left = TempXPlacement
                   .Top = TempYPlacement
                   .Visible = True
                   .Stretch = True
                   .ZOrder (1)
                   
              If i Mod 2 <> 0 Then
              
                 If j Mod 2 <> 0 Then
                 
                       .Picture = LoadPicture("pictures/black2.gif")
                       ArrayOfCells(j, i).Color = "black"
                 
                 Else
                   
                       .Picture = LoadPicture("pictures/white2.gif")
                       ArrayOfCells(j, i).Color = "white"
                   
                 End If
                   
              Else
              
              
                    If j Mod 2 <> 0 Then
                 
                       .Picture = LoadPicture("pictures/white2.gif")
                       ArrayOfCells(j, i).Color = "white"
                 
                 Else
                   
                       .Picture = LoadPicture("pictures/black2.gif")
                        ArrayOfCells(j, i).Color = "black"
                 End If
                   
                    
                   
              End If
            
            
            End With
             
            TempXPlacement = TempXPlacement + 1000
            
    With ArrayOfCells(j, i)
    
        .PositionX = j
        .PositionY = i
    
    End With
            
       
       Next j
       
    TempYPlacement = TempYPlacement + 1000
         
       
    Next i
       

 For i = -1 To 0

     For j = -1 To 10
       
       Set ArrayOfCells(i, j) = New TCell
       
     Next j
 
 Next i
 
 For i = 9 To 10

     For j = -1 To 10
       
       Set ArrayOfCells(i, j) = New TCell
       
     Next j
 
 Next i

 For i = 1 To 8

       
      Set ArrayOfCells(i, -1) = New TCell
      
      Set ArrayOfCells(i, 0) = New TCell
      
        
      Set ArrayOfCells(i, 9) = New TCell
      
      Set ArrayOfCells(i, 10) = New TCell

 Next i

End Sub
Public Sub SimpleMove(indicator As String, PieceIndex As Integer, X As Integer, Y As Integer)  ' the logic for move, involving simple moving or a single "jump" over another piece
 
   
 If ArrayOfCells(X, Y).Taken = False Then
      
     If ((MovingX = X) And (((MovingY = Y - 1) Or (MovingY = Y + 1)) Or (((MovingY = Y + 2) And (ArrayOfCells(X, Y + 1).Taken = True)) Or ((MovingY = Y - 2) And (ArrayOfCells(X, Y - 1).Taken = True))))) Or ((MovingY = Y) And (((MovingX = X - 1) Or (MovingX = X + 1)) Or (((MovingX = X + 2) And (ArrayOfCells(X + 1, Y).Taken = True)) Or ((MovingX = X - 2) And (ArrayOfCells(X - 1, Y).Taken = True))))) Then
        
      If ApplicationState = "server" Or ApplicationState = "client" Then
            
          
      Else
        
        Call ChangePlayer
      
      End If
        
        
        
       
        Call MovePiece(indicator, PieceIndex, X, Y)
         
      If ApplicationState = "server" Then Server.SendData "move|" & indicator & "|" & PieceIndex & "|" & X & "|" & Y    ' sending the information about the move to the reciever
          
      If ApplicationState = "client" Then Client.SendData "move|" & indicator & "|" & PieceIndex & "|" & X & "|" & Y
        
         
       Else
         
       MsgBox ("Illegal move!")
         
         
     End If
    
   
 End If

End Sub

Public Sub MultiMoves(indicator As String, PieceIndex As Integer, X As Integer, Y As Integer, recieved As Boolean)    ' logic for a multimove.
                                                                                                                      ' if multimove mode is activated then with each selected sell this procedure does:
                                                                                                                      ' 1. check if it is possible to move there
                                                                                                                      ' 2. if yes, then mark it as possible to move (change a picture, send information to reciever if playing trough network); else => illegal move
                                                                                                                      ' 3. collects coordinates of possible to move cells which were picked up because after moving the cells should be redrawn into their initial color
                                                                                                                      
 If ArrayOfCells(X, Y).Taken = False Then
    
    If ((MovingX = X) And ((MovingY = Y - 1) Or (MovingY = Y + 1))) Or ((MovingY = Y) And ((MovingX = X - 1) Or (MovingX = X + 1))) Then
    
         Call MovePiece(indicator, PieceIndex, X, Y)
         
         MultiMove = False
         
          If ApplicationState = "server" Or ApplicationState = "client" Then
            
          
          Else
        
                 Call ChangePlayer
      
          End If
                
         MultiMove = False
         
         Image1.Picture = LoadPicture("pictures/multimovefalse.gif")
                
         Exit Sub
    
    End If
    
    
    
    If TempMovingX = X Then
        
       If ((TempMovingY = Y + 2) And (ArrayOfCells(X, Y + 1).Taken = True)) Or ((TempMovingY = Y - 2) And (ArrayOfCells(X, Y - 1).Taken = True)) Then
        
               
               ArrayOfCells(X, Y).Cell.Picture = LoadPicture("pictures/possible.gif")
               
               TempMovingX = X
               TempMovingY = Y
               MultiRemove = True
               
               
               MultiNumber = MultiNumber + 1
               
              ReDim Preserve MultiIndexX(1 To MultiNumber) As Integer
              ReDim Preserve MultiIndexY(1 To MultiNumber) As Integer
              
              MultiIndexX(UBound(MultiIndexX)) = X
              MultiIndexY(UBound(MultiIndexY)) = Y
               
               
           ElseIf Not ((TempMovingX = X) And (TempMovingY = Y)) Then
              
               
             If recieved = False Then MsgBox ("Illegal move!")
               
           
        End If
        
        
       
    
    ElseIf TempMovingY = Y Then
        
         If ((TempMovingX = X + 2) And (ArrayOfCells(X + 1, Y).Taken = True)) Or ((TempMovingX = X - 2) And (ArrayOfCells(X - 1, Y).Taken = True)) Then
        
          ArrayOfCells(X, Y).Cell.Picture = LoadPicture("pictures/possible.gif")
               
               TempMovingX = X
               TempMovingY = Y
               
               MultiRemove = True
                  
               MultiNumber = MultiNumber + 1
               
              ReDim Preserve MultiIndexX(1 To MultiNumber) As Integer
              ReDim Preserve MultiIndexY(1 To MultiNumber) As Integer
              
              MultiIndexX(UBound(MultiIndexX)) = X
              MultiIndexY(UBound(MultiIndexY)) = Y
                
    
            ElseIf Not ((TempMovingX = X) And (TempMovingY = Y)) Then
            
            If recieved = False Then MsgBox ("Illegal move!")
         
         
         End If
         
      Else
    
    If recieved = False Then MsgBox ("Illegal move!")
         
         
   End If
    
 
    
 End If


End Sub

Public Sub RemoveSelection(X As Integer, Y As Integer)    ' changes cell pictures to their originals after move is done
  
  Dim i As Integer
  
  ArrayOfCells(X, Y).Cell.Picture = LoadPicture("pictures/" & ArrayOfCells(X, Y).Color & "2.gif")
  
  
  
  If MultiRemove = True Then
  
       For i = 1 To MultiNumber
       
         ArrayOfCells(MultiIndexX(i), MultiIndexY(i)).Cell.Picture = LoadPicture("pictures/" & ArrayOfCells(MultiIndexX(i), MultiIndexY(i)).Color & "2.gif")
       
       Next i
       
    MultiRemove = False
       
  End If

End Sub
Public Sub MovePiece(indicator As String, PieceIndex As Integer, X2 As Integer, Y2 As Integer) ' procedure moving pieces to their new coordinates
   
   If indicator = "white" Then
     
     WhitePiece(PieceIndex).Piece.Top = ArrayOfCells(X2, Y2).Cell.Top + 25
     WhitePiece(PieceIndex).Piece.Left = ArrayOfCells(X2, Y2).Cell.Left + 25
     WhitePiece(PieceIndex).Selected = False
     WhitePiece(PieceIndex).CurrentX = X2
     WhitePiece(PieceIndex).CurrentY = Y2

     MakingMove = False
     
    Call RemoveSelection(MovingX, MovingY)
    
     ArrayOfCells(MovingX, MovingY).Taken = False
     ArrayOfCells(X2, Y2).Taken = True
     
    WhiteCounter = WhiteCounter + 1
     
   End If
   
   If indicator = "black" Then
     
     BlackPiece(PieceIndex).Piece.Top = ArrayOfCells(X2, Y2).Cell.Top + 25
     BlackPiece(PieceIndex).Piece.Left = ArrayOfCells(X2, Y2).Cell.Left + 25
     BlackPiece(PieceIndex).Selected = False
     BlackPiece(PieceIndex).CurrentX = X2
     BlackPiece(PieceIndex).CurrentY = Y2

     MakingMove = False
     
    Call RemoveSelection(MovingX, MovingY)
    
     ArrayOfCells(MovingX, MovingY).Taken = False
     ArrayOfCells(X2, Y2).Taken = True
     
    BlackCounter = BlackCounter + 1
     
   End If
   
If ApplicationState = "server" Or ApplicationState = "client" Then Call ChangeNetPlayer
      
      
   
   Call CheckWin

    
End Sub

Private Sub BtnClient_Click()   ' button calling client form
FormClient.Show
End Sub

Private Sub BtnPlay_Click()    ' button starting the game on one computer

Call CreateField
Call CreatePieces
Call GenerateFirstPlayer

Call HideButtons

Form1.SetFocus

End Sub

Private Sub BtnServer_Click()  ' button calling server form
FormServer.Show
Server.Close
Server.Listen
End Sub


Private Sub Client_Connect()    ' when client is connected to server it sends a message, that connection is successful

FormClient.Status.Panels(1).Text = "Connected to " + FormClient.Text1.Text & "." & FormClient.Text2.Text & "." & FormClient.Text3.Text & "." & FormClient.Text4.Text


Client.SendData "status|Client " + Client.LocalHostName + " connected!|connected"
  
FormClient.BtnConnect.Enabled = False
  
End Sub

Private Sub Client_DataArrival(ByVal bytesTotal As Long)   ' this is a part responsible for the analysis of incoming data on client side
    
Dim Data() As String
   Dim DataString As String
   
   Client.GetData DataString         ' firstly, client gets data and puts it into string
   
   LblSystem.Caption = DataString ' that's for debugging
   
   Data() = Split(DataString, "|")    ' then client takes unbounded array and fills it by splitting the data string apart and making bounds according to
                                      ' how many pieces separated by "|" where in it
   
   If Data(0) = "disconnect" Then     '  after that each element of data array represent a piece of information needed for developing the gameplay, messaging, connecting, disconnecting, etc.
   
      MsgBox "Your opponent has been disconnected."
      
      Player = ""
   
      Client.SendData "disconnectapproved|"
   
   End If
   
   If Data(0) = "disconnectapproved" Then Client.Close
   
   
   If Data(0) = "status" Then
   
         FormClient.Status.Panels(1).Text = Data(1)
   
   End If
   
   If Data(0) = "state" Then
   
         If Data(1) = "client" Then ApplicationState = "client"
   
   End If
   
   If Data(0) = "select" Then
      
      Call SelectPiece(Int(Data(1)), Int(Data(2)))
      
      MovingX = Int(Data(1))
      
      MovingY = Int(Data(2))
      
      TempMovingX = Int(Data(1))
      
      TempMovingY = Int(Data(2))
      
      If Data(3) = "white" Then
      
        MovingWhiteIndex = Int(Data(4))
        
      
      Else
        
        MovingBlackIndex = Int(Data(4))
        
      
      End If
      
   End If
   
    If Data(0) = "unselect" Then
      
      Call RemoveSelection(Int(Data(1)), Int(Data(2)))
   
    End If
   
   If Data(0) = "move" Then
   
   
      Call MovePiece(Data(1), Int(Data(2)), Int(Data(3)), Int(Data(4)))
    
   
   End If
   
  If Data(0) = "multimove" Then
   
    Call MultiMoves(Data(1), Int(Data(2)), Int(Data(3)), Int(Data(4)), True)
     
    MultiRemove = True
   
   End If
   
   If Data(0) = "multiremoveboolean" Then MultiRemove = True
   
   
   If Data(0) = "change" Then Call ChangeNetPlayer
   
   If Data(0) = "system" Then
        
       If Data(1) = "start" Then
   
               Call CreateField
               Call CreatePieces
               Call HideButtons
               ApplicationState = "client"
               FormClient.Hide
               FormClient.BtnConnect.Enabled = True
               FormClient.Status.Panels(1).Text = "Waiting for IP..."
               FormClient.Text1.Text = ""
               FormClient.Text2.Text = ""
               FormClient.Text3.Text = ""
               FormClient.Text4.Text = ""
               
          Player = Data(2)
               
          If Player = NetPlayer Then
          
            Label1.Caption = "Your move."
          
          Else
               
            Label1.Caption = "Opponent's move"
               
          End If
          
               
       End If
       
   End If
    
  
    
End Sub


Private Sub Client_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
' if client can't connect to the given IP, it will indicate an error

FormClient.Status.Panels(1).Text = "Error occured"
FormClient.BtnConnect.Enabled = True
Client.Close
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)    ' needed for multimove change and exiting the current game

'MsgBox KeyCode

If KeyCode = 68 Then
   
   If LblSystem.Visible = False Then
   
      LblSystem.Visible = True
   
   Else
   
      LblSystem.Visible = False
   
   End If

End If

If KeyCode = 27 Then Call ReturnToMain   ' if ESC is pressed => field, pieces, etc. get hidden; needed variables are set to default

If KeyCode = 32 Then     ' if SPACEBAR is pressed, then
                         ' if multimove was made => finish multimove, move the image of piece, refresh selected cells, send data to reciever, give the piece new coordinates, change multimove indicator picture
                         ' else => change multimove indicator


   If MultiMove = True Then
              
    If WhiteMoves = True Then
              
       If MakingMove = True Then
        
         If Not ((TempMovingX = MovingX) And (TempMovingY = MovingY)) Then
         
         If ApplicationState = "server" Or ApplicationState = "client" Then
        
         
         Else
            Call ChangePlayer
         End If
         
         
         
         Call MovePiece("white", MovingWhiteIndex, TempMovingX, TempMovingY)
        
        If ApplicationState = "server" Then Server.SendData "move|white|" & MovingWhiteIndex & "|" & TempMovingX & "|" & TempMovingY
         
        If ApplicationState = "client" Then Client.SendData "move|white|" & MovingWhiteIndex & "|" & TempMovingX & "|" & TempMovingY
         
         
         End If
         
       End If
         
         MultiMove = False
         
         Image1.Picture = LoadPicture("pictures/multimovefalse.gif")
   
    Else
        
       If MakingMove = True Then
        
         If Not ((TempMovingX = MovingX) And (TempMovingY = MovingY)) Then
       
         If ApplicationState = "server" Or ApplicationState = "client" Then
         
         
         Else
            Call ChangePlayer
         End If
         
         
         
         Call MovePiece("black", MovingBlackIndex, TempMovingX, TempMovingY)
        
        If ApplicationState = "server" Then Server.SendData "move|black|" & MovingBlackIndex & "|" & TempMovingX & "|" & TempMovingY
         
        If ApplicationState = "client" Then Client.SendData "move|black|" & MovingBlackIndex & "|" & TempMovingX & "|" & TempMovingY
        
        
         End If
         
       End If
         
         MultiMove = False
         
         Image1.Picture = LoadPicture("pictures/multimovefalse.gif")
   
    End If
   
   Else
      
      MultiMove = True
      
      Image1.Picture = LoadPicture("pictures/multimovetrue.gif")
      
      MultiNumber = 0
      
   ReDim MultiIndexX(1 To 1) As Integer
   ReDim MultiIndexY(1 To 1) As Integer
      
      
      
   End If


End If
End Sub

Private Sub Form_Load()

FirstTime = True

Form1.Width = 9405

Server.LocalPort = 1111

'MMControl1.FileName = "sound\ABitOfJazz.wav"
'MMControl1.Command = "open"
'MMControl1.Command = "play"

Image1.Picture = LoadPicture("pictures/multimovefalse.gif")

MultiMove = False

NetPlayer = "white"       ' for the network play, white color always moves the first

ApplicationState = "self"

End Sub

Private Sub Form_Unload(Cancel As Integer) ' if main window is closed => close other windows too
   
   Unload FormServer
   Unload FormClient
   
End Sub

Private Sub Server_Connect()
FormServer.Command1.Enabled = True
End Sub

Private Sub Server_ConnectionRequest(ByVal requestID As Long)  ' if server recieves connection request => accept it

Server.Close

Server.Accept requestID

End Sub

Private Sub Server_DataArrival(ByVal bytesTotal As Long)      ' this is a part responsible for the analysis of incoming data on server side
                                                              ' in terms of construction is the same as for client part; doesn't have to be explained
   
   Dim Data() As String
   Dim DataString As String
   
   Server.GetData DataString
   
   LblSystem.Caption = DataString
   
   Data() = Split(DataString, "|")
   
   If Data(0) = "disconnect" Then
   
      MsgBox "Your opponent has been disconnected."
      
      Player = ""
      
      Server.SendData "disconnectapproved|"
      
   End If
   
   If Data(0) = "disconnectapproved" Then
   
   Server.Close
   
   Server.Listen
   
   End If
   
   If Data(0) = "status" Then
   
         FormServer.Status.Panels(1).Text = Data(1)
         
         If Data(2) = "connected" Then
         
         FormServer.Command1.Enabled = True
            
         Server.SendData "status|Waiting for server to start the game..."
            
         End If
   
   End If
   
   If Data(0) = "select" Then
      
      Call SelectPiece(Int(Data(1)), Int(Data(2)))
      
      MovingX = Int(Data(1))
      
      MovingY = Int(Data(2))
      
      TempMovingX = Int(Data(1))
      
      TempMovingY = Int(Data(2))
      
       If Data(3) = "white" Then
      
        MovingWhiteIndex = Int(Data(4))
        
      
      Else
        
        MovingBlackIndex = Int(Data(4))
        
      
      End If
      
   
   End If
   
   If Data(0) = "change" Then Call ChangeNetPlayer
   
   If Data(0) = "unselect" Then
      
      Call RemoveSelection(Int(Data(1)), Int(Data(2)))
   
    End If
   
   If Data(0) = "move" Then
   
   
      Call MovePiece(Data(1), Int(Data(2)), Int(Data(3)), Int(Data(4)))
   
   End If
   
    If Data(0) = "multiremoveboolean" Then MultiRemove = True
   
    If Data(0) = "multimove" Then
   
    Call MultiMoves(Data(1), Int(Data(2)), Int(Data(3)), Int(Data(4)), True)
    
    
     MultiRemove = True
   
   End If
   
   LblSystem.Caption = DataString
   
End Sub

Private Sub Server_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'server on error
FormServer.Status.Panels(1).Text = "Error occured"
End Sub

Private Sub Timer1_Timer() 'used for debugging (to trace the crucial variables) + for indicating moves counters

LabelMakingMove.Caption = "MakingMove = " + Str(MakingMove)

LabelMultiMove.Caption = "MultiMove = " + Str(MultiMove)

Label2.Caption = "Moves made by BLACK player: " + Str(BlackCounter)

Label3.Caption = "Moves made by WHITE player: " + Str(WhiteCounter)


If ApplicationState <> "self" Then

     Label4.Caption = "Your color: " + Player

       Else

     Label4.Caption = ""


End If

End Sub


