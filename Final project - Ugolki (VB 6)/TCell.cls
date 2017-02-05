VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Color As String, WithEvents Cell As Image
Attribute Cell.VB_VarHelpID = -1
Public PositionX As Integer, PositionY As Integer
Public Taken As Boolean, PossibleToMoveHere As Boolean

Public Sub Cell_Click()

If Form1.MakingMove = True Then

  
      If Form1.MultiMove = False Then

            If Form1.WhiteMoves = True Then
 
                Call Form1.SimpleMove("white", Form1.MovingWhiteIndex, PositionX, PositionY)
                   
                   
                   
              Else
 
                Call Form1.SimpleMove("black", Form1.MovingBlackIndex, PositionX, PositionY)
                
                
 
            End If


       Else

            If Form1.WhiteMoves = True Then
 
                Call Form1.MultiMoves("white", Form1.MovingWhiteIndex, PositionX, PositionY, False)
               
              If Form1.ApplicationState = "server" Then Form1.Server.SendData "multimove|white|" & Form1.MovingWhiteIndex & "|" & PositionX & "|" & PositionY
              If Form1.ApplicationState = "client" Then Form1.Client.SendData "multimove|white|" & Form1.MovingWhiteIndex & "|" & PositionX & "|" & PositionY
               
              Else
 
                Call Form1.MultiMoves("black", Form1.MovingBlackIndex, PositionX, PositionY, False)
                
              If Form1.ApplicationState = "server" Then Form1.Server.SendData "multimove|black|" & Form1.MovingBlackIndex & "|" & PositionX & "|" & PositionY
              If Form1.ApplicationState = "client" Then Form1.Client.SendData "multimove|black|" & Form1.MovingBlackIndex & "|" & PositionX & "|" & PositionY
                
 
            End If

      End If



End If

End Sub

