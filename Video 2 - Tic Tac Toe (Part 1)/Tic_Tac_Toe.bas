Attribute VB_Name = "Tic_Tac_Toe"
Option Explicit


'current player
'board
'game over?

Private player1 As String
Private player2 As String

Private currentPlayer As Integer

Private board(1 To 3, 1 To 3) As Integer

Private gameOver As Boolean

Private winner As Integer

Public Sub gameHandler()

    init
    
    While Not gameOver
    
        'turn handler
        
        
        currentPlayer = currentPlayer Mod 2 + 1
        
        gameOver = True
        
    Wend
    


End Sub

Public Sub init()
    
    gameOver = False
    
    'x
    player1 = InputBox("What is Player 1's name?", "Player 1 Name Entry")
    
    'o
    player2 = InputBox("What is Player 2's name?", "Player 2 Name Entry")
    
    currentPlayer = 1
    'build randomness / coin toss to determine first player
    
    
End Sub
