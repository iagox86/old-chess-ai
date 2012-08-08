Attribute VB_Name = "modChess"
Option Explicit

' The chessboard is going to be an array of 64 ints
' which will each represent a square.  Here are the
' piece names:
Public Const Blank = 0
Public Const bBishop = 1
Public Const bKing = 2
Public Const bKnight = 3
Public Const bPawn = 4
Public Const bQueen = 5
Public Const bRook = 6
Public Const wBishop = 7
Public Const wKing = 8
Public Const wKnight = 9
Public Const wPawn = 10
Public Const wQueen = 11
Public Const wRook = 12

Public Gameboard(0 To 7, 0 To 7) As Integer
Public Lastboard(0 To 7, 0 To 7) As Integer
Public Lastlastboard(0 To 7, 0 To 7) As Integer

Public EmptyBoard(0 To 7, 0 To 7) As Integer

Sub main()
    Dim row As Integer
    Dim column As Integer

    EmptyBoard(0, 0) = bRook
    EmptyBoard(0, 1) = bKnight
    EmptyBoard(0, 2) = bBishop
    EmptyBoard(0, 3) = bQueen
    EmptyBoard(0, 4) = bKing
    EmptyBoard(0, 5) = bBishop
    EmptyBoard(0, 6) = bKnight
    EmptyBoard(0, 7) = bRook
    EmptyBoard(1, 0) = bPawn
    EmptyBoard(1, 1) = bPawn
    EmptyBoard(1, 2) = bPawn
    EmptyBoard(1, 3) = bPawn
    EmptyBoard(1, 4) = bPawn
    EmptyBoard(1, 5) = bPawn
    EmptyBoard(1, 6) = bPawn
    EmptyBoard(1, 7) = bPawn
    
    For row = 2 To 5
        For column = 0 To 7
            EmptyBoard(row, column) = Blank
        Next
    Next
    
    EmptyBoard(6, 0) = wPawn
    EmptyBoard(6, 1) = wPawn
    EmptyBoard(6, 2) = wPawn
    EmptyBoard(6, 3) = wPawn
    EmptyBoard(6, 4) = wPawn
    EmptyBoard(6, 5) = wPawn
    EmptyBoard(6, 6) = wPawn
    EmptyBoard(6, 7) = wPawn
    EmptyBoard(7, 0) = wRook
    EmptyBoard(7, 1) = wKnight
    EmptyBoard(7, 2) = wBishop
    EmptyBoard(7, 3) = wKing
    EmptyBoard(7, 4) = wQueen
    EmptyBoard(7, 5) = wBishop
    EmptyBoard(7, 6) = wKnight
    EmptyBoard(7, 7) = wRook
    
    frmMain.Show

End Sub

Function GetColumn(Location As Integer) As Integer
    GetColumn = ((Location) Mod 8)
End Function

Function GetRow(Location As Integer) As Integer
    GetRow = Int(Location / 8)
End Function

Function GetLocation(row As Integer, col As Integer) As Integer
    GetLocation = col - 1 + ((row - 1) * 8)
End Function
