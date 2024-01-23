Sub Button3_Click()

Dim Deck As Worksheet
Set Deck = Worksheets("Deck_Order")

Dim Results As Worksheet
Set Results = Worksheets("Sim")

Dim Score As Worksheet
Set Score = Worksheets("Score")


suit = Array("h", "s", "d", "c")

rnk = Array("2", "3", "4", "5", "6", "7", "8", "9", "T", "J", "Q", "K", "A")

Dim Iterations As Range
Set Iterations = Results.Range("F2:R500000")

   
Iterations.Clear


Dim cards(0 To 51) As String
For i = 0 To 3
For j = 0 To 12
cards(13 * i + j) = rnk(j) & suit(i)
Next j
Next i

For Z = 1 To Results.Range("D2").Value

'Shuffle
For i = 1 To 1000
c1 = Int(52 * Rnd)
c2 = Int(52 * Rnd)
temp = cards(c1)
cards(c1) = cards(c2)
cards(c2) = temp
Next i

For X = 0 To 51
    Deck.Cells(X + 1, 1) = cards(X)
Next X

'Card Removal (cards in our hand are no longer in the deck)
For Y = 1 To 52

    If Deck.Cells(Y, 1) = Results.Range("B2") Then Deck.Cells(Y, 1).Delete
    If Deck.Cells(Y, 1) = Results.Range("B3") Then Deck.Cells(Y, 1).Delete
    
Next Y

'Deal first 5 cards

    Results.Cells(Z + 1, 11) = Deck.Cells(1, 1)
    Results.Cells(Z + 1, 12) = Deck.Cells(2, 1)
    Results.Cells(Z + 1, 13) = Deck.Cells(3, 1)
    Results.Cells(Z + 1, 14) = Deck.Cells(4, 1)
    Results.Cells(Z + 1, 15) = Deck.Cells(5, 1)
    Results.Cells(Z + 1, 6) = Z
    
'Scoring the Hand
    Results.Cells(Z + 1, 7) = Results.Range("B2")
    Results.Cells(Z + 1, 8) = Results.Range("C2")
    
    
    Score.Cells(1, 3) = Z
    Results.Cells(Z + 1, 16) = Score.Cells(29, 7)
    
    
Next Z


'Automatically Storing the Data
Dim Data As Worksheet
Set Data = Worksheets("Data")


nextrow = Data.Cells(Rows.Count, 1).End(xlUp).Row + 1


Data.Cells(nextrow, 1) = Results.Range("B2")
Data.Cells(nextrow, 2) = Results.Range("C2")

For X = 1 To 9

Data.Cells(nextrow, X + 2) = Results.Cells(X + 10, 2)
Data.Cells(nextrow, X + 12) = Results.Cells(X + 10, 3)

Next X

Data.Cells(nextrow, 12) = Results.Range("B21")


End Sub
