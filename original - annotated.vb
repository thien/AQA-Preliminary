'Skeleton Program code for the AQA COMP1 Summer 2014 examination
'this code should be used in conjunction with the Preliminary Material
'written by the AQA COMP1 Programmer Team
'developed in the Visual Studio 2008 (Console Mode) programming environment (VB.NET)

Module CardPredict
    Const NoOfRecentScores As Integer = 3 'global variable NoOfRecentScores is recalled throughout the code. This can be modified safely

    Structure TCard 'Defines a structure, used in arrays relating to Decks.
        Dim Suit As Integer 'Matrix variable, You can assume that this is a fixed variable and will not change. (Cards don't randomly change suits.)
        Dim Rank As Integer 'Matrix variable, You can assume that this is a fixed variable and will not change. (Cards don't randomly change suits.)
    End Structure 'ends the structure

    Structure TRecentScore 'declares structure for the RecentScore Array
        Dim Name As String 'declares name variable, inset for name for array matrix
        Dim Score As Integer 'declares score variable, inset for score for array matrix
    End Structure 'ends the structure, terminating its code. May conserve memory

    Sub Main() 'Declares Main() subroutine, this is the first subroutine that will be run, visible interface. This is used to show the menu.
        Dim Choice As Char 'Declares Choice, this is the input for the options given from the menu.
        Dim Deck(52) As TCard 'Declares deck, as this is used as a parameter
        Dim RecentScores(NoOfRecentScores) As TRecentScore 'declares RecentScores array, uses TRecentScore Structure. This is going to be used heavilly.
        Randomize() 'Initialises Randomiser. Randomiser is a prebuilt function and is part of VB libraries. 
        Do 'Initiates Do Loop, expect this to repeat untill a desired condition is met. 
            DisplayMenu() 'Runs DisplayMenu() subroutine into Main().
            Choice = GetMenuChoice() 'recalls GetMenuChoice() and assigns it to choice. I'm assuming this is done to make the code more pretty. 
            Select Case Choice 'Uses variable Choice and reads its value, and checks for match with the following variables below.
                Case "1" 'If choice has this value it would execeute following lines:
                    LoadDeck(Deck) 'Runs LoadDeck sub. Loads the prementioned deck array as LoadDeck's parameter. At this stage it assigns values into deck array.
                    ShuffleDeck(Deck) 'Runs ShuffleDeck subroutine. Runs shuffle algorithm onto deck array, since LoadDeck assigns the deck array in order.  
                    PlayGame(Deck, RecentScores) 'Runs the game algorithm, Uses pre-declared parameters that were initialised in Sub Main()
                Case "2" 'If choice has this value it would execeute following lines:
                    LoadDeck(Deck)
                    PlayGame(Deck, RecentScores)
                Case "3"
                    DisplayRecentScores(RecentScores)
                Case "4"
                    ResetRecentScores(RecentScores)
            End Select
        Loop Until Choice = "q" 'Will repeat the loop until the variable Choice matches "q".
    End Sub

  Function GetRank(ByVal RankNo As Integer) As String
    Dim Rank As String = ""
    Select Case RankNo
      Case 1 : Rank = "Ace"
      Case 2 : Rank = "Two"
      Case 3 : Rank = "Three"
      Case 4 : Rank = "Four"
      Case 5 : Rank = "Five"
      Case 6 : Rank = "Six"
      Case 7 : Rank = "Seven"
      Case 8 : Rank = "Eight"
      Case 9 : Rank = "Nine"
      Case 10 : Rank = "Ten"
      Case 11 : Rank = "Jack"
      Case 12 : Rank = "Queen"
      Case 13 : Rank = "King"
    End Select
    Return Rank
  End Function

  Function GetSuit(ByVal SuitNo As Integer) As String
    Dim Suit As String = ""
    Select Case SuitNo
      Case 1 : Suit = "Clubs"
      Case 2 : Suit = "Diamonds"
      Case 3 : Suit = "Hearts"
      Case 4 : Suit = "Spades"
    End Select
    Return Suit
  End Function

  Sub DisplayMenu()
    Console.WriteLine()
    Console.WriteLine("MAIN MENU")
    Console.WriteLine()
    Console.WriteLine("1.  Play game (with shuffle)")
    Console.WriteLine("2.  Play game (without shuffle)")
    Console.WriteLine("3.  Display recent scores")
    Console.WriteLine("4.  Reset recent scores")
    Console.WriteLine()
    Console.Write("Select an option from the menu (or enter q to quit): ")
  End Sub

  Function GetMenuChoice() As Char
    Dim Choice As Char
    Choice = Console.ReadLine
    Console.WriteLine()
    Return Choice
  End Function

  Sub LoadDeck(ByRef Deck() As TCard)
    Dim Count As Integer
    FileOpen(1, "deck.txt", OpenMode.Input)
    Count = 1
    While Not EOF(1)
      Deck(Count).Suit = CInt(LineInput(1))
      Deck(Count).Rank = CInt(LineInput(1))
      Count = Count + 1
    End While
    FileClose(1)
  End Sub

  Sub ShuffleDeck(ByRef Deck() As TCard)
    Dim NoOfSwaps As Integer
    Dim Position1 As Integer
    Dim Position2 As Integer
    Dim SwapSpace As TCard
    Dim NoOfSwapsMadeSoFar As Integer
    NoOfSwaps = 1000
    For NoOfSwapsMadeSoFar = 1 To NoOfSwaps
      Position1 = Int(Rnd() * 52) + 1
      Position2 = Int(Rnd() * 52) + 1
      SwapSpace = Deck(Position1)
      Deck(Position1) = Deck(Position2)
      Deck(Position2) = SwapSpace
    Next
  End Sub

  Sub DisplayCard(ByVal ThisCard As TCard)
    Console.WriteLine()
    Console.WriteLine("Card is the " & GetRank(ThisCard.Rank) & " of " & GetSuit(ThisCard.Suit))
    Console.WriteLine()
  End Sub

  Sub GetCard(ByRef ThisCard As TCard, ByRef Deck() As TCard, ByVal NoOfCardsTurnedOver As Integer)
    Dim Count As Integer
    ThisCard = Deck(1)
    For Count = 1 To (51 - NoOfCardsTurnedOver)
      Deck(Count) = Deck(Count + 1)
    Next
    Deck(52 - NoOfCardsTurnedOver).Suit = 0
    Deck(52 - NoOfCardsTurnedOver).Rank = 0
  End Sub

  Function IsNextCardHigher(ByVal LastCard As TCard, ByVal NextCard As TCard) As Boolean
    Dim Higher As Boolean
    Higher = False
    If NextCard.Rank > LastCard.Rank Then
      Higher = True
    End If
    Return Higher
  End Function

  Function GetPlayerName() As String
    Dim PlayerName As String
    Console.WriteLine()
    Console.Write("Please enter your name: ")
    PlayerName = Console.ReadLine
    Console.WriteLine()
    Return PlayerName
  End Function

  Function GetChoiceFromUser() As Char
    Dim Choice As Char
    Console.Write("Do you think the next card will be higher than the last card (enter y or n)? ")
    Choice = Console.ReadLine
    Return Choice
  End Function

  Sub DisplayEndOfGameMessage(ByVal Score As Integer)
    Console.WriteLine()
    Console.WriteLine("GAME OVER!")
    Console.WriteLine("Your score was " & Score)
    If Score = 51 Then
      Console.WriteLine("WOW!  You completed a perfect game.")
    End If
    Console.WriteLine()
  End Sub

  Sub DisplayCorrectGuessMessage(ByVal Score As Integer)
    Console.WriteLine()
    Console.WriteLine("Well done!  You guessed correctly.")
    Console.WriteLine("Your score is now " & Score & ".")
    Console.WriteLine()
  End Sub

  Sub ResetRecentScores(ByRef RecentScores() As TRecentScore)
    Dim Count As Integer
    For Count = 1 To NoOfRecentScores
      RecentScores(Count).Name = ""
      RecentScores(Count).Score = 0
    Next
  End Sub

  Sub DisplayRecentScores(ByVal RecentScores() As TRecentScore)
    Dim Count As Integer
    Console.WriteLine()
    Console.WriteLine("Recent scores:")
    Console.WriteLine()
    For Count = 1 To NoOfRecentScores
      Console.WriteLine(RecentScores(Count).Name & " got a score of " & RecentScores(Count).Score)
    Next
    Console.WriteLine()
    Console.WriteLine("Press the Enter key to return to the main menu")
    Console.WriteLine()
    Console.ReadLine()
  End Sub

  Sub UpdateRecentScores(ByRef RecentScores() As TRecentScore, ByVal Score As Integer)
    Dim PlayerName As String
    Dim Count As Integer
    Dim FoundSpace As Boolean
    PlayerName = GetPlayerName()
    FoundSpace = False
    Count = 1
    While Not FoundSpace And Count <= NoOfRecentScores
      If RecentScores(Count).Name = "" Then
        FoundSpace = True
      Else
        Count = Count + 1
      End If
    End While
    If Not FoundSpace Then
      For Count = 1 To NoOfRecentScores - 1
        RecentScores(Count) = RecentScores(Count + 1)
      Next
      Count = NoOfRecentScores
    End If
    RecentScores(Count).Name = PlayerName
    RecentScores(Count).Score = Score
  End Sub

  Sub PlayGame(ByVal Deck() As TCard, ByRef RecentScores() As TRecentScore)
    Dim NoOfCardsTurnedOver As Integer
    Dim GameOver As Boolean
    Dim NextCard As TCard
    Dim LastCard As TCard
    Dim Higher As Boolean
    Dim Choice As Char
    GameOver = False
    GetCard(LastCard, Deck, 0)
    DisplayCard(LastCard)
    NoOfCardsTurnedOver = 1
    While NoOfCardsTurnedOver < 52 And Not GameOver
      GetCard(NextCard, Deck, NoOfCardsTurnedOver)
      Do
        Choice = GetChoiceFromUser()
      Loop Until Choice = "y" Or Choice = "n"
      DisplayCard(NextCard)
      NoOfCardsTurnedOver = NoOfCardsTurnedOver + 1
      Higher = IsNextCardHigher(LastCard, NextCard)
      If Higher And Choice = "y" Or Not Higher And Choice = "n" Then
        DisplayCorrectGuessMessage(NoOfCardsTurnedOver - 1)
        LastCard = NextCard
      Else
        GameOver = True
      End If
    End While
    If GameOver Then
      DisplayEndOfGameMessage(NoOfCardsTurnedOver - 2)
      UpdateRecentScores(RecentScores, NoOfCardsTurnedOver - 2)
    Else
      DisplayEndOfGameMessage(51)
      UpdateRecentScores(RecentScores, 51)
    End If
  End Sub
End Module
