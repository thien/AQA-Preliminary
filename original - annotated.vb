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
                    LoadDeck(Deck) 'Same thing as beforehand. Necessary as code is wiped when it has finished running.
                    PlayGame(Deck, RecentScores) 'Runs game algorithm, but because ShuffleDeck hasn't been run beforehand, the deck here is unshuffled. 
                Case "3" 'If choice has this value it would execeute following lines:
                    DisplayRecentScores(RecentScores) 'Runs DisplayRecentScores subroutine with RecentScores as its variables. (parameters)'
                Case "4" 'If choice has this value it would execeute following lines:
                    ResetRecentScores(RecentScores) 'Runs ResetRecentScores subroutine with RecentScores as its variables. (parameters)'
            End Select 'Ends case statement, no more choices are needed for now. Of course, you can make a case for "q"
        Loop Until Choice = "q" 'Repeats the loop until the Choice matches "q". Do until loops are used as the condition would need to be met @ the end of the code 
    End Sub 'Ends the sub of Main. Necessary otherwise code is long as balls and tedious to read, and would make AQA arseholes'

  Function GetRank(ByVal RankNo As Integer) As String 'Declares function GetRank, and declares its parameters. Useful to use parameters instead of global variables
    Dim Rank As String = "" 'declares rank, these are going to be used to represent the different values of the Ranks a card would have. 
    Select Case RankNo 'Recalls parameter and matches it with one of its declared cases. This is the rank the individual card in the array would have.
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

  Function GetSuit(ByVal SuitNo As Integer) As String 'Literal same thing from GetRank, but for suits. These are used for TCard.
    Dim Suit As String = ""
    Select Case SuitNo
      Case 1 : Suit = "Clubs"
      Case 2 : Suit = "Diamonds"
      Case 3 : Suit = "Hearts"
      Case 4 : Suit = "Spades"
    End Select
    Return Suit
  End Function

  Sub DisplayMenu() 'simple subroutine that only shows text. Used to display the menu.'
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

  Function GetMenuChoice() As Char 'Simple input function.
    Dim Choice As Char
    Choice = Console.ReadLine
    Console.WriteLine()
    Return Choice
  End Function

  Sub LoadDeck(ByRef Deck() As TCard) 'Loads deck
    Dim Count As Integer
    FileOpen(1, "deck.txt", OpenMode.Input) 'Uses one of VB's premade functions (FileOpen), first variable is the file number, second is the location of txt. Third is the mode of the text. 
    Count = 1
    While Not EOF(1) 'While not end of file, this is the determiner of the while loop.'
      Deck(Count).Suit = CInt(LineInput(1)) 'Recalls one line of code and assigns it to Deck(Count).suit'
      Deck(Count).Rank = CInt(LineInput(1)) 'recalls following line of that code and assigns it to Deck(Count).rank'
      Count = Count + 1 'increments by 1, this count is incremented, so it acts as a stepper.'
    End While 'ends loop'
    FileClose(1) 'Closes file of 1. Good practice to close the file when you're done so it isn't hogging up memory.
  End Sub

  Sub ShuffleDeck(ByRef Deck() As TCard) 'Declares ShuffleDeck subroutine, uses Deck as parameter.'
    Dim NoOfSwaps As Integer
    Dim Position1 As Integer
    Dim Position2 As Integer
    Dim SwapSpace As TCard 'declared as TCard as it is holding a whole card and merely not a part.'
    Dim NoOfSwapsMadeSoFar As Integer
    NoOfSwaps = 1000 'This is going to be how many times the following loop is repeated, to ensure a good shuffle.'
    For NoOfSwapsMadeSoFar = 1 To NoOfSwaps 'Declares loop and its rules.'
      Position1 = Int(Rnd() * 52) + 1 'Int turns the position into a integer, its not nice to have half a card.
      Position2 = Int(Rnd() * 52) + 1 'adds 1 to as Random is a generated decimal between 0 and 1. '
      SwapSpace = Deck(Position1) 'Swapspace is a temporary holder. 
      Deck(Position1) = Deck(Position2) 'Most Recent Holders'
      Deck(Position2) = SwapSpace 'Most Recent Holders'
    Next
  End Sub

  Sub DisplayCard(ByVal ThisCard As TCard) 'simple display subroutine, parameters are used as good practice. This is going to be used a lot in the game. 
    Console.WriteLine()
    Console.WriteLine("Card is the " & GetRank(ThisCard.Rank) & " of " & GetSuit(ThisCard.Suit)) 'Displays card and rank'
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

  Function IsNextCardHigher(ByVal LastCard As TCard, ByVal NextCard As TCard) As Boolean 'simple comparison algorithm'
    Dim Higher As Boolean
    Higher = False
    If NextCard.Rank > LastCard.Rank Then
      Higher = True
    End If
    Return Higher
  End Function

  Function GetPlayerName() As String 'simple input function, serves to get players name'
    Dim PlayerName As String
    Console.WriteLine()
    Console.Write("Please enter your name: ")
    PlayerName = Console.ReadLine
    Console.WriteLine()
    Return PlayerName
  End Function

  Function GetChoiceFromUser() As Char 'simple input function'
    Dim Choice As Char
    Console.Write("Do you think the next card will be higher than the last card (enter y or n)? ")
    Choice = Console.ReadLine
    Return Choice
  End Function

  Sub DisplayEndOfGameMessage(ByVal Score As Integer) 'Displays scores when the game has ended. '
    Console.WriteLine()
    Console.WriteLine("GAME OVER!")
    Console.WriteLine("Your score was " & Score)
    If Score = 51 Then
      Console.WriteLine("WOW!  You completed a perfect game.")
    End If
    Console.WriteLine()
  End Sub

  Sub DisplayCorrectGuessMessage(ByVal Score As Integer) 'Displays scores in the progress of the game.
    Console.WriteLine()
    Console.WriteLine("Well done!  You guessed correctly.")
    Console.WriteLine("Your score is now " & Score & ".")
    Console.WriteLine()
  End Sub

  Sub ResetRecentScores(ByRef RecentScores() As TRecentScore) 
    Dim Count As Integer
    For Count = 1 To NoOfRecentScores 'Loop, sets the entire array's contents to how VB represents an empty variable. Blank.
      RecentScores(Count).Name = ""
      RecentScores(Count).Score = 0
    Next
  End Sub

  Sub DisplayRecentScores(ByVal RecentScores() As TRecentScore) 'simple display scores sub'
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

  Sub UpdateRecentScores(ByRef RecentScores() As TRecentScore, ByVal Score As Integer) 'Declares subroutine and its parameters'
    Dim PlayerName As String 
    Dim Count As Integer
    Dim FoundSpace As Boolean 'can only have two possible values, hence boolean. plus, boolean is very light in memory.
    PlayerName = GetPlayerName() 'Uses function to recall Playername.'
    FoundSpace = False 'safety first; Set to false before assuming anything.'
    Count = 1
    While Not FoundSpace And Count <= NoOfRecentScores 'Checks if there is space in the array'
      If RecentScores(Count).Name = "" Then
        FoundSpace = True 'If there is space it'll make the boolean true
      Else
        Count = Count + 1 'Keeps going. 
      End If
    End While
    If Not FoundSpace Then 'if theres no space then it'll do this
      For Count = 1 To NoOfRecentScores - 1 'moves everything in the array up by 1.'
        RecentScores(Count) = RecentScores(Count + 1)
      Next
      Count = NoOfRecentScores 
    End If
    RecentScores(Count).Name = PlayerName 'assigns playername to that specific matrice in the array'
    RecentScores(Count).Score = Score 'assigns score to that specific matrice in the array'
  End Sub

  Sub PlayGame(ByVal Deck() As TCard, ByRef RecentScores() As TRecentScore) 'declares playgame subroutine. 
    Dim NoOfCardsTurnedOver As Integer 
    Dim GameOver As Boolean
    Dim NextCard As TCard
    Dim LastCard As TCard
    Dim Higher As Boolean
    Dim Choice As Char 'input is only 1 character long so declaring it as a string is unnescessary and violates good programming principles'
    GameOver = False
    GetCard(LastCard, Deck, 0) 'represents the first card. 0 is there as there are no cards turned over before this.'
    DisplayCard(LastCard) 'shows card to player.'
    NoOfCardsTurnedOver = 1 'since card is shown, the card is turned over. (in order to see the damn card) #logic '
    While NoOfCardsTurnedOver < 52 And Not GameOver 'initates loop where it keeps going if theres still cards in the game, or if its not game over'
      GetCard(NextCard, Deck, NoOfCardsTurnedOver) 
      Do 'initiates do loop'
        Choice = GetChoiceFromUser() 'gets user input using the GetChoiceFromUser function. function is easier to plop throughout code.
      Loop Until Choice = "y" Or Choice = "n" 'repeats line untill correct input is plopped in'
      DisplayCard(NextCard) 'shows the next card, insinuates user wether they are unfortunately right or idiotically wrong'
      NoOfCardsTurnedOver = NoOfCardsTurnedOver + 1 'stepper increments by one (this is removed at game over)
      Higher = IsNextCardHigher(LastCard, NextCard) 'recalls function to determine if next card is higher or not'
      If Higher And Choice = "y" Or Not Higher And Choice = "n" Then 'uses if statement to determine wether the user is right (this is for right)'
        DisplayCorrectGuessMessage(NoOfCardsTurnedOver - 1)
        LastCard = NextCard
      Else 'indicates THE USER IS FFEEBLE AND SHOULD UNINSTALL THE GAME'
        GameOver = True
      End If 'ends if statement'
    End While 'ends while'
    If GameOver Then 'checks if the game is over, since the entire sub is repeated if it isnt, continuing the game'
      DisplayEndOfGameMessage(NoOfCardsTurnedOver - 2) 'displays endofgamemsg, parameter is score'
      UpdateRecentScores(RecentScores, NoOfCardsTurnedOver - 2) 'adds score to recentscores, name is asked in called subroutine'
    Else
      DisplayEndOfGameMessage(51) 'player got everything, technically wins the game so it displays full score. 
      UpdateRecentScores(RecentScores, 51) 'adds score to recentscores, name is asked in the called sub'
    End If
  End Sub
End Module 'end of entire code'
