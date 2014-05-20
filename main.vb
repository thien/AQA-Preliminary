Module CardPredict
    Dim playername As String = "Guest"
    Dim mode As Integer = 0
    Dim globcount As String
    Const NoOfRecentScores As Integer = 20
    Structure TCard
        Dim Suit As Integer
        Dim Rank As Integer
    End Structure
    Structure TRecentScore
        Dim Name As String
        Dim Score As String
        Dim istt As String
    End Structure
    Sub header(ByVal mode As Integer)
        Console.Clear()
        Console.WriteLine()
        Console.Write("Player: " & playername & " | Directory: Menu")
        Select Case mode
            Case Is = "1"
            Case Is = "2"
                Console.Write("/CardGame")
            Case Is = "3"
                Console.Write("/CardGame(anti-shuffle)")
            Case Is = "4"
                Console.Write("/Options")
            Case Is = "5"
                Console.Write("/Highscores")
            Case Is = "6"
                Console.Write("/Sign In")
            Case Is = "7"
                Console.Write("/Wipe Scores")
        End Select
        If mode = 2 Or mode = 3 Then
            Console.Write(" | Score: " & globcount)
        End If
        longline()
    End Sub
    Sub Main()
        Dim Choice As Char
        Dim options As String
        Dim Deck(52) As TCard
        Dim RecentScores(NoOfRecentScores) As TRecentScore
        'ResetRecentScores(RecentScores)
        loadscores(RecentScores)
        Randomize()

        Do
            header(1)
            DisplayMenu()
            Choice = GetMenuChoice()
            Select Case Choice
                Case "1"
                    Do
                        header(1)
                        Console.WriteLine("Your difficulty, good sir?")
                        Console.WriteLine("1. Easy")
                        Console.WriteLine("2. Medium")
                        Console.WriteLine("3. Hard")
                        Console.WriteLine("4. Rape")
                        Console.WriteLine("5. Time Attack")
                        Console.WriteLine("q. Back")
                        options = LCase(Console.ReadKey.KeyChar)
                        Console.WriteLine()
                        Select Case options
                            Case "1"
                                mode = 3
                                makedeck(Deck)
                                PlayGame(Deck, RecentScores)
                                mode = 0
                            Case "2"
                                mode = 2
                                makedeck(Deck)
                                ShuffleDeck(Deck)
                                PlayGame(Deck, RecentScores)
                                mode = 0
                            Case "5"
                                mode = 3
                                makedeck(Deck)
                                PlayTimeGame(Deck, RecentScores)
                            Case "q"
                                Console.WriteLine("")
                            Case Else
                                Console.WriteLine("invalid input, please retry")
                                Pause()
                        End Select
                    Loop Until options = "1" Or options = "2" Or options = "q"
                    mode = 0
                Case "2"
                    header(6)
                    GetPlayerName()
                Case "3"
                    mode = 5
                    DisplayRecentScores(RecentScores)
                    DisplayBestTimeRatioScores(RecentScores)
                Case "4"
                    header(7)
                    ResetRecentScores(RecentScores)
                Case "q"
                    safeg()
                Case Else
                    Console.WriteLine("invalid input - please retry")
                    Pause()
            End Select
        Loop Until Choice = "q"
    End Sub
    Sub options()

    End Sub
    Sub safeg()
        Console.WriteLine()
        Console.WriteLine("k thx bay")
        Console.ReadLine()
    End Sub
    Sub longline()
        Console.WriteLine()
        Console.WriteLine("________________________________________________________________________________")
    End Sub
    Sub Pause()
        Dim safety As Char
        Console.WriteLine()
        Console.WriteLine("Press any key to continue..")
        safety = Console.ReadKey.KeyChar
        Console.WriteLine()
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
        Console.WriteLine("1.  Play game")
        Console.WriteLine("2.  Sign In")
        Console.WriteLine("3.  Display recent scores")
        Console.WriteLine("4.  Reset recent scores")
        Console.WriteLine()
        Console.Write("Select an option from the menu (or enter q to quit): ")
    End Sub

    Function GetMenuChoice() As Char
        Dim Choice As Char
        Choice = LCase(Console.ReadKey.KeyChar)
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
    Sub makedeck(ByRef deck() As TCard)
        Dim suit, count, rank As Integer
        count = 1
        For rank = 1 To 13 Step 1
            For suit = 1 To 4 Step 1
                deck(count).Suit = suit
                deck(count).Rank = rank
                count = count + 1
            Next
        Next
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
    Function isnextcardequal(ByVal LastCard As TCard, ByVal NextCard As TCard)
        Dim equal As Boolean = False
        If NextCard.Rank = LastCard.Rank Then
            equal = True
        End If
        Return equal
    End Function

    Function GetPlayerName() As String
        Console.WriteLine()
        Do
            Console.Write("Please enter your name: ")
            playername = Console.ReadLine
            Console.WriteLine()
            If playername.Length = 0 Then
                Console.WriteLine("Invalid name, Too short or too long.")
            End If
        Loop Until playername <> "Guest"
        Return playername
    End Function
    Function GetChoiceFromUser() As Char
        Dim Choice As String
        Console.Write("Do you think the next card will be higher than the last card (enter y or n)? ")
        Choice = Console.ReadKey.KeyChar
        Select Case Choice
            Case "y"
            Case "n"
            Case Else
                Console.WriteLine()
                Console.WriteLine("wtf? try again")
                Pause()
        End Select
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
        Dim check As String
        Do
            header(7)
            Console.WriteLine()
            Console.WriteLine("WIPE SCORES")
            Console.WriteLine()
            Console.WriteLine("Are you sure you want to do this? type 'yes' or 'no'")
            check = Console.ReadLine()
            If check <> "yes" And check <> "no" Then
                Console.WriteLine("Invalid answer, try again.")
                Pause()
            End If
        Loop Until check = "yes" Or check = "no"
        If check = "yes" Then
            Dim Count As Integer
            For Count = 1 To NoOfRecentScores
                RecentScores(Count).Name = ""
                RecentScores(Count).Score = 0
                RecentScores(Count).istt = "False"
            Next
            Console.WriteLine("The scoreboard is reset.")
        ElseIf check = "no" Then
            Console.WriteLine("The scoreboard remains intact.")
        End If
        Pause()
    End Sub

    Sub DisplayRecentScores(ByVal RecentScores() As TRecentScore)
        header(5)
        Dim Count As Integer
        Dim score As Integer
        Console.WriteLine()
        Console.WriteLine("Scoreboard - Normal Mode")
        Console.WriteLine()

        For Count = 1 To NoOfRecentScores
            If RecentScores(Count).istt = "False" Then
                If RecentScores(Count).Name <> "" Then
                    score = score + 1
                    Console.WriteLine(RecentScores(Count).Name & " got a score of " & RecentScores(Count).Score & " at " & RecentScores(Count).istt)
                End If
            End If
        Next
        If score = 0 Then
            Console.WriteLine("there are no recent scores.")
        End If
        Pause()
    End Sub
    Sub DisplayBestTimeRatioScores(ByRef recentscores() As TRecentScore)
        header(5)
        Dim Count As Integer
        Dim temp As TRecentScore
        Dim bubblecount, mark As Integer
        For mark = NoOfRecentScores - 1 To 1 Step -1
            For Count = 1 To mark Step 1
                If recentscores(bubblecount).istt = "True" And recentscores(bubblecount + 1).istt = True Then
                    If recentscores(bubblecount).Score < recentscores(bubblecount + 1).Score Then
                        temp = recentscores(bubblecount + 1)
                        recentscores(bubblecount + 1) = recentscores(bubblecount)
                        recentscores(bubblecount) = temp
                    End If
                End If
            Next
        Next


        Console.WriteLine()
        Console.WriteLine("Scoreboard")
        Console.WriteLine()

        For Count = 1 To NoOfRecentScores Step 1
            If recentscores(Count).istt = True Then
                Console.WriteLine(recentscores(Count).Name & " got a score of " & recentscores(Count).Score)
            End If
        Next

        Pause()
    End Sub
    'Sub DisplayBestTimes(ByVal recentscores() As TRecentScore)
    '    Dim count, mark As Integer
    '    Dim count2 As Integer
    '    Dim temp As TRecentScore
    '    For mark = NoOfRecentScores - 1 To 1 Step -1
    '        For count = 1 To mark Step 1
    '            If recentscores(count).Period < recentscores(count + 1).Period Then
    '                temp = recentscores(count + 1)
    '                recentscores(count + 1) = recentscores(count)
    '                recentscores(count) = temp
    '            End If
    '        Next
    '    Next

    '    Console.WriteLine()
    '    Console.WriteLine("Scoreboard - Best Speed/Time Ratio")
    '    Console.WriteLine()
    '    Console.WriteLine("1st Place belongs to " & recentscores(1).Name & " with " & recentscores(1).Period & "pps")
    '    If recentscores(2).Name <> "" Then
    '        Console.WriteLine("2nd Place belongs to " & recentscores(2).Name & " with " & recentscores(2).Period & "pps")
    '    End If
    '    If recentscores(3).Name <> "" Then
    '        Console.WriteLine("3rd Place belongs to " & recentscores(3).Name & " with " & recentscores(3).Period & "pps")
    '    End If
    '    Console.WriteLine()
    '    Console.WriteLine("Runner Ups")
    '    Console.WriteLine()

    '    For count2 = 4 To NoOfRecentScores
    '        If recentscores(count2).Name <> "" Then
    '            Console.WriteLine(recentscores(count2).Name & " got a score of " & recentscores(count2).Score & " at " & recentscores(count2).Period & "pps")
    '        End If
    '    Next
    '    Pause()

    'End Sub


    Sub UpdateRecentScores(ByRef RecentScores() As TRecentScore, ByVal Score As Integer, ByVal istimeattack As Boolean)
        Pause()
        Dim Count As Integer
        Dim player As String = playername
        Dim save As Char
        Dim FoundSpace As Boolean
        If Score >= 5 Then
            Console.WriteLine("Woah, You scored high enough to be in the scoreboard!")
            Do
                Console.WriteLine("Would you like to save your score? type 'y' or 'n'. ")
                save = Console.ReadKey.KeyChar()
                If save = "y" Then
                    If playername = "Guest" Then
                        playername = GetPlayerName()
                    End If
                ElseIf save = "n" Then
                    Console.WriteLine("not saved.")
                    Pause()
                End If
            Loop Until save = "y" Or save = "n"

            If save = "y" Then
                player = playername
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
                RecentScores(Count).Name = player
                RecentScores(Count).Score = Score
                RecentScores(Count).istt = istimeattack
                Console.WriteLine("Score logged. :')")
                sortscores(RecentScores)
                savescores(RecentScores)

                Console.WriteLine("Score saved to scoreboard.")
                Pause()
            End If
        Else
            Console.WriteLine("You don't score high enough to join the high score.")
            Pause()
        End If


    End Sub


    Sub probability(ByVal deck() As TCard, ByVal currentrank As Integer)
        Dim probability, cardcount, tracker, numberhigher As Integer
        cardcount = 0
        numberhigher = 0
        For tracker = 1 To 52 Step 1
            If deck(tracker).Rank > currentrank Then
                numberhigher = numberhigher + 1
            End If
            If deck(tracker).Rank <> 0 Then
                cardcount = cardcount + 1
            End If
        Next
        probability = (numberhigher / cardcount) * 100
        Console.WriteLine("probability of card higher is " & probability & "%")
        Console.WriteLine()
    End Sub
    Function scoretimeratio(ByVal score As Integer, ByVal timer As Boolean, ByVal starttime As DateTime, ByVal cachetime As DateTime)
        Dim result As Decimal
        Dim timestring As Decimal
        Dim elapsed_cache As TimeSpan
        If timer = True Then
            elapsed_cache = cachetime.Subtract(starttime)
            timestring = elapsed_cache.TotalSeconds.ToString("0.000000")
            timestring = timestring ^ -0.5
            Select Case score
                Case Is < 10
                    score = score * 2.5
                Case Is > 10
                    score = score * 3.14
                Case Is > 15
                    score = score * 4.52
                Case Is > 20
                    score = score * 6.32
                Case Is > 25
                    score = score * 10
            End Select
            'Need to revisit multiplier
            result = FormatNumber(result, 2)
            result = score * timestring
            result = FormatNumber(result, 2)
        Else
            result = score
        End If
        Return result
    End Function
    Sub sortscores(ByRef recentscores() As TRecentScore)
        Dim count, mark As Integer
        Dim temp As TRecentScore
        For mark = NoOfRecentScores - 1 To 1 Step -1
            For count = 1 To mark Step 1
                If recentscores(count).Score < recentscores(count + 1).Score Then
                    temp = recentscores(count + 1)
                    recentscores(count + 1) = recentscores(count)
                    recentscores(count) = temp
                End If
            Next
        Next
    End Sub
    Sub savescores(ByVal recentscores() As TRecentScore)
        Dim count As Integer
        Dim write As New IO.StreamWriter("scores.txt", False)
        For count = 1 To NoOfRecentScores Step 1
            write.WriteLine(recentscores(count).Name)
            write.WriteLine(recentscores(count).Score)
            write.WriteLine(recentscores(count).istt)
        Next
        write.Close()
    End Sub
    Sub loadscores(ByRef recentscores() As TRecentScore)
        FileOpen(1, "scores.txt", OpenMode.Input)
        Dim count As Integer = 1
        While (Not EOF(1) And (Not count > NoOfRecentScores))
            recentscores(count).Name = LineInput(1)
            recentscores(count).Score = LineInput(1)
            recentscores(count).istt = LineInput(1)
            count = count + 1
        End While
        FileClose(1)
    End Sub
    Sub PlayTwoPlayerGame(ByVal Deck() As TCard, ByRef RecentScores() As TRecentScore)
        Dim NoOfCardsTurnedOver As Integer
        Dim GameOver As Boolean
        Dim NextCard As TCard
        Dim LastCard As TCard
        Dim player1Higher As Boolean
        Dim player2higher As Boolean
        Dim player1correct As Boolean
        Dim player2correct As Boolean
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
            player1Higher = IsNextCardHigher(LastCard, NextCard)
            If player1Higher And Choice = "y" Or Not player1Higher And Choice = "n" Then
                player1correct = True
                DisplayCorrectGuessMessage(NoOfCardsTurnedOver - 1)
                LastCard = NextCard
                Console.WriteLine("player1 got it right. player 2 gotta stay in the game to keep it real")
                Console.WriteLine()
            Else
                LastCard = NextCard
                player1correct = False
                Console.WriteLine("player 1 got it wrong. it's up to player 2 to stay in the game.")
                Console.WriteLine()
            End If
            GetCard(NextCard, Deck, NoOfCardsTurnedOver)
            Console.WriteLine("player 2's turn. go forth")
            Console.WriteLine()
            Do
                Choice = GetChoiceFromUser()
            Loop Until Choice = "y" Or Choice = "n"
            DisplayCard(NextCard)
            NoOfCardsTurnedOver = NoOfCardsTurnedOver + 1
            player2higher = IsNextCardHigher(LastCard, NextCard)
            If player2higher And Choice = "y" Or Not player2higher And Choice = "n" Then
                player2correct = True
                DisplayCorrectGuessMessage(NoOfCardsTurnedOver - 1)
                LastCard = NextCard
                Console.WriteLine("player 2 guessed right.")
            Else
                LastCard = NextCard
                player2correct = False
                Console.WriteLine("fail.")
            End If
            If player1correct <> player2correct Then
                GameOver = True
            End If
        End While

        If GameOver = True Then
            If player1correct = True Then
                Console.WriteLine("player 1 wins")
            Else
                Console.WriteLine("player 2 wins")
            End If
            DisplayEndOfGameMessage(NoOfCardsTurnedOver - 2)
        Else
            DisplayEndOfGameMessage(51)
        End If
    End Sub
    Sub PlayTimeGame(ByVal Deck() As TCard, ByRef RecentScores() As TRecentScore)
        Dim NoOfCardsTurnedOver As Integer
        Dim GameOver As Boolean
        Dim NextCard As TCard
        Dim LastCard As TCard
        Dim Higher As Boolean
        Dim equal As Boolean
        Static start_time As DateTime
        Static cache_time As DateTime
        Static stop_time As DateTime
        Dim elapsed_time As TimeSpan
        Dim speedscore As Integer
        Dim Choice As Char
        GameOver = False
        start_time = Now
        NoOfCardsTurnedOver = 1
        globcount = NoOfCardsTurnedOver
        header(2)
        GetCard(LastCard, Deck, 0)
        DisplayCard(LastCard)
        While NoOfCardsTurnedOver < 52 And Not GameOver
            GetCard(NextCard, Deck, NoOfCardsTurnedOver)
            Do
                Choice = GetChoiceFromUser()
            Loop Until Choice = "y" Or Choice = "n"
            cache_time = Now
            globcount = scoretimeratio(NoOfCardsTurnedOver, True, start_time, cache_time)
            header(2)
            DisplayCard(NextCard)
            NoOfCardsTurnedOver = NoOfCardsTurnedOver + 1
            Higher = IsNextCardHigher(LastCard, NextCard)
            equal = isnextcardequal(LastCard, NextCard)
            If Higher And Choice = "y" Or Not Higher And Choice = "n" Then
                DisplayCorrectGuessMessage(NoOfCardsTurnedOver - 1)
                globcount = NoOfCardsTurnedOver
                LastCard = NextCard
            ElseIf equal And Choice = "y" Or equal And Choice = "n" Then
                Console.WriteLine("equal value,carry on.")
                LastCard = NextCard
            Else
                GameOver = True
            End If
        End While
        If GameOver Then
            stop_time = Now
            elapsed_time = stop_time.Subtract(start_time)
            globcount = scoretimeratio(NoOfCardsTurnedOver, True, start_time, stop_time)
            speedscore = globcount
            header(2)
            DisplayEndOfGameMessage(NoOfCardsTurnedOver - 2)
            globcount = NoOfCardsTurnedOver
            UpdateRecentScores(RecentScores, globcount, True)
        Else
            DisplayEndOfGameMessage(51)
            UpdateRecentScores(RecentScores, globcount, True)
        End If
    End Sub
    Sub PlayGame(ByVal Deck() As TCard, ByRef RecentScores() As TRecentScore)
        Dim NoOfCardsTurnedOver As Integer
        Dim GameOver As Boolean
        Dim NextCard As TCard
        Dim LastCard As TCard
        Dim Higher As Boolean
        Dim equal As Boolean
        Static start_time As DateTime
        Static cache_time As DateTime
        Static stop_time As DateTime
        Dim elapsed_time As TimeSpan
        Dim speedscore As Integer
        Dim Choice As Char
        GameOver = False
        start_time = Now
        NoOfCardsTurnedOver = 1
        globcount = NoOfCardsTurnedOver
        header(2)
        GetCard(LastCard, Deck, 0)
        DisplayCard(LastCard)
        While NoOfCardsTurnedOver < 52 And Not GameOver
            GetCard(NextCard, Deck, NoOfCardsTurnedOver)
            Do
                Choice = GetChoiceFromUser()
            Loop Until Choice = "y" Or Choice = "n"
            cache_time = Now
            header(2)
            DisplayCard(NextCard)
            NoOfCardsTurnedOver = NoOfCardsTurnedOver + 1
            globcount = NoOfCardsTurnedOver
            Higher = IsNextCardHigher(LastCard, NextCard)
            equal = isnextcardequal(LastCard, NextCard)
            If Higher And Choice = "y" Or Not Higher And Choice = "n" Then
                DisplayCorrectGuessMessage(NoOfCardsTurnedOver - 1)
                globcount = NoOfCardsTurnedOver
                LastCard = NextCard
            ElseIf equal And Choice = "y" Or equal And Choice = "n" Then
                Console.WriteLine("equal value,carry on.")
                LastCard = NextCard
            Else
                GameOver = True
            End If
        End While
        If GameOver Then
            stop_time = Now
            elapsed_time = stop_time.Subtract(start_time)
            globcount = scoretimeratio(NoOfCardsTurnedOver, False, start_time, stop_time)
            speedscore = globcount
            header(2)
            DisplayEndOfGameMessage(NoOfCardsTurnedOver - 2)
            globcount = NoOfCardsTurnedOver
            UpdateRecentScores(RecentScores, NoOfCardsTurnedOver - 2, False)
        Else
            DisplayEndOfGameMessage(51)
            UpdateRecentScores(RecentScores, 51, False)
        End If
    End Sub
End Module