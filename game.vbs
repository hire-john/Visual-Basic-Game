
'Initialization Section

Option Explicit

Const cGreetingMsg = "Pick a number between 1 - 100"

'original app variables
Dim intUserNumber, intRandomNo, strOkToEnd, intNoGuesses

'declared the following variables to help us calculate hot/cold distances
Dim intPreviousGuess, intPreviousDistance, intCurrentGuess, intCurrentDistance

'initializing our variables to zero
intNoGuesses = 0
intPreviousGuess = 0 
intPreviousDistance = 0
intCurrentGuess = 0
intCurrentDistance = 0


'Main Processing Section

'Generate a random number
Randomize
intRandomNo = FormatNumber(Int((100 * Rnd) + 1))
MsgBox intRandomNo

'Loop until either the user guesses correctly or the user clicks on Cancel
Do Until strOkToEnd = "yes"

  'Prompt user to pick a number
  intUserNumber = InputBox("Type your guess:",cGreetingMsg)
  intUserNumber = FormatNumber(intUserNumber)
  intNoGuesses = intNoGuesses + 1
  
  'See if the user provided an answer
  If Len(intUserNumber) <> 0 Then

    'Make sure that the player typed a number
    If IsNumeric(intUserNumber) = True Then
    
     'wrap the initial conditionals with another conditional to check the guess count 
     'if its our first guess there will be no existing distance.
     If intNoGuesses = 1 Then
      
      'Test to see if the user's guess was too low
      If intUserNumber < intRandomNo Then
        MsgBox "Your guess was too low. Try again", ,cGreetingMsg
        strOkToEnd = "no"
        intPreviousGuess = intUserNumber 'store our first initial guess
      End If
      
      'Test to see if the user's guess was too high
      If intUserNumber > intRandomNo Then
        MsgBox "Your guess was too high. Try again", ,cGreetingMsg
        strOkToEnd = "no"
        intPreviousGuess = intUserNumber 'store our first initial guess
      End If

      'Test to see if the user's guess was correct
      If intUserNumber = intRandomNo Then
        MsgBox "Congratulations! You guessed it on the first try! your special! The number was " & _
          intUserNumber & "." & vbCrLf & vbCrLf & "You guessed it " & _
          "in " & intNoGuesses & " guesses.", ,cGreetingMsg
        strOkToEnd = "yes"
      End If

    ElseIf intNoGuesses > 1 Then

        intCurrentGuess = intUserNumber 'intentionally chose to use additional memory space for clarity
        intCurrentDistance = Abs(FormatNumber(IntRandomNo - IntCurrentGuess)) 'get the absolute value of the number to determine its actual distance.
        intPreviousDistance = Abs(FormatNumber(IntRandomNo - intPreviousGuess)) 'get the absolute value of the number to determine its actual distance.
        
        If intCurrentDistance > intPreviousDistance Then
           
           'you can remove the following conditional and the game will work better
            If intCurrentDistance <= 80 And intCurrentDistance >= 61 Then
               MsgBox "Your really, really, really, really cold!"
            ElseIf intCurrentDistance <= 60 And intCurrentDistance >= 41 Then
               MsgBox "Your really, really, really cold!"
            ElseIf intCurrentDistance <= 40 And intCurrentDistance >= 21 Then
               MsgBox "Your really, really cold!"
            ElseIf intCurrentDistance <= 20 And intCurrentDistance >= 11 Then
               MsgBox "Your really cold!"
            ElseIf intCurrentDistance <= 10 And intCurrentDistance >= 1 Then
               MsgBox "Your cold!"
            End If
            'stop remove

            strOkToEnd = "no"
            intPreviousGuess = intCurrentGuess 'set the previous guess equal to the current guess to check in the next iteration
        ElseIf intCurrentDistance < intPreviousDistance Then
            
            'you can remove the following conditional and the game will work better
            If intCurrentDistance >= 80 And intCurrentDistance >= 61 Then
                MsgBox "Your Hot!"
            ElseIf intCurrentDistance <= 60 And intCurrentDistance >= 41 Then
                MsgBox "Your really Hot!"
            ElseIf intCurrentDistance <= 40 And intCurrentDistance >= 21 Then
                MsgBox "Your really, really Hot!"
            ElseIf intCurrentDistance <= 20 And intCurrentDistance >= 11 Then
                MsgBox "Your really, really, really Hot!"
            ElseIf intCurrentDistance <= 10 And intCurrentDistance >= 1 Then
                MsgBox "Your on fire!"
            End If
            'stop remove


            strOkToEnd = "no"
            intPreviousGuess = intCurrentGuess 'set the previous guess equal to the current guess to check in the next iteration
        Else
            MsgBox "Your didn't change your guess!", ,cGreetingMsg
            strOkToEnd = "no"
            intPreviousGuess = intCurrentGuess 'set the previous guess equal to the current guess to check in the next iteration
        End If

        'Test to see if the user's guess was correct
        If intUserNumber = intRandomNo And intNoGuesses < 99 Then
            MsgBox "Congratulations! You guessed it pretty quick. The number was " & _
            intUserNumber & "." & vbCrLf & vbCrLf & "You guessed it " & _
            "in " & intNoGuesses & " guesses.", ,cGreetingMsg
            strOkToEnd = "yes"
        ElseIf intUserNumber = intRandomNo And intNoGuesses > 99 Then
            MsgBox "Your just plain unlucky. The number was " & _
            intUserNumber & "." & vbCrLf & vbCrLf & "You guessed it " & _
            "in " & intNoGuesses & " guesses.", ,cGreetingMsg
            strOkToEnd = "yes"
        End If

    Else
      MsgBox "Sorry. You did not enter a number. Try again.", , cGreetingMsg
    End If
  Else
    MsgBox "You either failed to type a value or you clicked on Cancel. " & _
      "Please play again soon!", , cGreetingMsg
    strOkToEnd = "yes"
  End If
 End If
Loop
WScript.Quit