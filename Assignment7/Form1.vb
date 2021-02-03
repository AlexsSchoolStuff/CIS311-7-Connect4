'------------------------------------------------------------
'-                File Name : Form1.frm                     - 
'-                Part of Project: Assignment 7                  -
'------------------------------------------------------------
'-                Written By: Alex Buckstiegel              -
'-                Written On: The date you wrote it         -
'------------------------------------------------------------
'- File Purpose:                                            -
'- Connect 4 game                                           - 
'------------------------------------------------------------
'- Program Purpose:                                         -
'-                                                          -
'- Connect 4 game-
'------------------------------------------------------------
'- Global Variable Dictionary (alphabetically):             -
'- currentButton - current button being evaluated
'- currentPlayer - boolean value of who the current player is
'- column1-7stack - stacks that represent the columns
'- poppedList - list of popped values from the stack. This
'- represents taken buttons
'------------------------------------------------------------




Public Class Form1

    Const Player1 = True
    Const Player2 = False
    Dim currentPlayer As Boolean = True
    Dim column1Stack As Stack = New Stack()
    Dim column2Stack As Stack = New Stack()
    Dim column3Stack As Stack = New Stack()
    Dim column4Stack As Stack = New Stack()
    Dim column5Stack As Stack = New Stack()
    Dim column6Stack As Stack = New Stack()
    Dim column7Stack As Stack = New Stack()
    Dim poppedList As New List(Of Integer)
    Dim currentButton As Button

    '------------------------------------------------------------
    '-                Subprogram Name: Form1_Load               -
    '------------------------------------------------------------
    '-                Written By: Alex Buckstiegel              -
    '-                Written On: 03-22-20                      -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called on form load, and it generates
    '- the stacks and then asks who will go first
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- goesFirst - MsgBox to see who goes first      
    '------------------------------------------------------------

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        listOpenButtons(pnl1, column1Stack)
        listOpenButtons(pnl2, column2Stack)
        listOpenButtons(pnl3, column3Stack)
        listOpenButtons(pnl4, column4Stack)
        listOpenButtons(pnl5, column5Stack)
        listOpenButtons(pnl6, column6Stack)
        listOpenButtons(pnl7, column7Stack)
        Dim goesFirst
        goesFirst = MsgBox("Will Player 1 go first?", vbYesNo, "Choose First Player")
        Select Case goesFirst

            Case 6
                currentPlayer = Player1
                lblWhoseTurn.Text = "Player 1"
                lblPlayer1.Text = "Player 1: Go!"
                btnPlayer2.Enabled = False
            Case 7
                currentPlayer = Player2
                lblWhoseTurn.Text = "Player2"
                lblPlayer2.Text = "Player 2: Go!"
                btnPlayer1.Enabled = False
        End Select

    End Sub

    '------------------------------------------------------------
    '-                Subprogram Name: btnPlayer_MouseDown               -
    '------------------------------------------------------------
    '-                Written By: Alex Buckstiegel              -
    '-                Written On: 03-22-20                      -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called when you press down on either player1
    '- or player 2 button and changes the cursor to the wait cursor
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------

    Private Sub btnPlayer_MouseDown(sender As Object, e As EventArgs) Handles btnPlayer1.MouseDown, btnPlayer2.MouseDown
        btnPlayer1.Capture = False
        btnPlayer2.Capture = False
        Me.Cursor = Cursors.WaitCursor
    End Sub

    '------------------------------------------------------------
    '-                Subprogram Name: btnColumn_MouseEnter               -
    '------------------------------------------------------------
    '-                Written By: Alex Buckstiegel              -
    '-                Written On: 03-22-20                      -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called when your cursor enters the drop buttons
    '- and changes the color if dropping is allowed
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- pnl - Panel to check
    '------------------------------------------------------------

    Private Sub btnColumn_MouseEnter(sender As Object, e As EventArgs) Handles btnColumn1.MouseEnter, btnColumn2.MouseEnter, btnColumn3.MouseEnter, btnColumn4.MouseEnter, btnColumn5.MouseEnter, btnColumn6.MouseEnter, btnColumn7.MouseEnter
        Dim pnl As Panel
        Select Case sender.name
            Case btnColumn1.Name
                pnl = pnl1
            Case btnColumn2.Name
                pnl = pnl2
            Case btnColumn3.Name
                pnl = pnl3
            Case btnColumn4.Name
                pnl = pnl4
            Case btnColumn5.Name
                pnl = pnl5
            Case btnColumn6.Name
                pnl = pnl6
            Case btnColumn7.Name
                pnl = pnl7

        End Select
        If Me.Cursor = Cursors.WaitCursor Then
            If isFull(pnl) Then
                sender.BackColor = Color.Red
            Else
                sender.backColor = Color.Green
            End If
        End If




        Console.WriteLine()
    End Sub


    '------------------------------------------------------------
    '-                Subprogram Name: btnColumn_MouseLeave               -
    '------------------------------------------------------------
    '-                Written By: Alex Buckstiegel              -
    '-                Written On: 03-22-20                      -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called when your cursor leaves the column
    '- buttons, and reverts them back to the original color
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    Private Sub btnColumn1_MouseLeave(sender As Object, e As EventArgs) Handles btnColumn1.MouseLeave, btnColumn2.MouseLeave, btnColumn3.MouseLeave, btnColumn4.MouseLeave, btnColumn5.MouseLeave, btnColumn6.MouseLeave, btnColumn7.MouseLeave
        sender.BackColor = DefaultBackColor()
        sender.UseVisualStyleBackColor = True
    End Sub

    '------------------------------------------------------------
    '-                Subprogram Name: btnColumn_MouseUp        -
    '------------------------------------------------------------
    '-                Written By: Alex Buckstiegel              -
    '-                Written On: 03-22-20                      -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called when the user drops the token
    '- It also resets the colors and then runs the dropToken()
    '- sub to the correct column
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    Private Sub btnColumn1_MouseUp(sender As Object, e As MouseEventArgs) Handles btnColumn1.MouseUp, btnColumn2.MouseUp, btnColumn3.MouseUp, btnColumn4.MouseUp, btnColumn5.MouseUp, btnColumn6.MouseUp, btnColumn7.MouseUp, MyBase.MouseUp, btnReset.MouseUp, lblWinner.MouseUp
        Me.Cursor = Cursors.Default
        For Each control In pnlColumns.Controls
            control.backColor = DefaultBackColor
            control.UseVisualStyleBackColor = True
        Next

        Select Case sender.name
            Case btnColumn1.Name
                dropToken(pnl1, column1Stack)
                changePlayer()
            Case btnColumn2.Name
                dropToken(pnl2, column2Stack)
                changePlayer()
            Case btnColumn3.Name
                dropToken(pnl3, column3Stack)
                changePlayer()
            Case btnColumn4.Name
                dropToken(pnl4, column4Stack)
                changePlayer()
            Case btnColumn5.Name
                dropToken(pnl5, column5Stack)
                changePlayer()
            Case btnColumn6.Name
                dropToken(pnl6, column6Stack)
                changePlayer()
            Case btnColumn7.Name
                dropToken(pnl7, column7Stack)
                changePlayer()
            Case Else


        End Select
    End Sub
    '------------------------------------------------------------
    '-                Subprogram Name: changePlayer               -
    '------------------------------------------------------------
    '-                Written By: Alex Buckstiegel              -
    '-                Written On: 03-22-20                      -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called to change the current player and
    '- update the labels
    '------------------------------------------------------------
    Private Sub changePlayer()
        currentPlayer = Not currentPlayer
        If btnPlayer1.Enabled Then
            lblPlayer2.Text = "Player 2: Go!"
            lblPlayer1.Text = "Player 1"
            btnPlayer2.Enabled = True
            btnPlayer1.Enabled = False
        Else
            lblPlayer1.Text = "Player 1: Go!"
            lblPlayer2.Text = "Player 2!"
            btnPlayer1.Enabled = True
            btnPlayer2.Enabled = False
        End If
    End Sub

    '------------------------------------------------------------
    '-                Subprogram Name: listOpenButtons          -
    '------------------------------------------------------------
    '-                Written By: Alex Buckstiegel              -
    '-                Written On: 03-21-20                      -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called to create the stacks of open 
    '- buttons
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- o - object (panel) to loop through
    '- s - object (stack)
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- arrTabIndex - Array of the tab indexes
    '- counter - counter variable
    '------------------------------------------------------------
    Private Sub listOpenButtons(o As Object, s As Object)
        Dim arrTabIndex(5)
        Dim counter = 0
        For Each control In o.Controls
            arrTabIndex(counter) = control.tabIndex
            counter += 1
        Next
        Array.Sort(arrTabIndex)
        For Each button In arrTabIndex
            s.Push(button)
        Next
        Console.WriteLine("For debugging")
    End Sub
    '------------------------------------------------------------
    '-                Subprogram Name: dropToken               -
    '------------------------------------------------------------
    '-                Written By: Alex Buckstiegel              -
    '-                Written On: 03-22-20                      -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called when you drop a token into a column
    '- it also checks for wins (kinda buggy) because I was having
    '- troubles moving it to a different subroutine
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- o - object (panel) to loop through
    '- s - object (stack)
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- allButtons() - array of all the buttons. changes size depending
    '- on whether it is checking for vertical or horizontal win
    '- checkColor - Checks the correct color for the corresponding
    '- player
    '- counter - counter variable
    '- counter2 - counter variable
    '- strLocation - String of the location of the win
    '- popped - contains the most recent popped value from the stack
    '------------------------------------------------------------
    Private Sub dropToken(o As Object, s As Object)
        If isFull(o) Then
            changePlayer()
            Exit Sub
        End If

        Dim checkColor As Color
        If currentPlayer = Player1 Then
            checkColor = Color.Green
        Else
            checkColor = Color.Yellow
        End If
        Dim strLocation As String


        Dim popped = s.Pop()

        Dim allButtons(5)
        Dim counter = 0

        For Each button In o.controls
            allButtons(counter) = button
            counter += 1
        Next
        allButtons = allButtons.OrderBy(Function(x) x.tabIndex).ToArray
        For Each button In allButtons
            poppedList.Add(popped)
            If poppedList.Last = button.tabIndex Then
                If currentPlayer = Player1 Then
                    button.BackColor = Color.Green
                Else
                    button.BackColor = Color.Yellow
                End If
                currentButton = button
            End If

            If Not poppedList.Contains(button.tabIndex) Then
                If currentPlayer = Player1 Then
                    button.BackColor = Color.Green
                Else
                    button.BackColor = Color.Yellow
                End If

                System.Threading.Thread.Sleep(20)
                Application.DoEvents()

                button.BackColor = DefaultBackColor
                button.UseVisualStyleBackColor = True

            End If
        Next
        'Vertical Win - can utilize allButtons to only check in the same column
        For Each button In allButtons
            If button.tabindex - 7 = currentButton.TabIndex And button.BackColor = checkColor Then
                For Each button2 In allButtons
                    If button2.tabindex - 14 = currentButton.TabIndex And button2.BackColor = checkColor Then
                        For Each button3 In allButtons
                            If button3.tabindex - 21 = currentButton.TabIndex And button3.BackColor = checkColor Then
                                If currentPlayer = Player1 Then
                                    strLocation = o.tag
                                    winner(strLocation, currentPlayer)
                                Else
                                    strLocation = o.tag
                                    winner(strLocation, currentPlayer)
                                End If
                            End If
                        Next
                    End If
                Next
            End If
        Next

        'Horizontal Win -Sadly will have to check every button each time

        ReDim allButtons(41)
        Dim counter2 = 0
        For Each panel In pnlGameBoard.Controls
            For Each button In panel.controls
                allButtons(counter2) = button
                counter2 += 1
            Next
        Next


        For Each button In allButtons
            If button.tabindex + 1 = currentButton.TabIndex And button.BackColor = checkColor Then
                For Each button2 In allButtons
                    If button2.tabindex + 2 = currentButton.TabIndex And button2.BackColor = checkColor Then
                        For Each button3 In allButtons
                            If button3.tabindex + 3 = currentButton.TabIndex And button3.BackColor = checkColor Then
                                If currentPlayer = Player1 Then
                                    strLocation = currentButton.Tag
                                    winner(strLocation, currentPlayer)
                                Else
                                    strLocation = currentButton.Tag
                                    winner(strLocation, currentPlayer)
                                End If
                            End If
                        Next
                    End If
                Next
            End If
        Next

        'Above does not work when the last placed item in on the left side, so this is the same except it checks to the left
        For Each button In allButtons
            If button.tabindex - 1 = currentButton.TabIndex And button.BackColor = checkColor Then
                For Each button2 In allButtons
                    If button2.tabindex - 2 = currentButton.TabIndex And button2.BackColor = checkColor Then
                        For Each button3 In allButtons
                            If button3.tabindex - 3 = currentButton.TabIndex And button3.BackColor = checkColor Then
                                If currentPlayer = Player1 Then
                                    strLocation = currentButton.Tag
                                    winner(strLocation, currentPlayer)
                                Else
                                    strLocation = currentButton.Tag
                                    winner(strLocation, currentPlayer)
                                End If
                            End If
                        Next
                    End If
                Next
            End If
        Next
        If isFull(pnl1) AndAlso isFull(pnl2) AndAlso isFull(pnl3) AndAlso isFull(pnl4) AndAlso isFull(pnl5) AndAlso isFull(pnl6) AndAlso isFull(pnl7) Then
            winner("Tie game!", currentPlayer)
        End If
        Console.WriteLine()
    End Sub

    '------------------------------------------------------------
    '-                Subprogram Name: winner               -
    '------------------------------------------------------------
    '-                Written By: Alex Buckstiegel              -
    '-                Written On: 03-22-20                      -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called when a winner has been detected
    '- and updates the lblWinner.text and disables input
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- location - String of the location saved in the control tag
    '- winner - currentPlayer value of whoever won
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- strWinner - String of the Winner
    '------------------------------------------------------------

    Private Sub winner(location As String, winner As Boolean)
        If location = "Tie game!" Then
            lblWinner.Text = location
            btnPlayer1.Enabled = False
            btnPlayer2.Enabled = False
            Exit Sub
        End If
        Dim strWinner
        If winner = Player1 Then
            strWinner = "Player 1"
        Else
            strWinner = "Player 2"
        End If
        lblWinner.Visible = True

        lblWinner.Text = "Game goes to: " & strWinner & " at " & location
        btnPlayer1.Enabled = False
        btnPlayer2.Enabled = False
    End Sub
    '------------------------------------------------------------
    '-                Subprogram Name: isFull               -
    '------------------------------------------------------------
    '-                Written By: Alex Buckstiegel              -
    '-                Written On: 03-22-20                      -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called to check if a column is full
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- pnl - Panel to check
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- placedCount - if it = 6 then the column is full
    '------------------------------------------------------------
    Private Function isFull(pnl As Panel)
        Dim placedCount = 0
        For Each button In pnl.Controls
            If Not button.BackColor = DefaultBackColor Then
                placedCount += 1
            End If
        Next
        If placedCount = 6 Then
            Return True
        Else
            Return False
        End If
    End Function
    '------------------------------------------------------------
    '-                Subprogram Name: btnReset_Click           -
    '------------------------------------------------------------
    '-                Written By: Alex Buckstiegel              -
    '-                Written On: 03-22-20                      -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called when you click btnReset. It resets
    '- the application
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        'Easiest way to ensure everything is reset and nothing needs to persist so this is the best method
        Application.Restart()
    End Sub
End Class


