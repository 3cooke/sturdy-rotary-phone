Public Class Form1


    Dim NumCards As Integer
    Dim ArLen As Integer
    Dim BingoCard(4, 4) As String                                   'Card Array Writer
    Dim Cardnames() As String                                       'Game Card ID #'s
    Dim CardLayouts() As Array                                      'Array of Card Arrays
    Dim CardHold() As Array                                         'Hold Cards between rounds
    Dim CardWin() As String                                         'Win/Fail holder for cards
    Dim dispCard As ListBox                                         'Userend display of cards

    Public Sub btnWrite_Click(sender As Object, e As EventArgs) Handles btnWrite.Click
        ArLen = NumCards - 1
        ReDim Preserve Cardnames(ArLen)
        ReDim Preserve CardLayouts(ArLen)
        ReDim Preserve CardHold(ArLen)
        ReDim Preserve CardWin(ArLen)

        If cmbCardNum.SelectedItem > 0 Then
            Dim CardNum As Integer = cmbCardNum.SelectedItem - 1
            Dim Spot As Integer
            Dim ErrorCheck As Boolean = False

            BingoCard(0, 0) = txtB1.Text
            BingoCard(0, 1) = txtB2.Text
            BingoCard(0, 2) = txtB3.Text
            BingoCard(0, 3) = txtB4.Text
            BingoCard(0, 4) = txtB5.Text
            BingoCard(1, 0) = txtI1.Text
            BingoCard(1, 1) = txtI2.Text
            BingoCard(1, 2) = txtI3.Text
            BingoCard(1, 3) = txtI4.Text
            BingoCard(1, 4) = txtI5.Text
            BingoCard(2, 0) = txtN1.Text
            BingoCard(2, 1) = txtN2.Text
            BingoCard(2, 2) = 999                              'Free Space
            BingoCard(2, 3) = txtN4.Text
            BingoCard(2, 4) = txtN5.Text
            BingoCard(3, 0) = txtG1.Text
            BingoCard(3, 1) = txtG2.Text
            BingoCard(3, 2) = txtG3.Text
            BingoCard(3, 3) = txtG4.Text
            BingoCard(3, 4) = txtG5.Text
            BingoCard(4, 0) = txtO1.Text
            BingoCard(4, 1) = txtO2.Text
            BingoCard(4, 2) = txtO3.Text
            BingoCard(4, 3) = txtO4.Text
            BingoCard(4, 4) = txtO5.Text

            For B As Integer = 0 To 4
                For I As Integer = 0 To 4
                    If Integer.TryParse(BingoCard(B, I), Spot) Then
                        If Spot < 1 Then
                            MessageBox.Show("Please enter positive integer values only; error at " & B & "," & I)
                            ErrorCheck = True
                            Exit For
                        End If
                    Else
                        MessageBox.Show("Please enter integer values only; error at " & B & "," & I)
                        ErrorCheck = True
                        Exit For
                    End If
                Next
                If ErrorCheck = True Then
                    Exit For
                End If
            Next
            If ErrorCheck = False Then
                BingoCard(2, 2) = 0
                CardLayouts(CardNum) = BingoCard
                CardHold(CardNum) = BingoCard
                Cardnames(CardNum) = txtIDname.Text

                'lstCard1.Items.Add(CardLayouts(0).ToString())
                'lstCard2.Items.Add(CardLayouts(1).ToString())
                'lstCard3.Items.Add(CardLayouts(2).ToString())
                'lstCard4.Items.Add(CardLayouts(3).ToString())
                'lstCard5.Items.Add(CardLayouts(4).ToString())
                'lstCard6.Items.Add(CardLayouts(5).ToString())
                'lstCard7.Items.Add(CardLayouts(6).ToString())
                'lstCard8.Items.Add(CardLayouts(7).ToString())
                'lstCard9.Items.Add(CardLayouts(8).ToString())
                'lstCard10.Items.Add(CardLayouts(9).ToString())
                'lstCard11.Items.Add(CardLayouts(10).ToString())
                'lstCard12.Items.Add(CardLayouts(11).ToString())

                For Each Wipe As Control In grpWriter.Controls
                    If TypeOf Wipe Is TextBox Then
                        Wipe.Text = Nothing
                    End If
                Next
            End If
        Else
            MessageBox.Show("Please select Card number")
        End If



    End Sub

    Private Sub btnCall_Click(sender As Object, e As EventArgs) Handles btnCall.Click

        Dim CalledStr As String = txtCalled.Text
        Dim called As Integer

        If Integer.TryParse(CalledStr, called) Then
            For B As Integer = 0 To 4
                For I As Integer = 0 To 4
                    For Each Card As Array In CardLayouts
                        If Card(B, I) = called Then
                            Card(B, I) = 0
                        End If
                    Next
                Next
            Next

            Dim N As Integer = 0
            For Each CardDisplay As ListBox In Me.Controls
                CardDisplay.Items.Clear()
                CardDisplay.Items.AddRange(CardLayouts(N))
                N += 1
            Next

            'For Each CardDisp As ListBox In CardDisplays
            ' CardDisp.Items.Clear()
            ' CardDisp.Items.AddRange(CardLayouts(CardCounter))
            'CardCounter += 1
            'Next

            If chkStandard.Checked Then
                For Each Card As Array In CardLayouts
                    For I = 0 To ArLen
                        CardWin(I) = StandardBingo(Card)
                        If CardWin(I) = "Bingo" Then
                            MessageBox.Show("Bingo on Card " & Cardnames(I))
                        End If
                    Next
                Next
            End If

            If chkAnyTwo.Checked Then
                For Each Card As Array In CardLayouts
                    For I = 0 To ArLen
                        CardWin(I) = AnyTwoBingo(Card)
                        If CardWin(I) = "Bingo" Then
                            MessageBox.Show("Bingo on Card " & Cardnames(I))
                        End If
                    Next
                Next
            End If

            If chkBlackout.Checked Then
                For Each Card As Array In CardLayouts
                    For I = 0 To ArLen
                        CardWin(I) = BlackoutBingo(Card)
                        If CardWin(I) = "Bingo" Then
                            MessageBox.Show("Bingo on Card " & Cardnames(I))
                        End If
                    Next
                Next
            End If

            If chkPlus.Checked Then
                For Each Card As Array In CardLayouts
                    For I = 0 To ArLen
                        CardWin(I) = PlusBingo(Card)
                        If CardWin(I) = "Bingo" Then
                            MessageBox.Show("Bingo on Card " & Cardnames(I))
                        End If
                    Next
                Next
            End If

            If chkX.Checked Then
                For Each Card As Array In CardLayouts
                    For I = 0 To ArLen
                        CardWin(I) = XBingo(Card)
                        If CardWin(I) = "Bingo" Then
                            MessageBox.Show("Bingo on Card " & Cardnames(I))
                        End If
                    Next
                Next
            End If

            If chkRing.Checked Then
                For Each Card As Array In CardLayouts
                    For I = 0 To ArLen
                        CardWin(I) = RingBingo(Card)
                        If CardWin(I) = "Bingo" Then
                            MessageBox.Show("Bingo on Card " & Cardnames(I))
                        End If
                    Next
                Next
            End If

            If chkTriple.Checked Then
                For Each card As Array In CardLayouts
                    For I = 0 To ArLen
                        CardWin(I) = TripleBingo(card)
                        If CardWin(I) = "Bingo" Then
                            MessageBox.Show("Bingo on Card " & Cardnames(I))
                        End If
                    Next
                Next
            End If

            If chkPyramid.Checked Then
                For Each Card As Array In CardLayouts
                    For I = 0 To ArLen
                        CardWin(I) = PyramidBingo(Card)
                        If CardWin(I) = "Bingo" Then
                            MessageBox.Show("Bingo on Card " & Cardnames(I))
                        End If
                    Next
                Next
            End If
        End If

        txtCalled.Clear()

    End Sub

    Public Function StandardBingo(CheckCard As Array) As String
        Dim I As Integer = 0
        Dim N As Integer = 0
        Dim G As Integer = 0
        Dim O As Integer = 0
        Dim I2 As Integer = 0
        Dim N2 As Integer = 0
        Dim G2 As Integer = 0
        Dim O2 As Integer = 0

        For B As Integer = 0 To 4
            I += CheckCard(B, 0)
            N += CheckCard(B, 1)
            G += CheckCard(B, 3)
            O += CheckCard(B, 4)
            I2 += CheckCard(0, B)
            N2 += CheckCard(1, B)
            G2 += CheckCard(3, B)
            O2 += CheckCard(4, B)
        Next

        If (I Or N Or G Or O Or I2 Or N2 Or G2 Or O2) = 0 Then
            Return "Bingo"
        Else
            Return "Fail"
        End If
    End Function

    Public Function PlusBingo(CheckCard As Array) As String
        Dim I As Integer = 0

        For B As Integer = 0 To 4
            I += CheckCard(2, B) + CheckCard(B, 2)
        Next

        If I = 0 Then
            Return "Bingo"
        Else
            Return "Fail"
        End If
    End Function

    Public Function XBingo(CheckCard As Array) As String
        Dim B As Integer = 0

        For I As Integer = 0 To 4
            B += CheckCard(I, I) + CheckCard(I, (4 - I))
        Next

        If B = 0 Then
            Return "Bingo"
        Else
            Return "Fail"
        End If
    End Function

    Public Function BlackoutBingo(CheckCard As Array) As String
        Dim N As Integer = 0

        For B As Integer = 0 To 4
            For I As Integer = 0 To 4
                N += CheckCard(B, I)
            Next
        Next

        If N = 0 Then
            Return "Bingo"
        Else
            Return "Fail"
        End If
    End Function

    Public Function PyramidBingo(CheckCard As Array) As String
        Dim I As Integer = 0
        Dim N As Integer = 0
        Dim G As Integer = 0
        Dim O As Integer = 0

        For B As Integer = 0 To 4
            I += CheckCard(B, 0)
            N += CheckCard(B, 4)
            G += CheckCard(0, B)
            O += CheckCard(4, B)
            If B <> (0 Or 4) Then
                I += CheckCard(B, 1)
                N += CheckCard(B, 3)
                G += CheckCard(1, B)
                O += CheckCard(3, B)
            End If
        Next

        If (I Or N Or G Or O) = 0 Then
            Return "Bingo"
        Else
            Return "Fail"
        End If

    End Function

    Public Function RingBingo(CheckCard As Array) As String
        Dim I As Integer = 0

        For B As Integer = 0 To 4
            I += CheckCard(B, 0) + CheckCard(B, 4)
            If B <> (0 Or 4) Then
                I += CheckCard(0, B) + CheckCard(4, B)
            End If
        Next

        If I = 0 Then
            Return "Bingo"
        Else
            Return "Fail"
        End If
    End Function

    Public Function AnyTwoBingo(CheckCard As Array) As String
        Dim B As Integer
        Dim I As Integer = 0
        Dim N As Integer = 0
        Dim G As Integer = 0
        Dim O As Integer = 0
        Dim Diag1 As Integer = 0
        Dim Diag2 As Integer = 0
        Dim B2 As Integer = 0
        Dim I2 As Integer = 0
        Dim N2 As Integer = 0
        Dim G2 As Integer = 0
        Dim O2 As Integer = 0

        For C As Integer = 0 To 4
            B += CheckCard(C, 0)
            I += CheckCard(C, 1)
            N += CheckCard(C, 2)
            G += CheckCard(C, 3)
            O += CheckCard(C, 4)
            B2 += CheckCard(0, C)
            I2 += CheckCard(1, C)
            N2 += CheckCard(2, C)
            G2 += CheckCard(3, C)
            O2 += CheckCard(4, C)
            Diag1 += CheckCard(C, C)
            Diag2 += CheckCard(C, (4 - C))
        Next

        If B = 0 Then
            If (I Or N Or G Or O Or B2 Or I2 Or N2 Or G2 Or O2 Or Diag1 Or Diag2) = 0 Then
                Return "Bingo"
            Else
                Return "Fail"
            End If
        ElseIf I = 0 Then
            If (N Or G Or O Or B2 Or I2 Or N2 Or G2 Or O2 Or Diag1 Or Diag2) = 0 Then
                Return "Bingo"
            Else
                Return "Fail"
            End If
        ElseIf N = 0 Then
            If (G Or O Or B2 Or I2 OrElse N2 Or G2 Or O2 Or Diag1 Or Diag2) = 0 Then
                Return "Bingo"
            Else
                Return "Fail"
            End If
        ElseIf G = 0 Then
            If (O Or B2 Or I2 Or N2 Or G2 Or O2 Or Diag1 Or Diag2) = 0 Then
                Return "Bingo"
            Else
                Return "Fail"
            End If
        ElseIf O = 0 Then
            If (B2 Or I2 Or N2 Or G2 Or O2 Or Diag1 Or Diag2) = 0 Then
                Return "Bingo"
            Else
                Return "Fail"
            End If
        ElseIf Diag1 = 0 Then
            If (B2 Or I2 Or N2 Or G2 Or O2 Or Diag2) = 0 Then
                Return "Bingo"
            Else
                Return "Fail"
            End If
        ElseIf Diag2 = 0 Then
            If (B2 Or I2 Or N2 Or G2 Or O2) = 0 Then
                Return "Bingo"
            Else
                Return "Fail"
            End If
        ElseIf B2 = 0 Then
            If (I2 Or N2 Or G2 Or O2) = 0 Then
                Return "Bingo"
            Else
                Return "Fail"
            End If
        ElseIf I2 = 0 Then
            If (N2 Or G2 Or O2) = 0 Then
                Return "Bingo"
            Else
                Return "Fail"
            End If
        ElseIf N2 = 0 Then
            If (G2 Or O2) = 0 Then
                Return "Bingo"
            Else
                Return "Fail"
            End If
        ElseIf (G2 And O2) = 0 Then
            Return "Bingo"
        Else
            Return "Fail"
        End If
    End Function

    Public Function TripleBingo(CheckCard As Array) As String
        Dim B As Integer
        Dim I As Integer = 0
        Dim N As Integer = 0
        Dim G As Integer = 0
        Dim O As Integer = 0
        Dim Diag1 As Integer = 0
        Dim Diag2 As Integer = 0
        Dim B2 As Integer = 0
        Dim I2 As Integer = 0
        Dim N2 As Integer = 0
        Dim G2 As Integer = 0
        Dim O2 As Integer = 0

        For C As Integer = 0 To 4
            B += CheckCard(C, 0)
            I += CheckCard(C, 1)
            N += CheckCard(C, 2)
            G += CheckCard(C, 3)
            O += CheckCard(C, 4)
            B2 += CheckCard(0, C)
            I2 += CheckCard(1, C)
            N2 += CheckCard(2, C)
            G2 += CheckCard(3, C)
            O2 += CheckCard(4, C)
            Diag1 += CheckCard(C, C)
            Diag2 += CheckCard(C, (4 - C))
        Next

        If (Diag1 Or Diag2) = 0 Then
            If (B Or I Or N Or G Or O) = 0 Then
                If (B2 Or I2 Or N2 Or G2 Or O2) = 0 Then
                    Return "Bingo"
                Else
                    Return "Fail"
                End If
            Else
                Return "Fail"
            End If
        Else
            Return "Fail"
        End If
    End Function

    Private Sub btnNewRound_Click(sender As Object, e As EventArgs) Handles btnNewRound.Click
        Dim Warning As DialogResult = MessageBox.Show("New Round?", "Confirm new round", MessageBoxButtons.OKCancel)

        If Warning = DialogResult.OK Then
            For B = 0 To ArLen
                CardLayouts(B) = CardHold(B)
            Next
        ElseIf Warning = DialogResult.Cancel Then
        End If
    End Sub

    Public Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'CardDisplays.Add(lstCard1)
        'CardDisplays.Add(lstCard2)
        'CardDisplays.Add(lstCard3)
        'CardDisplays.Add(lstCard4)
        'CardDisplays.Add(lstCard5)
        'CardDisplays.Add(lstCard6)
        'CardDisplays.Add(lstCard7)
        'CardDisplays.Add(lstCard8)
        'CardDisplays.Add(lstCard9)
        'CardDisplays.Add(lstCard10)
        'CardDisplays.Add(lstCard11)
        'CardDisplays.Add(lstCard12)

        Do
            Dim InputStr As String = InputBox("How many Bingo cards are you playing?", "Number of cards?", 3)
            Dim input As Integer

            If Integer.TryParse(InputStr, input) Then
                If input > 0 And input < 13 Then
                    NumCards = input
                Else
                    MessageBox.Show("Please enter an integer from 1 to 12")
                End If
            Else
                MessageBox.Show("Please enter an integer from 1 to 12")
            End If
        Loop Until NumCards <> 0

        For I = 1 To NumCards
            cmbCardNum.Items.Add(I)
            Me.dispCard = New System.Windows.Forms.ListBox
            Me.dispCard.FormattingEnabled = True
            Me.dispCard.Name = "lstCard" & I
            Me.dispCard.Size = New System.Drawing.Size(159, 173)
            Select Case I
                Case 1 - 4
                    Me.dispCard.Location = New System.Drawing.Point((37 + I * 175), 34)
                Case 5 - 8
                    Me.dispCard.Location = New System.Drawing.Point((37 + I * 175), 226)
                Case 9 - 12
                    Me.dispCard.Location = New System.Drawing.Point((37 + I * 175), 411)
            End Select
        Next

        'CardDisplays.RemoveRange((NumCards - 1), (12 - NumCards))

    End Sub
End Class
