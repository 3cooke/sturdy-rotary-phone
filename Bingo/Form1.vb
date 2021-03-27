Public Class Form1
    Dim NumCards, ArLen As Integer
    Dim BingoCard(4, 4) As String                                 'Card Array Writer
    Dim Cardnames() As String                                     'Game Card ID #'s
    Dim CardLayouts(4, 4, 15) As Integer                          'Array of Card Arrays
    Dim CardHold(4, 4, 15) As Integer                             'Hold Cards between rounds
    'Dim Custom(4, 4) As String                                    'Custom Win Condition writer
    'BUTTON CONTROLS
    Public Sub btnWrite_Click(sender As Object, e As EventArgs) Handles btnWrite.Click
        If cmbCardNum.SelectedItem <> -1 Then
            Dim CardNum As Integer = cmbCardNum.SelectedIndex
            Dim CheckSpot As Integer
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
                    If Integer.TryParse(BingoCard(B, I), CheckSpot) Then
                        If CheckSpot < 1 Then
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
                For B As Integer = 0 To 4
                    For I As Integer = 0 To 4
                        CardLayouts(B, I, CardNum) = BingoCard(B, I)
                        CardHold(B, I, CardNum) = CardLayouts(B, I, CardNum)
                    Next
                Next
                Cardnames(CardNum) = txtIDname.Text
                WriteDisplays()
                For Each Wipe As Control In grpWriter.Controls
                    If TypeOf Wipe Is TextBox Then
                        Wipe.Text = Nothing
                    End If
                Next
                txtB1.Focus()
            End If
        Else
            MessageBox.Show("Please select Card number")
        End If
    End Sub
    Private Sub btnCall_Click(sender As Object, e As EventArgs) Handles btnCall.Click
        Dim called As Integer
        If Integer.TryParse(txtCalled.Text, called) Then
            CallCheck(called)
            WriteDisplays()
            CheckWins()
            txtCalled.Clear()
        Else
            MessageBox.Show("Enter called number only - no letters")
        End If
        txtCalled.Focus()
    End Sub
    Private Sub btnNewRound_Click(sender As Object, e As EventArgs) Handles btnNewRound.Click
        Dim Warning As DialogResult = MessageBox.Show("New Round?", "Confirm new round", MessageBoxButtons.OKCancel)
        If Warning = DialogResult.OK Then
            For B = 0 To 4
                For I = 0 To 4
                    For N = 0 To ArLen
                        CardLayouts(B, I, N) = CardHold(B, I, N)
                    Next
                Next
            Next
            WriteDisplays()
        ElseIf Warning = DialogResult.Cancel Then
        End If
        txtCalled.Focus()
    End Sub
    'FORM LOAD, FUNCTIONS, SUBROUTINES
    Public Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Do
            Dim InputStr As String = InputBox("How many Bingo cards are you playing?", "Number of cards?", 3)
            Dim input As Integer
            If InputStr = DialogResult.Cancel Then
                Me.Close()
            End If
            If Integer.TryParse(InputStr, input) Then
                If input > 0 And input < 17 Then
                    NumCards = input
                Else
                    MessageBox.Show("Please enter an integer from 1 to 16")
                End If
            Else
                MessageBox.Show("Please enter an integer from 1 to 16")
            End If
        Loop Until NumCards <> 0
        For I = 1 To NumCards
            cmbCardNum.Items.Add(I)
        Next
        'Dim Custom As DialogResult = MessageBox.Show("Is there a custom Win Condition?", "Customs check", MessageBoxButtons.YesNo)
        'If Custom = DialogResult.Yes Then
        'cmbCardNum.Items.Add("CustomWin")
        'chkCustom.Visible = True
        'End If
        ArLen = (NumCards - 1)
        ReDim Preserve Cardnames(0 To ArLen)
        ReDim Preserve CardLayouts(4, 4, ArLen)
        ReDim Preserve CardHold(4, 4, ArLen)
    End Sub
    Sub CheckWins()
        If chkStandard.Checked Then
            StandardBingo()
        End If
        If chkAnyTwo.Checked Then
            AnyTwoBingo()
        End If
        If chkBlackout.Checked Then
            BlackoutBingo()
        End If
        If chkPlus.Checked Then
            PlusBingo()
        End If
        If chkX.Checked Then
            XBingo()
        End If
        If chkRing.Checked Then
            RingBingo()
        End If
        If chkTriple.Checked Then
            TripleBingo()
        End If
        If chkPyramid.Checked Then
            PyramidBingo()
        End If
    End Sub
    Sub CallCheck(called As Integer)
        For B As Integer = 0 To 4
            For I As Integer = 0 To 4
                For N As Integer = 0 To ArLen
                    If CardLayouts(B, I, N) = called Then
                        CardLayouts(B, I, N) = 0
                    End If
                Next
            Next
        Next
    End Sub
    Sub WriteDisplays()
        For G As Integer = 0 To ArLen
            If G = 0 Then
                DisplayWriter(lstCard1, G)
            ElseIf G = 1 Then
                DisplayWriter(lstCard2, G)
            ElseIf G = 2 Then
                DisplayWriter(lstCard3, G)
            ElseIf G = 3 Then
                DisplayWriter(lstCard4, G)
            ElseIf G = 4 Then
                DisplayWriter(lstCard5, G)
            ElseIf G = 5 Then
                DisplayWriter(lstCard6, G)
            ElseIf G = 6 Then
                DisplayWriter(lstCard7, G)
            ElseIf G = 7 Then
                DisplayWriter(lstCard8, G)
            ElseIf G = 8 Then
                DisplayWriter(lstCard9, G)
            ElseIf G = 9 Then
                DisplayWriter(lstCard10, G)
            ElseIf G = 10 Then
                DisplayWriter(lstCard11, G)
            ElseIf G = 11 Then
                DisplayWriter(lstCard12, G)
            ElseIf G = 12 Then
                DisplayWriter(lstCard13, G)
            ElseIf G = 13 Then
                DisplayWriter(lstCard14, G)
            ElseIf G = 14 Then
                DisplayWriter(lstCard15, G)
            ElseIf G = 15 Then
                DisplayWriter(lstCard16, G)
            End If
        Next
    End Sub
    Sub DisplayWriter(box As ListBox, cardnum As Integer)
        box.Visible = True
        box.Items.Clear()
        box.Items.Add("B" & vbTab & "I" & vbTab & "N" & vbTab & "G" & vbTab & "O")
        For N As Integer = 0 To 4
            box.Items.Add(CardLayouts(0, N, cardnum) & vbTab & CardLayouts(1, N, cardnum) & vbTab &
            CardLayouts(2, N, cardnum) & vbTab & CardLayouts(3, N, cardnum) & vbTab & CardLayouts(4, N, cardnum))
        Next
        box.Items.Add(" ")
        box.Items.Add("Card ID: " & Cardnames(cardnum))
    End Sub
    'WIN CONDITIONS
    Sub StandardBingo()
        Dim I, N, G, O, I2, N2, G2, O2 As Integer
        For B As Integer = 0 To ArLen
            For B2 As Integer = 0 To 4
                I += CardLayouts(B2, 0, B)
                N += CardLayouts(B2, 1, B)
                G += CardLayouts(B2, 3, B)
                O += CardLayouts(B2, 4, B)
                I2 += CardLayouts(0, B2, B)
                N2 += CardLayouts(1, B2, B)
                G2 += CardLayouts(3, B2, B)
                O2 += CardLayouts(4, B2, B)
            Next
            If I = 0 Or N = 0 Or G = 0 Or O = 0 Or I2 = 0 Or N2 = 0 Or G2 = 0 Or O2 = 0 Then
                MessageBox.Show("Bingo on card " & Cardnames(B))
            End If
            I = 0
            N = 0
            G = 0
            O = 0
            I2 = 0
            N2 = 0
            G2 = 0
            O2 = 0
        Next
    End Sub
    Sub PlusBingo()
        Dim N As Integer
        For B As Integer = 0 To ArLen
            For I As Integer = 0 To 4
                N += CardLayouts(2, I, B) + CardLayouts(I, 2, B)
            Next
            If N = 0 Then
                MessageBox.Show("Bingo on card " & Cardnames(B))
            End If
            N = 0
        Next
    End Sub
    Sub XBingo()
        Dim N As Integer
        For B = 0 To ArLen
            For I As Integer = 0 To 4
                N += CardLayouts(I, I, B) + CardLayouts(I, (4 - I), B)
            Next
            If N = 0 Then
                MessageBox.Show("Bingo on card " & Cardnames(B))
            End If
            N = 0
        Next
    End Sub
    Sub BlackoutBingo()
        Dim G As Integer = 0
        For B As Integer = 0 To ArLen
            For I As Integer = 0 To 4
                For N As Integer = 0 To 4
                    G += CardLayouts(I, N, B)
                Next
            Next
            If G = 0 Then
                MessageBox.Show("Bingo on card " & Cardnames(B))
            End If
            G = 0
        Next
    End Sub
    Sub PyramidBingo()
        Dim I, N, G, O As Integer
        For card As Integer = 0 To ArLen
            For B As Integer = 0 To 4
                I += CardLayouts(B, 0, card)
                N += CardLayouts(B, 4, card)
                G += CardLayouts(0, B, card)
                O += CardLayouts(4, B, card)
                If B = 1 Or 2 Or 3 Then
                    I += CardLayouts(B, 1, card)
                    N += CardLayouts(B, 3, card)
                    G += CardLayouts(1, B, card)
                    O += CardLayouts(3, B, card)
                End If
            Next
            If (I Or N Or G Or O) = 0 Then
                MessageBox.Show("Bingo on card " & Cardnames(card))
            End If
            I = 0
            N = 0
            G = 0
            O = 0
        Next
    End Sub
    Sub RingBingo()
        Dim I As Integer
        For card As Integer = 0 To ArLen
            For B As Integer = 0 To 4
                I += CardLayouts(B, 0, card) + CardLayouts(B, 4, card)
                If B <> (0 Or 4) Then
                    I += CardLayouts(0, B, card) + CardLayouts(4, B, card)
                End If
            Next
            If I = 0 Then
                MessageBox.Show("Bingo on card " & Cardnames(card))
            End If
            I = 0
        Next
    End Sub
    Sub AnyTwoBingo()
        Dim B, I, N, G, O, Diag1, Diag2, B2, I2, N2, G2, O2 As Integer
        For card As Integer = 0 To ArLen
            For C As Integer = 0 To 4
                B += CardLayouts(C, 0, card)
                I += CardLayouts(C, 1, card)
                N += CardLayouts(C, 2, card)
                G += CardLayouts(C, 3, card)
                O += CardLayouts(C, 4, card)
                B2 += CardLayouts(0, C, card)
                I2 += CardLayouts(1, C, card)
                N2 += CardLayouts(2, C, card)
                G2 += CardLayouts(3, C, card)
                O2 += CardLayouts(4, C, card)
                Diag1 += CardLayouts(C, C, card)
                Diag2 += CardLayouts(C, (4 - C), card)
            Next
            If B = 0 Then
                If I = 0 Or N = 0 Or G = 0 Or O = 0 Or B2 = 0 Or I2 = 0 Or N2 = 0 Or G2 = 0 Or O2 = 0 Or Diag1 = 0 Or Diag2 = 0 Then
                    MessageBox.Show("Bingo on card " & Cardnames(card))
                End If
            ElseIf I = 0 Then
                If N = 0 Or G = 0 Or O = 0 Or B2 = 0 Or I2 = 0 Or N2 = 0 Or G2 = 0 Or O2 = 0 Or Diag1 = 0 Or Diag2 = 0 Then
                    MessageBox.Show("Bingo on card " & Cardnames(card))
                End If
            ElseIf N = 0 Then
                If G = 0 Or O = 0 Or B2 = 0 Or I2 = 0 Or N2 = 0 Or G2 = 0 Or O2 = 0 Or Diag1 = 0 Or Diag2 = 0 Then
                    MessageBox.Show("Bingo on card " & Cardnames(card))
                End If
            ElseIf G = 0 Then
                If O = 0 Or B2 = 0 Or I2 = 0 Or N2 = 0 Or G2 = 0 Or O2 = 0 Or Diag1 = 0 Or Diag2 = 0 Then
                    MessageBox.Show("Bingo on card " & Cardnames(card))
                End If
            ElseIf O = 0 Then
                If B2 = 0 Or I2 = 0 Or N2 = 0 Or G2 = 0 Or O2 = 0 Or Diag1 = 0 Or Diag2 = 0 Then
                    MessageBox.Show("Bingo on card " & Cardnames(card))
                End If
            ElseIf Diag1 = 0 Then
                If B2 = 0 Or I2 = 0 Or N2 = 0 Or G2 = 0 Or O2 = 0 Or Diag2 = 0 Then
                    MessageBox.Show("Bingo on card " & Cardnames(card))
                End If
            ElseIf Diag2 = 0 Then
                If B2 = 0 Or I2 = 0 Or N2 = 0 Or G2 = 0 Or O2 = 0 Then
                    MessageBox.Show("Bingo on card " & Cardnames(card))
                End If
            ElseIf B2 = 0 Then
                If I2 = 0 Or N2 = 0 Or G2 = 0 Or O2 = 0 Then
                    MessageBox.Show("Bingo on card " & Cardnames(card))
                End If
            ElseIf I2 = 0 Then
                If N2 = 0 Or G2 = 0 Or O2 = 0 Then
                    MessageBox.Show("Bingo on card " & Cardnames(card))
                End If
            ElseIf N2 = 0 Then
                If G2 = 0 Or O2 = 0 Then
                    MessageBox.Show("Bingo on card " & Cardnames(card))
                End If
            ElseIf (G2 And O2) = 0 Then
                MessageBox.Show("Bingo on card " & Cardnames(card))
            End If
            B = 0
            I = 0
            N = 0
            G = 0
            O = 0
            B2 = 0
            I2 = 0
            N2 = 0
            G2 = 0
            O2 = 0
            Diag1 = 0
            Diag2 = 0
        Next
    End Sub
    Sub TripleBingo()
        Dim B, I, N, G, O, Diag1, Diag2, B2, I2, N2, G2, O2 As Integer
        For card As Integer = 0 To ArLen
            For C As Integer = 0 To 4
                B += CardLayouts(C, 0, card)
                I += CardLayouts(C, 1, card)
                N += CardLayouts(C, 2, card)
                G += CardLayouts(C, 3, card)
                O += CardLayouts(C, 4, card)
                B2 += CardLayouts(0, C, card)
                I2 += CardLayouts(1, C, card)
                N2 += CardLayouts(2, C, card)
                G2 += CardLayouts(3, C, card)
                O2 += CardLayouts(4, C, card)
                Diag1 += CardLayouts(C, C, card)
                Diag2 += CardLayouts(C, (4 - C), card)
            Next
            If Diag1 = 0 Or Diag2 = 0 Then
                If B = 0 Or I = 0 Or N = 0 Or G = 0 Or O = 0 Then
                    If B2 = 0 Or I2 = 0 Or N2 = 0 Or G2 = 0 Or O2 = 0 Then
                        MessageBox.Show("Bingo on card " & Cardnames(card))
                    End If
                End If
            End If
            B = 0
            I = 0
            N = 0
            G = 0
            O = 0
            B2 = 0
            I2 = 0
            N2 = 0
            G2 = 0
            O2 = 0
            Diag1 = 0
            Diag2 = 0
        Next
    End Sub
End Class
