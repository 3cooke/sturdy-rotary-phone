Public Class Form1


    Dim BingoCard(4, 4) As Integer          'Actual Card Array
    Dim CardLayouts(11) As Array            'Array of Card Arrays
    Dim Cardnames(11) As String             'Game Card ID #'s
    Dim CardHold As Array                   'Hold Cards between rounds
    Dim CardWin(11) As String

    Public Sub btnWrite_Click(sender As Object, e As EventArgs) Handles btnWrite.Click

        BingoCard(0, 0) = Val(txtB1.Text)
        BingoCard(0, 1) = Val(txtB2.Text)
        BingoCard(0, 2) = Val(txtB3.Text)
        BingoCard(0, 3) = Val(txtB4.Text)
        BingoCard(0, 4) = Val(txtB5.Text)
        BingoCard(1, 0) = Val(txtI1.Text)
        BingoCard(1, 1) = Val(txtI2.Text)
        BingoCard(1, 2) = Val(txtI3.Text)
        BingoCard(1, 3) = Val(txtI4.Text)
        BingoCard(1, 4) = Val(txtI5.Text)
        BingoCard(2, 0) = Val(txtN1.Text)
        BingoCard(2, 1) = Val(txtN2.Text)
        BingoCard(2, 2) = 0                              'Free Space
        BingoCard(2, 3) = Val(txtN4.Text)
        BingoCard(2, 4) = Val(txtN5.Text)
        BingoCard(3, 0) = Val(txtG1.Text)
        BingoCard(3, 2) = Val(txtG3.Text)
        BingoCard(3, 3) = Val(txtG4.Text)
        BingoCard(3, 4) = Val(txtG5.Text)
        BingoCard(4, 0) = Val(txtO1.Text)
        BingoCard(4, 1) = Val(txtO2.Text)
        BingoCard(4, 2) = Val(txtO3.Text)
        BingoCard(4, 3) = Val(txtO4.Text)
        BingoCard(4, 4) = Val(txtO5.Text)

        CardLayouts((cmbCardNum.SelectedItem - 1)) = BingoCard
        CardHold((cmbCardNum.SelectedItem - 1)) = BingoCard

        Cardnames((cmbCardNum.SelectedItem - 1)) = txtIDname.Text


        Dim Wipe As Control                     'Clear textboxes when card written

        For Each Wipe In grpWriter.Controls
            If TypeOf Wipe Is TextBox Then
                Wipe.Text = Nothing
            End If
        Next

    End Sub

    Private Sub btnCall_Click(sender As Object, e As EventArgs) Handles btnCall.Click
        Dim B As Integer
        Dim I As Integer = 0
        Dim Called As Integer = Val(txtCalled.Text)
        Dim CardHitTemp(4, 4) As Integer

        For B = 0 To 4
            For I = 0 To 4
                For Each Card As Array In CardLayouts
                    If Called = Val(Card(I, B)) Then
                        Card(I, B) = 0
                    End If
                Next
            Next
        Next

        If chkStandard.Checked Then
            For Each Card As Array In CardLayouts
                For I = 0 To 11
                    CardWin(I) = StandardBingo(Card)
                    If CardWin(I) = "Bingo" Then
                        MessageBox.Show("Bingo on Card " & Cardnames(I))
                    End If
                Next
            Next
        End If

        If chkAnyTwo.Checked Then
            For Each Card As Array In CardLayouts
                For I = 0 To 11
                    CardWin(I) = AnyTwoBingo(Card)
                    If CardWin(I) = "Bingo" Then
                        MessageBox.Show("Bingo on Card " & Cardnames(I))
                    End If
                Next
            Next
        End If

        If chkBlackout.Checked Then
            For Each Card As Array In CardLayouts
                For I = 0 To 11
                    CardWin(I) = BlackoutBingo(Card)
                    If CardWin(I) = "Bingo" Then
                        MessageBox.Show("Bingo on Card " & Cardnames(I))
                    End If
                Next
            Next
        End If

        If chkPlus.Checked Then
            For Each Card As Array In CardLayouts
                For I = 0 To 11
                    CardWin(I) = PlusBingo(Card)
                    If CardWin(I) = "Bingo" Then
                        MessageBox.Show("Bingo on Card " & Cardnames(I))
                    End If
                Next
            Next
        End If

        If chkX.Checked Then
            For Each Card As Array In CardLayouts
                For I = 0 To 11
                    CardWin(I) = XBingo(Card)
                    If CardWin(I) = "Bingo" Then
                        MessageBox.Show("Bingo on Card " & Cardnames(I))
                    End If
                Next
            Next
        End If

        If chkRing.Checked Then
            For Each Card As Array In CardLayouts
                For I = 0 To 11
                    CardWin(I) = RingBingo(Card)
                    If CardWin(I) = "Bingo" Then
                        MessageBox.Show("Bingo on Card " & Cardnames(I))
                    End If
                Next
            Next
        End If

        If chkTriple.Checked Then
            For Each card As Array In CardLayouts
                For I = 0 To 11
                    CardWin(I) = TripleBingo(card)
                    If CardWin(I) = "Bingo" Then
                        MessageBox.Show("Bingo on Card " & Cardnames(I))
                    End If
                Next
            Next
        End If

        If chkPyramid.Checked Then
            For Each Card As Array In CardLayouts
                For I = 0 To 11
                    CardWin(I) = PyramidBingo(Card)
                    If CardWin(I) = "Bingo" Then
                        MessageBox.Show("Bingo on Card " & Cardnames(I))
                    End If
                Next
            Next
        End If

        txtCalled.Clear()

    End Sub

    Public Function StandardBingo(CheckCard As Array) As String
        Dim B As Integer = 0
        Dim I As Integer = 0
        Dim N As Integer = 0
        Dim G As Integer = 0
        Dim O As Integer = 0
        Dim I2 As Integer = 0
        Dim N2 As Integer = 0
        Dim G2 As Integer = 0
        Dim O2 As Integer = 0

        For B = 0 To 4
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
        Dim B As Integer = 0
        Dim I As Integer = 0

        For B = 0 To 4
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

        For I = 0 To 4
            B += CheckCard(I, I) + CheckCard(I, (4 - I))
        Next

        If B = 0 Then
            Return "Bingo"
        Else
            Return "Fail"
        End If
    End Function

    Public Function BlackoutBingo(CheckCard As Array) As String
        Dim B As Integer = 0
        Dim I As Integer = 0
        Dim N As Integer = 0

        For B = 0 To 4
            For I = 0 To 4
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
        Dim B As Integer = 0
        Dim I As Integer = 0
        Dim N As Integer = 0
        Dim G As Integer = 0
        Dim O As Integer = 0

        For B = 0 To 4
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
        Dim B As Integer
        Dim I As Integer

        For B = 0 To 4
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
        Dim C As Integer
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

        For C = 0 To 4
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
        Dim C As Integer
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

        For C = 0 To 4
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
        Dim ClearHits(4, 4) As Integer

        If Warning = DialogResult.OK Then
            For B = 0 To 11
                CardLayouts(B) = CardHold(B)
            Next
        ElseIf Warning = DialogResult.Cancel Then
        End If
    End Sub

End Class
