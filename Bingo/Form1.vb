Public Class Form1
    Dim Card1(4, 4) As Integer
    Dim Card1Hits(4, 4) As Integer
    Dim Card2(4, 4) As Integer
    Dim Card2Hits(4, 4) As Integer
    Dim Card3(4, 4) As Integer
    Dim Card3Hits(4, 4) As Integer
    Dim Card4(4, 4) As Integer
    Dim Card4Hits(4, 4) As Integer
    Dim Card5(4, 4) As Integer
    Dim Card5Hits(4, 4) As Integer
    Dim Card6(4, 4) As Integer
    Dim Card6Hits(4, 4) As Integer
    Dim Card7(4, 4) As Integer
    Dim Card7Hits(4, 4) As Integer
    Dim Card8(4, 4) As Integer
    Dim Card8Hits(4, 4) As Integer
    Dim Card9(4, 4) As Integer
    Dim Card9Hits(4, 4) As Integer
    Dim Card10(4, 4) As Integer
    Dim Card10Hits(4, 4) As Integer
    Dim Card11(4, 4) As Integer
    Dim Card11Hits(4, 4) As Integer
    Dim Card12(4, 4) As Integer
    Dim Card12Hits(4, 4) As Integer

    Dim Cardnames(11) As String
    Dim CardWin(11) As String

    Public Sub btnWrite_Click(sender As Object, e As EventArgs) Handles btnWrite.Click
        Dim CardTemp(4, 4) As Integer

        CardTemp(0, 0) = Val(txtB1.Text)
        CardTemp(0, 1) = Val(txtB2.Text)
        CardTemp(0, 2) = Val(txtB3.Text)
        CardTemp(0, 3) = Val(txtB4.Text)
        CardTemp(0, 4) = Val(txtB5.Text)
        CardTemp(1, 0) = Val(txtI1.Text)
        CardTemp(1, 1) = Val(txtI2.Text)
        CardTemp(1, 2) = Val(txtI3.Text)
        CardTemp(1, 3) = Val(txtI4.Text)
        CardTemp(1, 4) = Val(txtI5.Text)
        CardTemp(2, 0) = Val(txtN1.Text)
        CardTemp(2, 1) = Val(txtN2.Text)
        CardTemp(2, 2) = 0                              'Free Space
        CardTemp(2, 3) = Val(txtN4.Text)
        CardTemp(2, 4) = Val(txtN5.Text)
        CardTemp(3, 0) = Val(txtG1.Text)
        CardTemp(3, 2) = Val(txtG3.Text)
        CardTemp(3, 3) = Val(txtG4.Text)
        CardTemp(3, 4) = Val(txtG5.Text)
        CardTemp(4, 0) = Val(txtO1.Text)
        CardTemp(4, 1) = Val(txtO2.Text)
        CardTemp(4, 2) = Val(txtO3.Text)
        CardTemp(4, 3) = Val(txtO4.Text)
        CardTemp(4, 4) = Val(txtO5.Text)

        If cmbCardNum.SelectedItem = 1 Then
            Array.Copy(CardTemp, Card1, 25)
            Cardnames(1) = txtIDname.Text
        ElseIf cmbCardNum.SelectedItem = 2 Then
            Array.Copy(CardTemp, Card2, 25)
            Cardnames(2) = txtIDname.Text
        ElseIf cmbCardNum.SelectedItem = 3 Then
            Array.Copy(CardTemp, Card3, 25)
            Cardnames(3) = txtIDname.Text
        ElseIf cmbCardNum.SelectedItem = 4 Then
            Array.Copy(CardTemp, Card4, 25)
            Cardnames(4) = txtIDname.Text
        ElseIf cmbCardNum.SelectedItem = 5 Then
            Array.Copy(CardTemp, Card5, 25)
            Cardnames(5) = txtIDname.Text
        ElseIf cmbCardNum.SelectedItem = 6 Then
            Array.Copy(CardTemp, Card6, 25)
            Cardnames(6) = txtIDname.Text
        ElseIf cmbCardNum.SelectedItem = 7 Then
            Array.Copy(CardTemp, Card7, 25)
            Cardnames(7) = txtIDname.Text
        ElseIf cmbCardNum.SelectedItem = 8 Then
            Array.Copy(CardTemp, Card8, 25)
            Cardnames(8) = txtIDname.Text
        ElseIf cmbCardNum.SelectedItem = 9 Then
            Array.Copy(CardTemp, Card9, 25)
            Cardnames(9) = txtIDname.Text
        ElseIf cmbCardNum.SelectedItem = 10 Then
            Array.Copy(CardTemp, Card10, 25)
            Cardnames(10) = txtIDname.Text
        ElseIf cmbCardNum.SelectedItem = 11 Then
            Array.Copy(CardTemp, Card11, 25)
            Cardnames(11) = txtIDname.Text
        ElseIf cmbCardNum.SelectedItem = 12 Then
            Array.Copy(CardTemp, Card12, 25)
            Cardnames(12) = txtIDname.Text
        End If

    End Sub

    Private Sub btnCall_Click(sender As Object, e As EventArgs) Handles btnCall.Click
        Dim B As Integer
        Dim Called As Integer = Val(txtCalled.Text)

        Select Case Called
            Case 1 To 15
                For B = 0 To 4
                    If Called = Val(Card1(0, B)) Then
                        Card1Hits(0, B) = 1
                    End If
                    If Called = Val(Card2(0, B)) Then
                        Card2Hits(0, B) = 1
                    End If
                    If Called = Val(Card3(0, B)) Then
                        Card3Hits(0, B) = 1
                    End If
                    If Called = Val(Card4(0, B)) Then
                        Card4Hits(0, B) = 1
                    End If
                    If Called = Val(Card5(0, B)) Then
                        Card5Hits(0, B) = 1
                    End If
                    If Called = Val(Card6(0, B)) Then
                        Card6Hits(0, B) = 1
                    End If
                    If Called = Val(Card7(0, B)) Then
                        Card7Hits(0, B) = 1
                    End If
                    If Called = Val(Card8(0, B)) Then
                        Card8Hits(0, B) = 1
                    End If
                    If Called = Val(Card9(0, B)) Then
                        Card9Hits(0, B) = 1
                    End If
                    If Called = Val(Card10(0, B)) Then
                        Card10Hits(0, B) = 1
                    End If
                    If Called = Val(Card11(0, B)) Then
                        Card11Hits(0, B) = 1
                    End If
                    If Called = Val(Card12(0, B)) Then
                        Card12Hits(0, B) = 1
                    End If
                Next

            Case 16 To 30
                For B = 0 To 4
                    If Called = Val(Card1(1, B)) Then
                        Card1Hits(1, B) = 1
                    End If
                    If Called = Val(Card2(1, B)) Then
                        Card2Hits(1, B) = 1
                    End If
                    If Called = Val(Card3(1, B)) Then
                        Card3Hits(1, B) = 1
                    End If
                    If Called = Val(Card4(1, B)) Then
                        Card4Hits(1, B) = 1
                    End If
                    If Called = Val(Card5(1, B)) Then
                        Card5Hits(1, B) = 1
                    End If
                    If Called = Val(Card6(1, B)) Then
                        Card6Hits(1, B) = 1
                    End If
                    If Called = Val(Card7(1, B)) Then
                        Card7Hits(1, B) = 1
                    End If
                    If Called = Val(Card8(1, B)) Then
                        Card8Hits(1, B) = 1
                    End If
                    If Called = Val(Card9(1, B)) Then
                        Card9Hits(1, B) = 1
                    End If
                    If Called = Val(Card10(1, B)) Then
                        Card10Hits(1, B) = 1
                    End If
                    If Called = Val(Card11(1, B)) Then
                        Card11Hits(1, B) = 1
                    End If
                    If Called = Val(Card12(1, B)) Then
                        Card12Hits(1, B) = 1
                    End If
                Next

            Case 31 To 45
                For B = 0 To 4
                    If Called = Val(Card1(2, B)) Then
                        Card1Hits(2, B) = 1
                    End If
                    If Called = Val(Card2(2, B)) Then
                        Card2Hits(2, B) = 1
                    End If
                    If Called = Val(Card3(2, B)) Then
                        Card3Hits(2, B) = 1
                    End If
                    If Called = Val(Card4(2, B)) Then
                        Card4Hits(2, B) = 1
                    End If
                    If Called = Val(Card5(2, B)) Then
                        Card5Hits(2, B) = 1
                    End If
                    If Called = Val(Card6(2, B)) Then
                        Card6Hits(2, B) = 1
                    End If
                    If Called = Val(Card7(2, B)) Then
                        Card7Hits(2, B) = 1
                    End If
                    If Called = Val(Card8(2, B)) Then
                        Card8Hits(2, B) = 1
                    End If
                    If Called = Val(Card9(2, B)) Then
                        Card9Hits(2, B) = 1
                    End If
                    If Called = Val(Card10(2, B)) Then
                        Card10Hits(2, B) = 1
                    End If
                    If Called = Val(Card11(2, B)) Then
                        Card11Hits(2, B) = 1
                    End If
                    If Called = Val(Card12(2, B)) Then
                        Card12Hits(2, B) = 1
                    End If
                Next

            Case 46 To 60
                For B = 0 To 4
                    If Called = Val(Card1(3, B)) Then
                        Card1Hits(3, B) = 1
                    End If
                    If Called = Val(Card2(3, B)) Then
                        Card2Hits(3, B) = 1
                    End If
                    If Called = Val(Card3(3, B)) Then
                        Card3Hits(3, B) = 1
                    End If
                    If Called = Val(Card4(3, B)) Then
                        Card4Hits(3, B) = 1
                    End If
                    If Called = Val(Card5(3, B)) Then
                        Card5Hits(3, B) = 1
                    End If
                    If Called = Val(Card6(3, B)) Then
                        Card6Hits(3, B) = 1
                    End If
                    If Called = Val(Card7(3, B)) Then
                        Card7Hits(3, B) = 1
                    End If
                    If Called = Val(Card8(3, B)) Then
                        Card8Hits(3, B) = 1
                    End If
                    If Called = Val(Card9(3, B)) Then
                        Card9Hits(3, B) = 1
                    End If
                    If Called = Val(Card10(3, B)) Then
                        Card10Hits(3, B) = 1
                    End If
                    If Called = Val(Card11(3, B)) Then
                        Card11Hits(3, B) = 1
                    End If
                    If Called = Val(Card12(3, B)) Then
                        Card12Hits(3, B) = 1
                    End If
                Next
            Case 61 To 75
                For B = 0 To 4
                    If Called = Val(Card1(4, B)) Then
                        Card1Hits(4, B) = 1
                    End If
                    If Called = Val(Card2(4, B)) Then
                        Card2Hits(4, B) = 1
                    End If
                    If Called = Val(Card3(4, B)) Then
                        Card3Hits(4, B) = 1
                    End If
                    If Called = Val(Card4(4, B)) Then
                        Card4Hits(4, B) = 1
                    End If
                    If Called = Val(Card5(4, B)) Then
                        Card5Hits(4, B) = 1
                    End If
                    If Called = Val(Card6(4, B)) Then
                        Card6Hits(4, B) = 1
                    End If
                    If Called = Val(Card7(4, B)) Then
                        Card7Hits(4, B) = 1
                    End If
                    If Called = Val(Card8(4, B)) Then
                        Card8Hits(4, B) = 1
                    End If
                    If Called = Val(Card9(4, B)) Then
                        Card9Hits(4, B) = 1
                    End If
                    If Called = Val(Card10(4, B)) Then
                        Card10Hits(4, B) = 1
                    End If
                    If Called = Val(Card11(4, B)) Then
                        Card11Hits(4, B) = 1
                    End If
                    If Called = Val(Card12(4, B)) Then
                        Card12Hits(4, B) = 1
                    End If
                Next
        End Select

        If chkStandard.Checked Then
            CardWin(0) = StandardBingo(Card1Hits)
            CardWin(1) = StandardBingo(Card2Hits)
            CardWin(2) = StandardBingo(Card3Hits)
            CardWin(3) = StandardBingo(Card4Hits)
            CardWin(4) = StandardBingo(Card5Hits)
            CardWin(5) = StandardBingo(Card6Hits)
            CardWin(6) = StandardBingo(Card7Hits)
            CardWin(7) = StandardBingo(Card8Hits)
            CardWin(8) = StandardBingo(Card9Hits)
            CardWin(9) = StandardBingo(Card10Hits)
            CardWin(10) = StandardBingo(Card11Hits)
            CardWin(11) = StandardBingo(Card12Hits)

            For B = 0 To 11
                If CardWin(B) = "Bingo" Then
                    MessageBox.Show("Bingo on Card " & Cardnames(B))
                End If
            Next
        End If

        If chkAnyTwo.Checked Then
            CardWin(0) = AnyTwoBingo(Card1Hits)
            CardWin(1) = AnyTwoBingo(Card2Hits)
            CardWin(2) = AnyTwoBingo(Card3Hits)
            CardWin(3) = AnyTwoBingo(Card4Hits)
            CardWin(4) = AnyTwoBingo(Card5Hits)
            CardWin(5) = AnyTwoBingo(Card6Hits)
            CardWin(6) = AnyTwoBingo(Card7Hits)
            CardWin(7) = AnyTwoBingo(Card8Hits)
            CardWin(8) = AnyTwoBingo(Card9Hits)
            CardWin(9) = AnyTwoBingo(Card10Hits)
            CardWin(10) = AnyTwoBingo(Card11Hits)
            CardWin(11) = AnyTwoBingo(Card12Hits)

            For B = 0 To 11
                If CardWin(B) = "Bingo" Then
                    MessageBox.Show("Bingo on Card " & Cardnames(B))
                End If
            Next
        End If

        If chkBlackout.Checked Then
            CardWin(0) = BlackoutBingo(Card1Hits)
            CardWin(1) = BlackoutBingo(Card2Hits)
            CardWin(2) = BlackoutBingo(Card3Hits)
            CardWin(3) = BlackoutBingo(Card4Hits)
            CardWin(4) = BlackoutBingo(Card5Hits)
            CardWin(5) = BlackoutBingo(Card6Hits)
            CardWin(6) = BlackoutBingo(Card7Hits)
            CardWin(7) = BlackoutBingo(Card8Hits)
            CardWin(8) = BlackoutBingo(Card9Hits)
            CardWin(9) = BlackoutBingo(Card10Hits)
            CardWin(10) = BlackoutBingo(Card11Hits)
            CardWin(11) = BlackoutBingo(Card12Hits)

            For B = 0 To 11
                If CardWin(B) = "Bingo" Then
                    MessageBox.Show("Bingo on Card " & Cardnames(B))
                End If
            Next
        End If

        If chkPlus.Checked Then
            CardWin(0) = PlusBingo(Card1Hits)
            CardWin(1) = PlusBingo(Card2Hits)
            CardWin(2) = PlusBingo(Card3Hits)
            CardWin(3) = PlusBingo(Card4Hits)
            CardWin(4) = PlusBingo(Card5Hits)
            CardWin(5) = PlusBingo(Card6Hits)
            CardWin(6) = PlusBingo(Card7Hits)
            CardWin(7) = PlusBingo(Card8Hits)
            CardWin(8) = PlusBingo(Card9Hits)
            CardWin(9) = PlusBingo(Card10Hits)
            CardWin(10) = PlusBingo(Card11Hits)
            CardWin(11) = PlusBingo(Card12Hits)

            For B = 0 To 11
                If CardWin(B) = "Bingo" Then
                    MessageBox.Show("Bingo on Card " & Cardnames(B))
                End If
            Next
        End If

        If chkX.Checked Then
            CardWin(0) = XBingo(Card1Hits)
            CardWin(1) = XBingo(Card2Hits)
            CardWin(2) = XBingo(Card3Hits)
            CardWin(3) = XBingo(Card4Hits)
            CardWin(4) = XBingo(Card5Hits)
            CardWin(5) = XBingo(Card6Hits)
            CardWin(6) = XBingo(Card7Hits)
            CardWin(7) = XBingo(Card8Hits)
            CardWin(8) = XBingo(Card9Hits)
            CardWin(9) = XBingo(Card10Hits)
            CardWin(10) = XBingo(Card11Hits)
            CardWin(11) = XBingo(Card12Hits)

            For B = 0 To 11
                If CardWin(B) = "Bingo" Then
                    MessageBox.Show("Bingo on Card " & Cardnames(B))
                End If
            Next
        End If

        If chkRing.Checked Then
            CardWin(0) = RingBingo(Card1Hits)
            CardWin(1) = RingBingo(Card2Hits)
            CardWin(2) = RingBingo(Card3Hits)
            CardWin(3) = RingBingo(Card4Hits)
            CardWin(4) = RingBingo(Card5Hits)
            CardWin(5) = RingBingo(Card6Hits)
            CardWin(6) = RingBingo(Card7Hits)
            CardWin(7) = RingBingo(Card8Hits)
            CardWin(8) = RingBingo(Card9Hits)
            CardWin(9) = RingBingo(Card10Hits)
            CardWin(10) = RingBingo(Card11Hits)
            CardWin(11) = RingBingo(Card12Hits)

            For B = 0 To 11
                If CardWin(B) = "Bingo" Then
                    MessageBox.Show("Bingo on Card " & Cardnames(B))
                End If
            Next
        End If

        If chkTriple.Checked Then
            CardWin(0) = TripleBingo(Card1Hits)
            CardWin(1) = TripleBingo(Card2Hits)
            CardWin(2) = TripleBingo(Card3Hits)
            CardWin(3) = TripleBingo(Card4Hits)
            CardWin(4) = TripleBingo(Card5Hits)
            CardWin(5) = TripleBingo(Card6Hits)
            CardWin(6) = TripleBingo(Card7Hits)
            CardWin(7) = TripleBingo(Card8Hits)
            CardWin(8) = TripleBingo(Card9Hits)
            CardWin(9) = TripleBingo(Card10Hits)
            CardWin(10) = TripleBingo(Card11Hits)
            CardWin(11) = TripleBingo(Card12Hits)

            For B = 0 To 11
                If CardWin(B) = "Bingo" Then
                    MessageBox.Show("Bingo on Card " & Cardnames(B))
                End If
            Next
        End If

        If chkPyramid.Checked Then
            CardWin(0) = PyramidBingo(Card1Hits)
            CardWin(1) = PyramidBingo(Card2Hits)
            CardWin(2) = PyramidBingo(Card3Hits)
            CardWin(3) = PyramidBingo(Card4Hits)
            CardWin(4) = PyramidBingo(Card5Hits)
            CardWin(5) = PyramidBingo(Card6Hits)
            CardWin(6) = PyramidBingo(Card7Hits)
            CardWin(7) = PyramidBingo(Card8Hits)
            CardWin(8) = PyramidBingo(Card9Hits)
            CardWin(9) = PyramidBingo(Card10Hits)
            CardWin(10) = PyramidBingo(Card11Hits)
            CardWin(11) = PyramidBingo(Card12Hits)

            For B = 0 To 11
                If CardWin(B) = "Bingo" Then
                    MessageBox.Show("Bingo on Card " & Cardnames(B))
                End If
            Next
        End If


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

        If I = 5 Or N = 5 Or G = 5 Or O = 5 Or I2 = 5 Or N2 = 5 Or G2 = 5 Or O2 = 5 Then
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

        If I = 8 Then
            Return "Bingo"
        Else
            Return "Fail"
        End If
    End Function

    Public Function XBingo(CheckCard As Array) As String
        Dim B As Integer = 0
        B = CheckCard(0, 0) + CheckCard(0, 4) + CheckCard(4, 0) + CheckCard(4, 4) + CheckCard(1, 1) + CheckCard(1, 3) + CheckCard(3, 1) + CheckCard(3, 3)
        If B = 8 Then
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

        If N = 24 Then
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

        If I = 8 Or N = 8 Or G = 8 Or O = 8 Then
            Return "Bingo"
        Else
            Return "Fail"
        End If

    End Function

    Public Function RingBingo(CheckCard As Array) As String
        Dim B As Integer
        Dim I As Integer

        For B = 0 To 4
            I += CheckCard(B, 0)
            I += CheckCard(B, 4)
            If B <> (0 Or 4) Then
                I += CheckCard(0, B)
                I += CheckCard(4, B)
            End If
        Next
        If I = 14 Then
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
        Dim Diag1 As Integer = CheckCard(0, 0) + CheckCard(1, 1) + CheckCard(3, 3) + CheckCard(4, 4)
        Dim Diag2 As Integer = CheckCard(0, 4) + CheckCard(1, 3) + CheckCard(3, 1) + CheckCard(4, 0)
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
        Next

        If B = 5 Then
            If ((I Or G Or O Or B2 Or I2 Or G2 Or O2) = 5) Or ((N Or Diag1 Or Diag2) = 4) Then
                Return "Bingo"
            Else
                Return "Fail"
            End If
        ElseIf I = 5 Then
            If ((G Or O Or B2 Or I2 Or G2 Or O2) = 5) Or ((N Or N2 Or Diag1 Or Diag2) = 4) Then
                Return "Bingo"
            Else
                Return "Fail"
            End If
        ElseIf N = 4 Then
            If ((I Or G Or O Or B2 Or I2 Or G2 Or O2) = 5) Or ((N2 Or Diag1 Or Diag2) = 4) Then
                Return "Bingo"
            Else
                Return "Fail"
            End If
        ElseIf G = 5 Then
            If ((O Or B2 Or I2 Or G2 Or O2) = 5) Or ((N2 Or Diag1 Or Diag2) = 4) Then
                Return "Bingo"
            Else
                Return "Fail"
            End If
        ElseIf O = 5 Then
            If ((B2 Or I2 Or G2 Or O2) = 5) Or ((N2 Or Diag1 Or Diag2) = 4) Then
                Return "Bingo"
            Else
                Return "Fail"
            End If
        ElseIf Diag1 = 4 Then
            If ((B2 Or I2 Or G2 Or O2) = 5) Or ((N2 Or Diag2) = 4) Then
                Return "Bingo"
            Else
                Return "Fail"
            End If
        ElseIf Diag2 = 4 Then
            If ((B2 Or I2 Or G2 Or O2) = 5) Or (N2 = 4) Then
                Return "Bingo"
            Else
                Return "Fail"
            End If
        ElseIf B2 = 5 Then
            If ((I2 Or G2 Or O2) = 5) Or (N2 = 4) Then
                Return "Bingo"
            Else
                Return "Fail"
            End If
        ElseIf I2 = 5 Then
            If ((G2 Or O2) = 5) Or (N2 = 4) Then
                Return "Bingo"
            Else
                Return "Fail"
            End If
        ElseIf N2 = 4 Then
            If (G2 Or O2) = 5 Then
                Return "Bingo"
            Else
                Return "Fail"
            End If
        ElseIf (G2 And O2) = 5 Then
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
        Dim Diag1 As Integer = CheckCard(0, 0) + CheckCard(1, 1) + CheckCard(3, 3) + CheckCard(4, 4)
        Dim Diag2 As Integer = CheckCard(0, 4) + CheckCard(1, 3) + CheckCard(3, 1) + CheckCard(4, 0)
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
        Next

        If (Diag1 Or Diag2) = 4 Then
            If ((B Or I Or G Or O) = 5) Or (N = 4) Then
                If ((B2 Or I2 Or G2 Or O2) = 5) Or (N2 = 4) Then
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
            Array.Copy(ClearHits, Card1, 25)
            Array.Copy(ClearHits, Card2, 25)
            Array.Copy(ClearHits, Card3, 25)
            Array.Copy(ClearHits, Card4, 25)
            Array.Copy(ClearHits, Card5, 25)
            Array.Copy(ClearHits, Card6, 25)
            Array.Copy(ClearHits, Card7, 25)
            Array.Copy(ClearHits, Card8, 25)
            Array.Copy(ClearHits, Card9, 25)
            Array.Copy(ClearHits, Card10, 25)
            Array.Copy(ClearHits, Card11, 25)
            Array.Copy(ClearHits, Card12, 25)
        ElseIf Warning = DialogResult.Cancel Then

        End If

    End Sub
End Class
