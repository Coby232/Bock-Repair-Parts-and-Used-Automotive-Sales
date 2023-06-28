Public Class Form1

    'Error provider
    Dim ErrorProvider As New ErrorProvider

    'Exterior Extras
    Const NoneDecimal As Double = 0.0D
    Const PaintTouchUpDecimal As Double = 250D
    Const UndercoatDecimal As Double = 300D
    Const BothExteriorDecimal As Double = 550D

    'Accessory Extras
    Const NewTiresDecimal As Double = 450D
    Const NewHdDecimal As Double = 190.95D
    Const GPS_GPS_Decimal As Double = 700D
    Const NewFloorMatsDecimal As Double = 55D

    'Local accumulator
    Public TotalDuesAllSales, AverageTotalDues As Double
    Public TotalVehiclesSold As Integer


    'resetfunction Button
    Friend Function resetfunction$()
        'clear text boxes
        PriceTextBox.Clear()
        DiscountTextBox.Clear()
        ExtrasTextBox.Clear()
        SubtotalTextBox.Clear()
        SalesTaxTextBox.Clear()
        TradeInTextBox.Text = "0.00" 'exception
        TotalDueTextBox.Clear()
        LotNumTextBox.Clear()
        VehicleModelTextBox.Clear()

        'uncheck checkboxes
        NoneRadioButton.Checked = True
        NewTireCheckBox.Checked = False
        NewHDCheckBox.Checked = False
        NewFloorCheckBox.Checked = False
        GPSCheckBox.Checked = False

        'set focus to lotTextbox
        LotNumTextBox.Focus()
    End Function


    'Function for Exterior extras
    Friend Function checkboxfunction$(checked As Boolean, CheckPrice As Double)
        Dim count As Integer = 0
        If checked = True And count = 0 Then
            Try
                Dim ExtraPrice As Double
                ExtraPrice = Decimal.Parse(ExtrasTextBox.Text) + CheckPrice
                ExtrasTextBox.Text = ExtraPrice.ToString("N2")
                count += 1
            Catch ex As Exception
                MessageBox.Show("Error")
            End Try
        End If
        If checked = False Then
            Try
                Dim cost As Double
                cost = Decimal.Parse(ExtrasTextBox.Text) - CheckPrice
                ExtrasTextBox.Text = cost.ToString("N2")
                count = 0
            Catch ex As Exception
                MessageBox.Show("Error")
            End Try
        End If
    End Function

    'compute button
    Private Sub ComputeButton_Click(sender As Object, e As EventArgs) Handles ComputeButton.Click
        Try
            If IsNumeric(LotNumTextBox.Text) = True And LotNumTextBox.Text <> "" And YearTextBox.Text <> "" _
                And IsNumeric(YearTextBox.Text) And VehicleModelTextBox.Text <> "" Then

                If Decimal.Parse(PriceTextBox.Text) > 0 Then
                    If Decimal.Parse(TradeInTextBox.Text) >= 0 Then 'RULE #1 to #4



                        Const SalesTaxDecimal As Double = 0.05D
                        Const WholesaleDiscountDecimal As Double = 0.2D
                        Dim TotalDue, TradeIn, Extras, Discounts, Subtotal, Tax, cost As Double

                        'wholesale discount
                        If WholesaleCheckBox.Checked = True Then
                            Discounts = Decimal.Parse(PriceTextBox.Text) * WholesaleDiscountDecimal
                            DiscountTextBox.Text = Discounts.ToString("N2")
                            SalesTaxTextBox.Text = "0.00"
                        Else
                            DiscountTextBox.Text = "0.00"
                            Tax = Decimal.Parse(PriceTextBox.Text) * SalesTaxDecimal
                            SalesTaxTextBox.Text = Tax.ToString("N2")
                        End If

                        cost = Decimal.Parse(PriceTextBox.Text, Globalization.NumberStyles.Currency)
                        Extras = Decimal.Parse(ExtrasTextBox.Text)
                        Subtotal = (cost + Extras) - (Discounts + Tax)
                        SubtotalTextBox.Text = Subtotal.ToString("N2")
                        TradeIn = Decimal.Parse(TradeInTextBox.Text)
                        TotalDue = Subtotal - Tax - TradeIn
                        TotalDueTextBox.Text = TotalDue.ToString("C2")

                        'Accumulation
                        TotalDuesAllSales += TotalDue
                        TotalVehiclesSold += 1
                        AverageTotalDues = TotalDuesAllSales / TotalVehiclesSold

                        'custom error handling
                    ElseIf IsNumeric(TradeInTextBox.Text.Trim) = True Then
                        ErrorProvider.SetError(Me.TradeInTextBox, "the trade-in TextBox must contain a numeric value that is greater than or equal to zero")
                        TradeInTextBox.Focus()
                    End If
                ElseIf IsNumeric(PriceTextBox.Text.Trim) = True Then
                    ErrorProvider.SetError(Me.PriceTextBox, "the cost TextBox must contain a numeric value that is greater than zero")
                    PriceTextBox.Focus()
                End If
            ElseIf String.IsNullOrEmpty(LotNumTextBox.Text.Trim) Then 'RULE #1
                ErrorProvider.SetError(Me.LotNumTextBox, "the lot number TextBox cannot be blank.")
                LotNumTextBox.Focus()
            ElseIf String.IsNullOrEmpty(YearTextBox.Text.Trim) Then 'RULE #2
                ErrorProvider.SetError(Me.YearTextBox, "the year TextBox cannot be blank.")
                YearTextBox.Focus()
            ElseIf String.IsNullOrEmpty(VehicleModelTextBox.Text.Trim) Then 'RULE #2
                ErrorProvider.SetError(Me.VehicleModelTextBox, "the vehicle make/model TextBox cannot be blank")
                VehicleModelTextBox.Focus()
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub UndercoatRadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles UndercoatRadioButton.CheckedChanged
        checkboxfunction(UndercoatRadioButton.checked, UndercoatDecimal)
    End Sub

    Private Sub PaintRadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles PaintRadioButton.CheckedChanged
        checkboxfunction(PaintRadioButton.checked, PaintTouchUpDecimal)
    End Sub

    Private Sub BothRadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles BothRadioButton.CheckedChanged
        checkboxfunction(BothRadioButton.checked, BothExteriorDecimal)
    End Sub

    'Accessory Checkbox Events
    Private Sub NewTireCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles NewTireCheckBox.CheckedChanged
        checkboxfunction(NewTireCheckBox.checked, NewTiresDecimal)
    End Sub

    Private Sub NewHDCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles NewHDCheckBox.CheckedChanged
        checkboxfunction(NewHDCheckBox.checked, NewHdDecimal)
    End Sub

    Private Sub GPSCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles GPSCheckBox.CheckedChanged
        checkboxfunction(GPSCheckBox.checked, GPS_GPS_Decimal)
    End Sub

    Private Sub NewFloorCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles NewFloorCheckBox.CheckedChanged
        checkboxfunction(NewFloorCheckBox.checked, NewFloorMatsDecimal)
    End Sub

    'MAPPPING KEYS
    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown,
        YearTextBox.KeyDown, LotNumTextBox.KeyDown, WholesaleCheckBox.KeyDown, VehicleModelTextBox.KeyDown, UndercoatRadioButton.KeyDown,
        TradeInTextBox.KeyDown, TotalDueTextBox.KeyDown, SubtotalTextBox.KeyDown, SalesTaxTextBox.KeyDown, PriceTextBox.KeyDown,
        PaintRadioButton.KeyDown, NoneRadioButton.KeyDown, NewTireCheckBox.KeyDown, NewHDCheckBox.KeyDown, NewFloorCheckBox.KeyDown,
        GPSCheckBox.KeyDown, ExtrasTextBox.KeyDown, DiscountTextBox.KeyDown, BothRadioButton.KeyDown

        'If enter key is pressed
        If e.KeyCode = Keys.Enter Then
            ComputeButton.PerformClick()
        End If

        'ESC for resetfunction
        If e.KeyCode = Keys.Escape Then
            resetfunctionButton.PerformClick()
        End If

        'Capslock for Totals
        If e.KeyCode = Keys.ControlKey Then
            TotalButton.PerformClick()
        End If

        'Numlock must be on plus 9 for Exit
        If e.KeyCode = Keys.NumPad9 Then
            ExitButton.PerformClick()
        End If

        'Hot keys for input textboxes

        If e.KeyCode = Keys.F2 Then
            LotNumTextBox.Focus()
        End If

        If e.KeyCode = Keys.F1 Then
            VehicleModelTextBox.Focus()
        End If

        If e.KeyCode = Keys.F3 Then
            YearTextBox.Focus()
        End If

        If e.KeyCode = Keys.F4 Then
            PriceTextBox.Focus()
        End If

        If e.KeyCode = Keys.F7 Then
            TradeInTextBox.Focus()
        End If

        'Hot keys for check boxes
        If e.KeyCode = Keys.F6 Then
            NoneRadioButton.checked = True
        End If

        If e.KeyCode = Keys.F5 Then
            PaintRadioButton.checked = True
        End If

        If e.KeyCode = Keys.F8 Then
            UndercoatRadioButton.checked = True
        End If

        If e.KeyCode = Keys.F9 Then
            BothRadioButton.checked = True
        End If

        If e.KeyCode = Keys.F10 Then
            NewTireCheckBox.checked = True
        End If

        If e.KeyCode = Keys.F11 Then
            NewHDCheckBox.checked = True
        End If

        If e.KeyCode = Keys.F12 Then
            GPSCheckBox.checked = True
        End If

        If e.KeyCode = Keys.CapsLock Then
            NewFloorCheckBox.checked = True
        End If



    End Sub

    Private Sub NoneRadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles NoneRadioButton.CheckedChanged
        checkboxfunction(NoneRadioButton.checked, NoneDecimal)
    End Sub

    'Total button
    Private Sub TotalButton_Click(sender As Object, e As EventArgs) Handles TotalButton.Click
        Try
            If TotalVehiclesSold > 0 Then
                MsgBox("Total due all sales: $" & TotalDuesAllSales & vbNewLine & "Total vehicles sold: " & TotalVehiclesSold & vbNewLine &
                                "Average total due: $" & AverageTotalDues, MsgBoxStyle.Information, "Totals and Averages")
            End If
        Catch ex As Exception
            MsgBox("Error")
        End Try
    End Sub

    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        Try
            Dim Convo As MsgBoxResult
            Convo = MsgBox("Exit application (Y/N)?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question, "Exit ?")
            If Convo = MsgBoxResult.Yes Then
                Application.ExitThread()
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub resetfunctionButton_Click(sender As Object, e As EventArgs) Handles resetfunctionButton.Click
        resetfunction()
    End Sub

    'Private Sub GroupBox_PreviewKeyDown(sender As Object, e As PreviewKeyDownEventArgs) Handles ExteriorGroupBox.KeyDown, AutomotiveInformationGroupBox.KeyDown
    '    If e.KeyCode = Keys.Enter Then
    '        ComputeButton.PerformClick()
    '    End If
    'End Sub

End Class
