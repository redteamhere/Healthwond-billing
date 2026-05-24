Namespace Utilities

    Public NotInheritable Class UiStyler

        Private Sub New()
        End Sub

        Public Shared Sub StylePrimaryButton(button As Button)
            button.FlatStyle = FlatStyle.Flat
            button.FlatAppearance.BorderSize = 0
            button.BackColor = ThemePalette.BrandBlue
            button.ForeColor = Color.White
            button.Cursor = Cursors.Hand
            button.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
            button.Height = 44
        End Sub

        Public Shared Sub StyleSecondaryButton(button As Button)
            button.FlatStyle = FlatStyle.Flat
            button.FlatAppearance.BorderColor = Color.FromArgb(208, 214, 222)
            button.FlatAppearance.BorderSize = 1
            button.BackColor = Color.White
            button.ForeColor = ThemePalette.TextPrimary
            button.Cursor = Cursors.Hand
            button.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
            button.Height = 44
        End Sub

        Public Shared Sub StyleInput(textBox As TextBox)
            textBox.Font = New Font("Segoe UI", 10.5F, FontStyle.Regular)
            textBox.Margin = New Padding(0)
        End Sub

        Public Shared Sub StyleCard(control As Control)
            control.BackColor = ThemePalette.CardBackground
            control.Padding = New Padding(24)
        End Sub

    End Class

End Namespace
