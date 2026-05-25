Imports HealthwondBilling.Utilities

Namespace Forms

    Public Class FrmTextPrompt
        Inherits Form

        Private ReadOnly txtValue As New TextBox()
        Private ReadOnly btnOk As New Button()
        Private ReadOnly btnCancel As New Button()

        Public Property PromptValue As String
            Get
                Return txtValue.Text.Trim()
            End Get
            Set(value As String)
                txtValue.Text = value
            End Set
        End Property

        Public Sub New(title As String, promptText As String, Optional initialValue As String = "", Optional isPassword As Boolean = False)
            Text = title
            StartPosition = FormStartPosition.CenterParent
            FormBorderStyle = FormBorderStyle.FixedDialog
            MaximizeBox = False
            MinimizeBox = False
            ShowInTaskbar = False
            ClientSize = New Size(520, 188)
            BackColor = ThemePalette.AppBackground
            Font = New Font("Segoe UI", 10.0F, FontStyle.Regular)

            Dim root As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 4,
                .Padding = New Padding(18)
            }
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 48))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 64))
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 42))

            Dim lblPrompt As New Label With {
                .Dock = DockStyle.Fill,
                .Text = promptText,
                .Font = New Font("Segoe UI", 11.0F, FontStyle.Regular),
                .ForeColor = ThemePalette.TextPrimary
            }

            txtValue.Dock = DockStyle.Top
            txtValue.Height = 36
            txtValue.BorderStyle = BorderStyle.FixedSingle
            txtValue.Font = New Font("Segoe UI", 11.0F, FontStyle.Regular)
            txtValue.Text = initialValue
            txtValue.UseSystemPasswordChar = isPassword

            Dim note As New Label With {
                .Dock = DockStyle.Fill,
                .ForeColor = ThemePalette.TextMuted,
                .Text = "Press Enter to confirm or Esc to cancel.",
                .TextAlign = ContentAlignment.TopLeft
            }

            Dim buttonPanel As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.RightToLeft,
                .WrapContents = False
            }

            btnOk.Text = "OK"
            btnOk.Width = 96
            btnOk.Height = 34
            btnOk.FlatStyle = FlatStyle.Flat
            btnOk.FlatAppearance.BorderSize = 0
            btnOk.BackColor = ThemePalette.BrandBlue
            btnOk.ForeColor = Color.White
            AddHandler btnOk.Click, Sub() DialogResult = DialogResult.OK

            btnCancel.Text = "Cancel"
            btnCancel.Width = 96
            btnCancel.Height = 34
            btnCancel.FlatStyle = FlatStyle.Flat
            btnCancel.FlatAppearance.BorderSize = 1
            btnCancel.FlatAppearance.BorderColor = ThemePalette.TextMuted
            btnCancel.BackColor = Color.White
            btnCancel.ForeColor = ThemePalette.TextPrimary
            AddHandler btnCancel.Click, Sub() DialogResult = DialogResult.Cancel

            buttonPanel.Controls.Add(btnOk)
            buttonPanel.Controls.Add(btnCancel)

            root.Controls.Add(lblPrompt, 0, 0)
            root.Controls.Add(txtValue, 0, 1)
            root.Controls.Add(note, 0, 2)
            root.Controls.Add(buttonPanel, 0, 3)

            Controls.Add(root)

            AcceptButton = btnOk
            CancelButton = btnCancel
        End Sub

        Protected Overrides Sub OnShown(e As EventArgs)
            MyBase.OnShown(e)
            txtValue.Focus()
            txtValue.SelectionStart = txtValue.TextLength
        End Sub

    End Class

End Namespace
