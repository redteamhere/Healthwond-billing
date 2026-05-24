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

        Public Shared Sub StyleDangerButton(button As Button)
            button.FlatStyle = FlatStyle.Flat
            button.FlatAppearance.BorderSize = 0
            button.BackColor = ThemePalette.DangerRed
            button.ForeColor = Color.White
            button.Cursor = Cursors.Hand
            button.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
            button.Height = 44
        End Sub

        Public Shared Sub StyleDataGrid(grid As DataGridView)
            grid.BackgroundColor = Color.White
            grid.BorderStyle = BorderStyle.None
            grid.AutoGenerateColumns = False
            grid.AllowUserToAddRows = False
            grid.AllowUserToDeleteRows = False
            grid.AllowUserToResizeRows = False
            grid.MultiSelect = False
            grid.ReadOnly = True
            grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            grid.RowHeadersVisible = False
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            grid.EnableHeadersVisualStyles = False
            grid.ColumnHeadersHeight = 40
            grid.ColumnHeadersDefaultCellStyle.BackColor = ThemePalette.BrandBlue
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
            grid.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
            grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(216, 229, 255)
            grid.DefaultCellStyle.SelectionForeColor = ThemePalette.TextPrimary
            grid.DefaultCellStyle.BackColor = Color.White
            grid.DefaultCellStyle.ForeColor = ThemePalette.TextPrimary
            grid.DefaultCellStyle.Font = New Font("Segoe UI", 9.75F, FontStyle.Regular)
            grid.RowsDefaultCellStyle.BackColor = Color.White
            grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(247, 249, 252)
            grid.GridColor = Color.FromArgb(230, 233, 240)
        End Sub

        Public Shared Sub StyleInput(textBox As TextBox)
            textBox.Font = New Font("Segoe UI", 10.5F, FontStyle.Regular)
            textBox.Margin = New Padding(0)
        End Sub

        Public Shared Sub StyleCard(control As Control)
            control.BackColor = ThemePalette.CardBackground
            If control.Padding = Padding.Empty Then
                control.Padding = New Padding(24)
            End If
        End Sub

        Public Shared Function CreateScrollableHost(content As Control) As Panel
            Dim host As New Panel With {
                .Dock = DockStyle.Fill,
                .AutoScroll = True,
                .BackColor = ThemePalette.CardBackground
            }

            content.Dock = DockStyle.Top
            content.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            content.AutoSize = True

            host.Controls.Add(content)

            AddHandler host.SizeChanged,
                Sub()
                    Dim targetWidth As Integer = host.DisplayRectangle.Width
                    If targetWidth <= 0 Then
                        targetWidth = host.ClientSize.Width
                    End If

                    content.Width = Math.Max(targetWidth, 0)
                End Sub

            Return host
        End Function

    End Class

End Namespace
