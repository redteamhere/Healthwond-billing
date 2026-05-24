Imports HealthwondBilling.Utilities

Namespace Controls

    Public Class StatCardControl
        Inherits UserControl

        Private ReadOnly _accentPanel As Panel
        Private ReadOnly _titleLabel As Label
        Private ReadOnly _valueLabel As Label
        Private ReadOnly _subtitleLabel As Label

        Public Sub New()
            DoubleBuffered = True
            BackColor = ThemePalette.CardBackground
            Padding = New Padding(18)
            Size = New Size(215, 132)
            Margin = New Padding(0, 0, 18, 18)

            _accentPanel = New Panel With {
                .Dock = DockStyle.Top,
                .Height = 6,
                .BackColor = ThemePalette.AccentGreen
            }

            _titleLabel = New Label With {
                .Dock = DockStyle.Top,
                .Height = 28,
                .Font = New Font("Segoe UI Semibold", 9.75F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextMuted,
                .TextAlign = ContentAlignment.MiddleLeft
            }

            _valueLabel = New Label With {
                .Dock = DockStyle.Top,
                .Height = 42,
                .Font = New Font("Segoe UI Semibold", 21.0F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary,
                .TextAlign = ContentAlignment.MiddleLeft
            }

            _subtitleLabel = New Label With {
                .Dock = DockStyle.Fill,
                .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular),
                .ForeColor = ThemePalette.TextMuted,
                .TextAlign = ContentAlignment.TopLeft
            }

            Controls.Add(_subtitleLabel)
            Controls.Add(_valueLabel)
            Controls.Add(_titleLabel)
            Controls.Add(_accentPanel)

            CardTitle = "Title"
            ValueText = "0"
            SubtitleText = String.Empty
        End Sub

        Public Property CardTitle As String
            Get
                Return _titleLabel.Text
            End Get
            Set(value As String)
                _titleLabel.Text = value
            End Set
        End Property

        Public Property ValueText As String
            Get
                Return _valueLabel.Text
            End Get
            Set(value As String)
                _valueLabel.Text = value
            End Set
        End Property

        Public Property SubtitleText As String
            Get
                Return _subtitleLabel.Text
            End Get
            Set(value As String)
                _subtitleLabel.Text = value
            End Set
        End Property

        Public Property AccentColor As Color
            Get
                Return _accentPanel.BackColor
            End Get
            Set(value As Color)
                _accentPanel.BackColor = value
            End Set
        End Property

    End Class

End Namespace
