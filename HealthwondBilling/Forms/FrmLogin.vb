Imports HealthwondBilling.Models
Imports HealthwondBilling.Services
Imports HealthwondBilling.Utilities
Imports System.Linq

Namespace Forms

    Public Class FrmLogin
        Inherits Form

        Private ReadOnly _authService As AuthService
        Private ReadOnly _dashboardService As DashboardService
        Private ReadOnly _productService As ProductService
        Private ReadOnly _customerService As CustomerService
        Private ReadOnly _billingService As BillingService

        Private ReadOnly txtUsername As TextBox
        Private ReadOnly txtPassword As TextBox
        Private ReadOnly chkShowPassword As CheckBox
        Private ReadOnly btnLogin As Button
        Private ReadOnly btnExit As Button
        Private ReadOnly lblStatus As Label

        Public Sub New(authService As AuthService, dashboardService As DashboardService, productService As ProductService, customerService As CustomerService, billingService As BillingService)
            _authService = authService
            _dashboardService = dashboardService
            _productService = productService
            _customerService = customerService
            _billingService = billingService

            Text = "Healthwond Billing System - Login"
            StartPosition = FormStartPosition.CenterScreen
            MinimumSize = New Size(1080, 680)
            Size = New Size(1180, 720)
            BackColor = ThemePalette.AppBackground
            Font = New Font("Segoe UI", 10.0F, FontStyle.Regular)

            Dim shellLayout As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2,
                .RowCount = 1,
                .BackColor = ThemePalette.AppBackground
            }
            shellLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 55.0F))
            shellLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 45.0F))

            Dim brandPanel As Panel = BuildBrandPanel()
            Dim formPanel As Panel = BuildLoginPanel()

            shellLayout.Controls.Add(brandPanel, 0, 0)
            shellLayout.Controls.Add(formPanel, 1, 0)
            Controls.Add(shellLayout)

            txtUsername = CType(formPanel.Controls.Find("txtUsername", True).Single(), TextBox)
            txtPassword = CType(formPanel.Controls.Find("txtPassword", True).Single(), TextBox)
            chkShowPassword = CType(formPanel.Controls.Find("chkShowPassword", True).Single(), CheckBox)
            btnLogin = CType(formPanel.Controls.Find("btnLogin", True).Single(), Button)
            btnExit = CType(formPanel.Controls.Find("btnExit", True).Single(), Button)
            lblStatus = CType(formPanel.Controls.Find("lblStatus", True).Single(), Label)

            AddHandler btnLogin.Click, AddressOf btnLogin_Click
            AddHandler btnExit.Click, AddressOf btnExit_Click
            AddHandler chkShowPassword.CheckedChanged, AddressOf chkShowPassword_CheckedChanged

            txtUsername.Text = "admin"

            AcceptButton = btnLogin
            CancelButton = btnExit
        End Sub

        Private Function BuildBrandPanel() As Panel
            Dim panel As New Panel With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(48),
                .BackColor = ThemePalette.BrandBlue
            }

            Dim titleLabel As New Label With {
                .Dock = DockStyle.Top,
                .Height = 94,
                .Font = New Font("Segoe UI Semibold", 28.0F, FontStyle.Bold),
                .ForeColor = Color.White,
                .Text = "Healthwond Billing System",
                .TextAlign = ContentAlignment.BottomLeft
            }

            Dim subtitleLabel As New Label With {
                .Dock = DockStyle.Top,
                .Height = 84,
                .Font = New Font("Segoe UI", 12.0F, FontStyle.Regular),
                .ForeColor = Color.FromArgb(225, 232, 241),
                .Text = "Pharmaceutical billing, stock control, GST invoicing, and audit-ready workflows built on a layered WinForms architecture.",
                .TextAlign = ContentAlignment.TopLeft
            }

            Dim pointsPanel As New FlowLayoutPanel With {
                .Dock = DockStyle.Top,
                .Height = 230,
                .FlowDirection = FlowDirection.TopDown,
                .WrapContents = False,
                .BackColor = Color.Transparent,
                .Padding = New Padding(0, 18, 0, 0)
            }

            pointsPanel.Controls.Add(CreateFeatureLabel("Role-based authentication for admin and staff users"))
            pointsPanel.Controls.Add(CreateFeatureLabel("SQLite schema bootstrap with secure seeded sample data"))
            pointsPanel.Controls.Add(CreateFeatureLabel("Dashboard metrics for sales, stock, expiry, low stock, and dues"))
            pointsPanel.Controls.Add(CreateFeatureLabel("Future-ready structure for billing, purchase, reporting, and template workflows"))

            Dim credentialsLabel As New Label With {
                .Dock = DockStyle.Bottom,
                .Height = 66,
                .Font = New Font("Segoe UI", 10.0F, FontStyle.Regular),
                .ForeColor = Color.FromArgb(225, 232, 241),
                .Text = "Default demo users: admin / Admin@123 and staff / Staff@123",
                .TextAlign = ContentAlignment.BottomLeft
            }

            panel.Controls.Add(credentialsLabel)
            panel.Controls.Add(pointsPanel)
            panel.Controls.Add(subtitleLabel)
            panel.Controls.Add(titleLabel)

            Return panel
        End Function

        Private Function CreateFeatureLabel(text As String) As Label
            Return New Label With {
                .Width = 420,
                .Height = 38,
                .Font = New Font("Segoe UI", 11.0F, FontStyle.Regular),
                .ForeColor = Color.White,
                .Text = $"- {text}",
                .TextAlign = ContentAlignment.MiddleLeft,
                .Margin = New Padding(0, 0, 0, 10)
            }
        End Function

        Private Function BuildLoginPanel() As Panel
            Dim hostPanel As New Panel With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(50),
                .BackColor = ThemePalette.AppBackground
            }

            Dim card As New Panel With {
                .Dock = DockStyle.None,
                .BackColor = ThemePalette.CardBackground,
                .Size = New Size(420, 460),
                .Anchor = AnchorStyles.None
            }
            UiStyler.StyleCard(card)

            Dim shell As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .RowCount = 9,
                .ColumnCount = 1,
                .Padding = New Padding(36)
            }
            shell.RowStyles.Add(New RowStyle(SizeType.Absolute, 56))
            shell.RowStyles.Add(New RowStyle(SizeType.Absolute, 52))
            shell.RowStyles.Add(New RowStyle(SizeType.Absolute, 72))
            shell.RowStyles.Add(New RowStyle(SizeType.Absolute, 72))
            shell.RowStyles.Add(New RowStyle(SizeType.Absolute, 32))
            shell.RowStyles.Add(New RowStyle(SizeType.Absolute, 64))
            shell.RowStyles.Add(New RowStyle(SizeType.Absolute, 54))
            shell.RowStyles.Add(New RowStyle(SizeType.Absolute, 40))
            shell.RowStyles.Add(New RowStyle(SizeType.Percent, 100))

            Dim heading As New Label With {
                .Dock = DockStyle.Fill,
                .Font = New Font("Segoe UI Semibold", 20.0F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary,
                .Text = "Sign in",
                .TextAlign = ContentAlignment.BottomLeft
            }

            Dim intro As New Label With {
                .Dock = DockStyle.Fill,
                .Font = New Font("Segoe UI", 10.0F, FontStyle.Regular),
                .ForeColor = ThemePalette.TextMuted,
                .Text = "Use your Healthwond account to access pharmacy billing operations.",
                .TextAlign = ContentAlignment.MiddleLeft
            }

            Dim usernamePanel As Panel = CreateLabeledInput("Username", "txtUsername", False)
            Dim passwordPanel As Panel = CreateLabeledInput("Password", "txtPassword", True)

            Dim optionsPanel As New Panel With {.Dock = DockStyle.Fill}
            Dim showPassword As New CheckBox With {
                .Name = "chkShowPassword",
                .Text = "Show password",
                .AutoSize = True,
                .ForeColor = ThemePalette.TextMuted,
                .Location = New Point(0, 6)
            }
            optionsPanel.Controls.Add(showPassword)

            Dim actionPanel As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2,
                .RowCount = 1
            }
            actionPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 65.0F))
            actionPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 35.0F))

            Dim loginButton As New Button With {.Name = "btnLogin", .Text = "Sign In", .Dock = DockStyle.Fill}
            Dim exitButton As New Button With {.Name = "btnExit", .Text = "Exit", .Dock = DockStyle.Fill}
            UiStyler.StylePrimaryButton(loginButton)
            UiStyler.StyleSecondaryButton(exitButton)

            actionPanel.Controls.Add(loginButton, 0, 0)
            actionPanel.Controls.Add(exitButton, 1, 0)

            Dim shortcutLabel As New Label With {
                .Dock = DockStyle.Fill,
                .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular),
                .ForeColor = ThemePalette.TextMuted,
                .Text = "Shortcut: Press Enter to sign in.",
                .TextAlign = ContentAlignment.MiddleLeft
            }

            Dim statusLabel As New Label With {
                .Name = "lblStatus",
                .Dock = DockStyle.Fill,
                .Font = New Font("Segoe UI Semibold", 9.5F, FontStyle.Bold),
                .ForeColor = ThemePalette.DangerRed,
                .TextAlign = ContentAlignment.MiddleLeft
            }

            shell.Controls.Add(heading, 0, 0)
            shell.Controls.Add(intro, 0, 1)
            shell.Controls.Add(usernamePanel, 0, 2)
            shell.Controls.Add(passwordPanel, 0, 3)
            shell.Controls.Add(optionsPanel, 0, 4)
            shell.Controls.Add(actionPanel, 0, 5)
            shell.Controls.Add(shortcutLabel, 0, 6)
            shell.Controls.Add(statusLabel, 0, 7)

            card.Controls.Add(shell)
            hostPanel.Controls.Add(card)

            AddHandler hostPanel.Resize,
                Sub()
                    card.Location = New Point(Math.Max((hostPanel.ClientSize.Width - card.Width) \ 2, 0), Math.Max((hostPanel.ClientSize.Height - card.Height) \ 2, 0))
                End Sub

            Return hostPanel
        End Function

        Private Function CreateLabeledInput(labelText As String, textBoxName As String, isPassword As Boolean) As Panel
            Dim wrapper As New Panel With {.Dock = DockStyle.Fill}

            Dim caption As New Label With {
                .Dock = DockStyle.Top,
                .Height = 24,
                .Font = New Font("Segoe UI Semibold", 9.5F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary,
                .Text = labelText
            }

            Dim textbox As New TextBox With {
                .Name = textBoxName,
                .Dock = DockStyle.Bottom,
                .Height = 42,
                .BorderStyle = BorderStyle.FixedSingle,
                .UseSystemPasswordChar = isPassword
            }
            UiStyler.StyleInput(textbox)

            wrapper.Controls.Add(textbox)
            wrapper.Controls.Add(caption)

            Return wrapper
        End Function

        Private Async Sub btnLogin_Click(sender As Object, e As EventArgs)
            Await LoginAsync()
        End Sub

        Private Sub btnExit_Click(sender As Object, e As EventArgs)
            Close()
        End Sub

        Private Sub chkShowPassword_CheckedChanged(sender As Object, e As EventArgs)
            txtPassword.UseSystemPasswordChar = Not chkShowPassword.Checked
        End Sub

        Private Async Function LoginAsync() As Task
            lblStatus.Text = String.Empty

            ToggleUi(False)

            Try
                Dim result As LoginResult = Await _authService.AuthenticateAsync(txtUsername.Text, txtPassword.Text)
                If Not result.IsSuccess Then
                    lblStatus.Text = result.Message
                    Return
                End If

                SessionManager.StartSession(result.User)

                Using dashboard As New FrmDashboard(_dashboardService, _productService, _customerService, _billingService)
                    Hide()
                    dashboard.ShowDialog(Me)

                    If dashboard.RequestedLogout Then
                        SessionManager.Clear()
                        txtPassword.Clear()
                        lblStatus.ForeColor = ThemePalette.AccentGreen
                        lblStatus.Text = "You have been signed out."
                        Show()
                        Activate()
                        txtPassword.Focus()
                        lblStatus.ForeColor = ThemePalette.DangerRed
                    Else
                        Close()
                    End If
                End Using
            Catch ex As Exception
                AppLogger.Error("Login workflow failed.", ex)
                lblStatus.Text = "Unable to complete the sign-in request."
            Finally
                ToggleUi(True)
            End Try
        End Function

        Private Sub ToggleUi(isEnabled As Boolean)
            txtUsername.Enabled = isEnabled
            txtPassword.Enabled = isEnabled
            chkShowPassword.Enabled = isEnabled
            btnLogin.Enabled = isEnabled
            btnExit.Enabled = isEnabled
            btnLogin.Text = If(isEnabled, "Sign In", "Signing In...")
        End Sub

    End Class

End Namespace
