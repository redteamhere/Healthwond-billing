Imports HealthwondBilling.Models
Imports HealthwondBilling.Services
Imports HealthwondBilling.Utilities

Namespace Forms

    Public Class FrmLogin
        Inherits Form

        Private ReadOnly _authService As AuthService
        Private ReadOnly _dashboardService As DashboardService
        Private ReadOnly _productService As ProductService
        Private ReadOnly _customerService As CustomerService
        Private ReadOnly _supplierService As SupplierService
        Private ReadOnly _billingService As BillingService
        Private ReadOnly _purchaseService As PurchaseService
        Private ReadOnly _purchasePrintService As PurchasePrintService
        Private ReadOnly _invoiceExportService As InvoiceExportService
        Private ReadOnly _reportService As ReportService
        Private ReadOnly _inventoryService As InventoryService
        Private ReadOnly _stockOperationService As StockOperationService
        Private ReadOnly _settlementService As SettlementService
        Private ReadOnly _settingsService As SettingsService
        Private ReadOnly _maintenanceService As MaintenanceService
        Private ReadOnly _accountingService As AccountingService
        Private ReadOnly _clockTimer As New Timer()

        Private ReadOnly lstCompanies As New ListBox()
        Private ReadOnly txtUsername As New TextBox()
        Private ReadOnly txtPassword As New TextBox()
        Private ReadOnly chkShowPassword As New CheckBox()
        Private ReadOnly btnLogin As New Button()
        Private ReadOnly btnDemoAdmin As New Button()
        Private ReadOnly btnDemoStaff As New Button()
        Private ReadOnly btnExit As New Button()
        Private ReadOnly lblStatus As New Label()
        Private ReadOnly lblTopClock As New Label()
        Private ReadOnly lblSelectedCompany As New Label()
        Private ReadOnly lblDateValue As New Label()
        Private ReadOnly lblDayValue As New Label()
        Private ReadOnly lblTimeValue As New Label()
        Private ReadOnly lblShortcutHints As New Label()

        Private _workspaceShadow As Panel
        Private _workspaceCard As Panel
        Private _loginShadow As Panel
        Private _loginCard As Panel
        Private _workspaceHost As Panel

        Public Sub New(authService As AuthService, dashboardService As DashboardService, productService As ProductService, customerService As CustomerService, supplierService As SupplierService, billingService As BillingService, purchaseService As PurchaseService, purchasePrintService As PurchasePrintService, invoiceExportService As InvoiceExportService, reportService As ReportService, inventoryService As InventoryService, stockOperationService As StockOperationService, settlementService As SettlementService, settingsService As SettingsService, maintenanceService As MaintenanceService, accountingService As AccountingService)
            _authService = authService
            _dashboardService = dashboardService
            _productService = productService
            _customerService = customerService
            _supplierService = supplierService
            _billingService = billingService
            _purchaseService = purchaseService
            _purchasePrintService = purchasePrintService
            _invoiceExportService = invoiceExportService
            _reportService = reportService
            _inventoryService = inventoryService
            _stockOperationService = stockOperationService
            _settlementService = settlementService
            _settingsService = settingsService
            _maintenanceService = maintenanceService
            _accountingService = accountingService

            Text = "Healthwond Billing System - Login"
            StartPosition = FormStartPosition.CenterScreen
            MinimumSize = New Size(1340, 860)
            Size = New Size(1560, 920)
            BackColor = ThemePalette.ClassicShellBlue
            Font = New Font("Segoe UI", 10.0F, FontStyle.Regular)
            KeyPreview = True

            BuildLayout()
            ConfigureControls()
            WireEvents()

            AcceptButton = btnLogin
            CancelButton = btnExit
        End Sub

        Private Sub BuildLayout()
            Dim root As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 4,
                .BackColor = ThemePalette.ClassicShellBlue
            }
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 42))
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 34))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 194))

            root.Controls.Add(BuildTopBar(), 0, 0)
            root.Controls.Add(BuildWorkspaceHost(), 0, 1)
            root.Controls.Add(BuildHintBar(), 0, 2)
            root.Controls.Add(BuildFooterPanel(), 0, 3)

            Controls.Add(root)
        End Sub

        Private Function BuildTopBar() As Control
            Dim panel As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = ThemePalette.ClassicShellBlue,
                .Padding = New Padding(10, 4, 10, 4)
            }

            Dim titleLabel As New Label With {
                .Dock = DockStyle.Left,
                .Width = 920,
                .Font = New Font("Segoe UI", 9.75F, FontStyle.Regular),
                .ForeColor = Color.Black,
                .Text = "Healthwond ERP | Pharma Billing | Inventory | Accounts | Licensed Workstation: Sam",
                .TextAlign = ContentAlignment.MiddleLeft
            }

            lblTopClock.Dock = DockStyle.Right
            lblTopClock.Width = 170
            lblTopClock.Font = New Font("Consolas", 12.0F, FontStyle.Bold)
            lblTopClock.ForeColor = Color.White
            lblTopClock.BackColor = ThemePalette.ClassicShellGreen
            lblTopClock.TextAlign = ContentAlignment.MiddleCenter

            panel.Controls.Add(lblTopClock)
            panel.Controls.Add(titleLabel)
            Return panel
        End Function

        Private Function BuildWorkspaceHost() As Control
            _workspaceHost = New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = ThemePalette.ClassicShellCream,
                .Padding = New Padding(18)
            }

            _workspaceShadow = New Panel With {
                .BackColor = ThemePalette.ClassicShellShadow,
                .Size = New Size(1480, 660)
            }

            _workspaceCard = BuildWorkspaceCard()
            _workspaceHost.Controls.Add(_workspaceShadow)
            _workspaceHost.Controls.Add(_workspaceCard)

            AddHandler _workspaceHost.Resize, AddressOf WorkspaceHost_Resize
            Return _workspaceHost
        End Function

        Private Function BuildWorkspaceCard() As Panel
            Dim card As New Panel With {
                .BackColor = ThemePalette.ClassicShellCream,
                .Size = New Size(1480, 660),
                .Padding = New Padding(18),
                .BorderStyle = BorderStyle.FixedSingle
            }

            Dim headerBand As New Label With {
                .Dock = DockStyle.Top,
                .Height = 34,
                .BackColor = ThemePalette.ClassicShellGreen,
                .ForeColor = Color.White,
                .Font = New Font("Segoe UI", 18.0F, FontStyle.Regular),
                .Text = "LIST OF COMPANIES",
                .Padding = New Padding(16, 0, 0, 0),
                .TextAlign = ContentAlignment.MiddleLeft
            }

            Dim contentShell As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2,
                .RowCount = 2,
                .BackColor = ThemePalette.ClassicShellCream,
                .Padding = New Padding(18, 16, 18, 16)
            }
            contentShell.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 420))
            contentShell.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            contentShell.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            contentShell.RowStyles.Add(New RowStyle(SizeType.Absolute, 88))

            contentShell.Controls.Add(BuildCompanyListPanel(), 0, 0)
            contentShell.Controls.Add(BuildLoginZone(), 1, 0)
            contentShell.Controls.Add(BuildSelectedCompanyPrompt(), 0, 1)
            contentShell.Controls.Add(BuildDateInfoPanel(), 1, 1)

            card.Controls.Add(contentShell)
            card.Controls.Add(headerBand)
            Return card
        End Function

        Private Function BuildCompanyListPanel() As Control
            Dim panel As New Panel With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(18),
                .BackColor = ThemePalette.ClassicShellCream
            }

            Dim titleBar As New Label With {
                .Dock = DockStyle.Top,
                .Height = 34,
                .BackColor = ThemePalette.ClassicShellDarkGreen,
                .ForeColor = Color.White,
                .Font = New Font("Consolas", 15.0F, FontStyle.Bold),
                .Text = " AVAILABLE COMPANIES ",
                .TextAlign = ContentAlignment.MiddleLeft
            }

            lstCompanies.Dock = DockStyle.Fill
            lstCompanies.BackColor = Color.FromArgb(250, 250, 242)
            lstCompanies.ForeColor = Color.Black
            lstCompanies.BorderStyle = BorderStyle.FixedSingle
            lstCompanies.Font = New Font("Consolas", 12.0F, FontStyle.Bold)
            lstCompanies.ItemHeight = 24

            Dim helpText As New Label With {
                .Dock = DockStyle.Bottom,
                .Height = 54,
                .ForeColor = ThemePalette.TextPrimary,
                .Font = New Font("Consolas", 10.0F, FontStyle.Regular),
                .Text = "Use Up/Down to choose the active company." & Environment.NewLine & "Press Enter or F5 to sign in.",
                .TextAlign = ContentAlignment.MiddleLeft
            }

            panel.Controls.Add(lstCompanies)
            panel.Controls.Add(helpText)
            panel.Controls.Add(titleBar)
            Return panel
        End Function

        Private Function BuildLoginZone() As Control
            Dim host As New Panel With {.Dock = DockStyle.Fill, .BackColor = ThemePalette.ClassicShellCream}

            _loginShadow = New Panel With {
                .BackColor = ThemePalette.ClassicShellShadow,
                .Size = New Size(620, 360)
            }

            _loginCard = BuildLoginCard()
            host.Controls.Add(_loginShadow)
            host.Controls.Add(_loginCard)

            AddHandler host.Resize,
                Sub()
                    Dim cardLeft As Integer = Math.Max((host.ClientSize.Width - _loginCard.Width) \ 2, 36)
                    Dim cardTop As Integer = Math.Max((host.ClientSize.Height - _loginCard.Height) \ 2 - 10, 40)
                    _loginCard.Location = New Point(cardLeft, cardTop)
                    _loginShadow.Location = New Point(cardLeft + 22, cardTop + 18)
                End Sub

            Return host
        End Function

        Private Function BuildLoginCard() As Panel
            Dim card As New Panel With {
                .BackColor = ThemePalette.ClassicShellPaleBlue,
                .BorderStyle = BorderStyle.FixedSingle,
                .Size = New Size(620, 360)
            }

            Dim titleBar As New Label With {
                .Dock = DockStyle.Top,
                .Height = 32,
                .BackColor = ThemePalette.ClassicShellGreen,
                .ForeColor = Color.White,
                .Font = New Font("Segoe UI", 16.0F, FontStyle.Regular),
                .Padding = New Padding(18, 0, 0, 0),
                .Text = "OPERATOR ACCESS",
                .TextAlign = ContentAlignment.MiddleLeft
            }

            Dim content As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .Padding = New Padding(34, 18, 34, 18)
            }
            content.RowStyles.Add(New RowStyle(SizeType.Absolute, 38))
            content.RowStyles.Add(New RowStyle(SizeType.Absolute, 30))
            content.RowStyles.Add(New RowStyle(SizeType.Absolute, 72))
            content.RowStyles.Add(New RowStyle(SizeType.Absolute, 72))
            content.RowStyles.Add(New RowStyle(SizeType.Absolute, 32))
            content.RowStyles.Add(New RowStyle(SizeType.Absolute, 56))
            content.RowStyles.Add(New RowStyle(SizeType.Absolute, 44))
            content.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            Dim prompt As New Label With {
                .Dock = DockStyle.Fill,
                .BackColor = ThemePalette.ClassicShellDarkGreen,
                .ForeColor = Color.White,
                .Font = New Font("Consolas", 17.0F, FontStyle.Bold),
                .Text = " SIGN IN TO SELECTED COMPANY ",
                .TextAlign = ContentAlignment.MiddleLeft
            }

            lblSelectedCompany.Dock = DockStyle.Fill
            lblSelectedCompany.Font = New Font("Consolas", 11.0F, FontStyle.Regular)
            lblSelectedCompany.ForeColor = ThemePalette.ClassicShellBorder
            lblSelectedCompany.TextAlign = ContentAlignment.MiddleLeft

            content.Controls.Add(prompt, 0, 0)
            content.Controls.Add(lblSelectedCompany, 0, 1)
            content.Controls.Add(CreateClassicInputHost("Operator Name", txtUsername), 0, 2)
            content.Controls.Add(CreateClassicInputHost("Password", txtPassword), 0, 3)

            chkShowPassword.Text = "Show Password"
            chkShowPassword.AutoSize = True
            chkShowPassword.Font = New Font("Consolas", 10.0F, FontStyle.Regular)
            chkShowPassword.ForeColor = ThemePalette.ClassicShellBorder
            chkShowPassword.BackColor = Color.Transparent
            content.Controls.Add(chkShowPassword, 0, 4)

            Dim buttonRow As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 4
            }
            buttonRow.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 31.0F))
            buttonRow.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 23.0F))
            buttonRow.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 23.0F))
            buttonRow.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 23.0F))

            ConfigureClassicButton(btnLogin, "Sign In", True, AddressOf btnLogin_Click)
            ConfigureClassicButton(btnDemoAdmin, "Demo Admin", False, AddressOf btnDemoAdmin_Click)
            ConfigureClassicButton(btnDemoStaff, "Demo Staff", False, AddressOf btnDemoStaff_Click)
            ConfigureClassicButton(btnExit, "Exit", False, AddressOf btnExit_Click)

            buttonRow.Controls.Add(btnLogin, 0, 0)
            buttonRow.Controls.Add(btnDemoAdmin, 1, 0)
            buttonRow.Controls.Add(btnDemoStaff, 2, 0)
            buttonRow.Controls.Add(btnExit, 3, 0)

            Dim optionsLabel As New Label With {
                .Dock = DockStyle.Fill,
                .Font = New Font("Consolas", 10.5F, FontStyle.Regular),
                .ForeColor = ThemePalette.ClassicShellBorder,
                .Text = "F2 Demo Admin    F3 Demo Staff    F5 Sign In    F8 Show / Hide Password",
                .TextAlign = ContentAlignment.MiddleLeft
            }

            lblStatus.Dock = DockStyle.Fill
            lblStatus.Font = New Font("Consolas", 10.5F, FontStyle.Bold)
            lblStatus.ForeColor = ThemePalette.DangerRed
            lblStatus.TextAlign = ContentAlignment.MiddleLeft

            content.Controls.Add(buttonRow, 0, 5)
            content.Controls.Add(optionsLabel, 0, 6)
            content.Controls.Add(lblStatus, 0, 7)

            card.Controls.Add(content)
            card.Controls.Add(titleBar)
            Return card
        End Function

        Private Function CreateClassicInputHost(labelText As String, editorControl As Control) As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill}

            Dim caption As New Label With {
                .Dock = DockStyle.Top,
                .Height = 24,
                .Font = New Font("Consolas", 11.0F, FontStyle.Bold),
                .ForeColor = ThemePalette.ClassicShellBorder,
                .Text = labelText
            }

            editorControl.Dock = DockStyle.Bottom
            editorControl.Height = 38
            panel.Controls.Add(editorControl)
            panel.Controls.Add(caption)
            Return panel
        End Function

        Private Function BuildSelectedCompanyPrompt() As Control
            Dim panel As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = ThemePalette.ClassicShellCream,
                .Padding = New Padding(18, 18, 18, 0)
            }

            Dim labelText As New Label With {
                .Dock = DockStyle.Left,
                .Width = 220,
                .Font = New Font("Consolas", 18.0F, FontStyle.Regular),
                .ForeColor = Color.Black,
                .Text = "Selected Company ?",
                .TextAlign = ContentAlignment.MiddleLeft
            }

            Dim valueLabel As New Label With {
                .Dock = DockStyle.Fill,
                .Font = New Font("Consolas", 18.0F, FontStyle.Bold),
                .ForeColor = ThemePalette.ClassicShellDarkGreen,
                .BorderStyle = BorderStyle.FixedSingle,
                .Padding = New Padding(10, 0, 10, 0),
                .TextAlign = ContentAlignment.MiddleLeft
            }
            AddHandler lstCompanies.SelectedIndexChanged,
                Sub()
                    valueLabel.Text = GetSelectedCompanyName()
                    lblSelectedCompany.Text = $"Company : {GetSelectedCompanyName()}"
                End Sub

            panel.Controls.Add(valueLabel)
            panel.Controls.Add(labelText)
            Return panel
        End Function

        Private Function BuildDateInfoPanel() As Control
            Dim panel As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = ThemePalette.ClassicShellCream,
                .Padding = New Padding(20, 12, 12, 0)
            }

            Dim table As New TableLayoutPanel With {
                .Dock = DockStyle.Right,
                .Width = 390,
                .ColumnCount = 2
            }
            table.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 110))
            table.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            table.RowStyles.Add(New RowStyle(SizeType.Absolute, 30))
            table.RowStyles.Add(New RowStyle(SizeType.Absolute, 30))
            table.RowStyles.Add(New RowStyle(SizeType.Absolute, 30))

            table.Controls.Add(CreateClassicInfoLabel("Date :"), 0, 0)
            table.Controls.Add(CreateClassicValueLabel(lblDateValue), 1, 0)
            table.Controls.Add(CreateClassicInfoLabel("Day  :"), 0, 1)
            table.Controls.Add(CreateClassicValueLabel(lblDayValue), 1, 1)
            table.Controls.Add(CreateClassicInfoLabel("Time :"), 0, 2)
            table.Controls.Add(CreateClassicValueLabel(lblTimeValue), 1, 2)

            panel.Controls.Add(table)
            Return panel
        End Function

        Private Function CreateClassicInfoLabel(text As String) As Label
            Return New Label With {
                .Dock = DockStyle.Fill,
                .Font = New Font("Segoe UI", 16.0F, FontStyle.Regular),
                .ForeColor = Color.Black,
                .Text = text,
                .TextAlign = ContentAlignment.MiddleLeft
            }
        End Function

        Private Function CreateClassicValueLabel(target As Label) As Label
            target.Dock = DockStyle.Fill
            target.Font = New Font("Segoe UI", 17.0F, FontStyle.Regular)
            target.ForeColor = Color.Black
            target.TextAlign = ContentAlignment.MiddleLeft
            Return target
        End Function

        Private Function BuildHintBar() As Control
            Dim panel As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = ThemePalette.ClassicShellHintBar,
                .Padding = New Padding(10, 0, 10, 0)
            }

            lblShortcutHints.Dock = DockStyle.Fill
            lblShortcutHints.Font = New Font("Segoe UI", 10.5F, FontStyle.Regular)
            lblShortcutHints.ForeColor = Color.White
            lblShortcutHints.Text = "F2-Demo Admin   F3-Demo Staff   Enter/F5-Sign In   F8-Show Password   Up/Down-Select Company   Esc-Exit"
            lblShortcutHints.TextAlign = ContentAlignment.MiddleLeft

            panel.Controls.Add(lblShortcutHints)
            Return panel
        End Function

        Private Function BuildFooterPanel() As Control
            Dim panel As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = ThemePalette.ClassicShellFooter,
                .Padding = New Padding(16, 10, 16, 10)
            }

            Dim layout As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 3
            }
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 42.0F))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 22.0F))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 36.0F))

            Dim supportLabel As New Label With {
                .Dock = DockStyle.Fill,
                .Font = New Font("Segoe UI", 10.5F, FontStyle.Regular),
                .ForeColor = ThemePalette.TextPrimary,
                .Text = "Developed & Managed By :" & Environment.NewLine &
                        "HEALTHWOND BILLING SYSTEM" & Environment.NewLine &
                        "Support : support@healthwond.local" & Environment.NewLine &
                        "Desktop ERP Workspace for Pharmaceutical Billing",
                .TextAlign = ContentAlignment.TopLeft
            }

            Dim centerPanel As New Panel With {.Dock = DockStyle.Fill, .BackColor = ThemePalette.ClassicShellFooter}
            Dim badge As New Label With {
                .Size = New Size(180, 112),
                .BorderStyle = BorderStyle.FixedSingle,
                .BackColor = ThemePalette.ClassicShellCream,
                .Font = New Font("Segoe UI", 13.0F, FontStyle.Bold),
                .ForeColor = ThemePalette.BrandBlue,
                .Text = "Healthwond" & Environment.NewLine & "ERP Workspace" & Environment.NewLine & "Desktop Edition",
                .TextAlign = ContentAlignment.MiddleCenter
            }
            centerPanel.Controls.Add(badge)
            AddHandler centerPanel.Resize,
                Sub()
                    badge.Location = New Point(Math.Max((centerPanel.ClientSize.Width - badge.Width) \ 2, 0), Math.Max((centerPanel.ClientSize.Height - badge.Height) \ 2, 0))
                End Sub

            Dim rightLabel As New Label With {
                .Dock = DockStyle.Fill,
                .Font = New Font("Segoe UI", 13.0F, FontStyle.Regular),
                .ForeColor = ThemePalette.TextPrimary,
                .Text = "Authorised Access :" & Environment.NewLine &
                        "Admin / Staff Accounts" & Environment.NewLine & Environment.NewLine &
                        "Classic startup shell inspired by legacy desktop ERP workflows.",
                .TextAlign = ContentAlignment.TopLeft
            }

            layout.Controls.Add(supportLabel, 0, 0)
            layout.Controls.Add(centerPanel, 1, 0)
            layout.Controls.Add(rightLabel, 2, 0)
            panel.Controls.Add(layout)
            Return panel
        End Function

        Private Sub ConfigureControls()
            txtUsername.Text = "admin"
            txtUsername.Font = New Font("Consolas", 13.0F, FontStyle.Bold)
            txtUsername.BackColor = Color.White
            txtUsername.BorderStyle = BorderStyle.FixedSingle

            txtPassword.Font = New Font("Consolas", 13.0F, FontStyle.Bold)
            txtPassword.BackColor = Color.White
            txtPassword.BorderStyle = BorderStyle.FixedSingle
            txtPassword.UseSystemPasswordChar = True

            _clockTimer.Interval = 1000
            UpdateClock()
        End Sub

        Private Sub WireEvents()
            AddHandler Load, AddressOf FrmLogin_Load
            AddHandler btnLogin.Click, AddressOf btnLogin_Click
            AddHandler btnDemoAdmin.Click, AddressOf btnDemoAdmin_Click
            AddHandler btnDemoStaff.Click, AddressOf btnDemoStaff_Click
            AddHandler btnExit.Click, AddressOf btnExit_Click
            AddHandler chkShowPassword.CheckedChanged, AddressOf chkShowPassword_CheckedChanged
            AddHandler _clockTimer.Tick, AddressOf ClockTimer_Tick
        End Sub

        Private Async Sub FrmLogin_Load(sender As Object, e As EventArgs)
            _clockTimer.Start()
            Await LoadStartupContextAsync()
        End Sub

        Private Async Function LoadStartupContextAsync() As Task
            Try
                Dim profile As AppSettingsProfile = Await _settingsService.LoadAsync()
                Dim companyName As String = If(profile.CompanyName, String.Empty).Trim()
                If companyName = String.Empty Then
                    companyName = "Healthwond Pharmacy"
                End If

                lstCompanies.Items.Clear()
                lstCompanies.Items.Add(companyName)
                lstCompanies.Items.Add(companyName & " - Training")
                lstCompanies.SelectedIndex = 0
                lblStatus.ForeColor = ThemePalette.AccentGreen
                lblStatus.Text = "Startup workspace ready. Select company and sign in."
            Catch ex As Exception
                AppLogger.Error("Startup shell initialization failed.", ex)
                lstCompanies.Items.Clear()
                lstCompanies.Items.Add("Healthwond Pharmacy")
                lstCompanies.SelectedIndex = 0
                lblStatus.ForeColor = ThemePalette.DangerRed
                lblStatus.Text = "Settings could not be loaded. Using fallback company profile."
            End Try
        End Function

        Private Sub WorkspaceHost_Resize(sender As Object, e As EventArgs)
            If _workspaceCard Is Nothing OrElse _workspaceShadow Is Nothing Then
                Return
            End If

            Dim cardWidth As Integer = Math.Min(Math.Max(_workspaceHost.ClientSize.Width - 72, 1200), 1480)
            Dim cardHeight As Integer = Math.Min(Math.Max(_workspaceHost.ClientSize.Height - 56, 560), 680)
            _workspaceCard.Size = New Size(cardWidth, cardHeight)
            _workspaceShadow.Size = _workspaceCard.Size

            Dim left As Integer = Math.Max((_workspaceHost.ClientSize.Width - _workspaceCard.Width) \ 2 - 8, 12)
            Dim top As Integer = Math.Max((_workspaceHost.ClientSize.Height - _workspaceCard.Height) \ 2 - 8, 12)

            _workspaceCard.Location = New Point(left, top)
            _workspaceShadow.Location = New Point(left + 18, top + 18)
        End Sub

        Private Sub ConfigureClassicButton(button As Button, text As String, isPrimary As Boolean, handler As EventHandler)
            If isPrimary Then
                button.FlatStyle = FlatStyle.Flat
                button.FlatAppearance.BorderSize = 1
                button.FlatAppearance.BorderColor = ThemePalette.ClassicShellBorder
                button.BackColor = ThemePalette.ClassicShellDarkGreen
                button.ForeColor = Color.White
            Else
                button.FlatStyle = FlatStyle.Flat
                button.FlatAppearance.BorderSize = 1
                button.FlatAppearance.BorderColor = ThemePalette.ClassicShellBorder
                button.BackColor = ThemePalette.ClassicShellCream
                button.ForeColor = ThemePalette.ClassicShellBorder
            End If

            button.Cursor = Cursors.Hand
            button.Font = New Font("Consolas", 10.5F, FontStyle.Bold)
            button.Height = 42
            button.Text = text
            AddHandler button.Click, handler
        End Sub

        Private Sub btnDemoAdmin_Click(sender As Object, e As EventArgs)
            txtUsername.Text = "admin"
            txtPassword.Text = "Admin@123"
            lblStatus.ForeColor = ThemePalette.ClassicShellBorder
            lblStatus.Text = "Demo admin credentials loaded. Press Sign In."
            txtPassword.Focus()
            txtPassword.SelectionStart = txtPassword.TextLength
        End Sub

        Private Sub btnDemoStaff_Click(sender As Object, e As EventArgs)
            txtUsername.Text = "staff"
            txtPassword.Text = "Staff@123"
            lblStatus.ForeColor = ThemePalette.ClassicShellBorder
            lblStatus.Text = "Demo staff credentials loaded. Press Sign In."
            txtPassword.Focus()
            txtPassword.SelectionStart = txtPassword.TextLength
        End Sub

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
                    lblStatus.ForeColor = ThemePalette.DangerRed
                    lblStatus.Text = result.Message
                    Return
                End If

                SessionManager.StartSession(result.User)

                Using dashboard As New FrmDashboard(_dashboardService, _productService, _customerService, _supplierService, _billingService, _purchaseService, _purchasePrintService, _invoiceExportService, _reportService, _inventoryService, _stockOperationService, _settlementService, _settingsService, _maintenanceService, _accountingService)
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
                    Else
                        Close()
                    End If
                End Using
            Catch ex As Exception
                AppLogger.Error("Login workflow failed.", ex)
                lblStatus.ForeColor = ThemePalette.DangerRed
                lblStatus.Text = "Unable to complete the sign-in request."
            Finally
                ToggleUi(True)
            End Try
        End Function

        Private Sub ToggleUi(isEnabled As Boolean)
            lstCompanies.Enabled = isEnabled
            txtUsername.Enabled = isEnabled
            txtPassword.Enabled = isEnabled
            chkShowPassword.Enabled = isEnabled
            btnLogin.Enabled = isEnabled
            btnDemoAdmin.Enabled = isEnabled
            btnDemoStaff.Enabled = isEnabled
            btnExit.Enabled = isEnabled
            btnLogin.Text = If(isEnabled, "Sign In", "Signing In...")
        End Sub

        Private Sub ClockTimer_Tick(sender As Object, e As EventArgs)
            UpdateClock()
        End Sub

        Private Sub UpdateClock()
            Dim nowValue As DateTime = DateTime.Now
            lblTopClock.Text = nowValue.ToString("HH:mm:ss")
            lblDateValue.Text = nowValue.ToString("dd MMM, yyyy")
            lblDayValue.Text = nowValue.ToString("dddd")
            lblTimeValue.Text = nowValue.ToString("hh:mm:ss tt")
        End Sub

        Private Function GetSelectedCompanyName() As String
            If lstCompanies.SelectedItem Is Nothing Then
                Return "Healthwond Pharmacy"
            End If

            Return Convert.ToString(lstCompanies.SelectedItem)
        End Function

        Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
            Select Case keyData
                Case Keys.F2
                    btnDemoAdmin.PerformClick()
                    Return True
                Case Keys.F3
                    btnDemoStaff.PerformClick()
                    Return True
                Case Keys.F5
                    btnLogin.PerformClick()
                    Return True
                Case Keys.F8
                    chkShowPassword.Checked = Not chkShowPassword.Checked
                    Return True
                Case Keys.Escape
                    Close()
                    Return True
            End Select

            Return MyBase.ProcessCmdKey(msg, keyData)
        End Function

        Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
            _clockTimer.Stop()
            MyBase.OnFormClosing(e)
        End Sub

    End Class

End Namespace
