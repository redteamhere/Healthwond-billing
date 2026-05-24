Imports HealthwondBilling.Controls
Imports HealthwondBilling.Models
Imports HealthwondBilling.Services
Imports HealthwondBilling.Utilities

Namespace Forms

    Public Class FrmDashboard
        Inherits Form

        Private ReadOnly _dashboardService As DashboardService
        Private ReadOnly _productService As ProductService
        Private ReadOnly _customerService As CustomerService
        Private ReadOnly _supplierService As SupplierService
        Private ReadOnly _billingService As BillingService
        Private ReadOnly _purchaseService As PurchaseService
        Private ReadOnly _invoiceExportService As InvoiceExportService
        Private ReadOnly _reportService As ReportService
        Private ReadOnly _clockTimer As Timer

        Private lblGreeting As Label
        Private lblRole As Label
        Private lblClock As Label
        Private btnSettings As Button

        Private salesCard As StatCardControl
        Private stockCard As StatCardControl
        Private expiryCard As StatCardControl
        Private lowStockCard As StatCardControl
        Private pendingCard As StatCardControl

        Public Property RequestedLogout As Boolean

        Public Sub New(dashboardService As DashboardService, productService As ProductService, customerService As CustomerService, supplierService As SupplierService, billingService As BillingService, purchaseService As PurchaseService, invoiceExportService As InvoiceExportService, reportService As ReportService)
            _dashboardService = dashboardService
            _productService = productService
            _customerService = customerService
            _supplierService = supplierService
            _billingService = billingService
            _purchaseService = purchaseService
            _invoiceExportService = invoiceExportService
            _reportService = reportService

            Text = "Healthwond Billing System - Dashboard"
            WindowState = FormWindowState.Maximized
            MinimumSize = New Size(1220, 760)
            BackColor = ThemePalette.AppBackground
            Font = New Font("Segoe UI", 10.0F, FontStyle.Regular)

            Dim root As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2
            }
            root.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 240))
            root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))

            Dim sidebar As Panel = BuildSidebar()
            Dim content As Panel = BuildContentPanel(lblGreeting, lblRole, lblClock, salesCard, stockCard, expiryCard, lowStockCard, pendingCard)

            root.Controls.Add(sidebar, 0, 0)
            root.Controls.Add(content, 1, 0)
            Controls.Add(root)

            ConfigureAccess()

            _clockTimer = New Timer With {.Interval = 1000}
            AddHandler _clockTimer.Tick, AddressOf ClockTimer_Tick
            AddHandler Load, AddressOf FrmDashboard_Load
        End Sub

        Private Function BuildSidebar() As Panel
            Dim sidebar As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = ThemePalette.BrandBlue,
                .Padding = New Padding(18)
            }

            Dim title As New Label With {
                .Dock = DockStyle.Top,
                .Height = 86,
                .Font = New Font("Segoe UI Semibold", 18.0F, FontStyle.Bold),
                .ForeColor = Color.White,
                .Text = "Healthwond" & Environment.NewLine & "Dashboard",
                .TextAlign = ContentAlignment.MiddleLeft
            }

            Dim navPanel As New FlowLayoutPanel With {
                .Dock = DockStyle.Top,
                .FlowDirection = FlowDirection.TopDown,
                .WrapContents = False,
                .AutoSize = True,
                .BackColor = Color.Transparent,
                .Padding = New Padding(0, 16, 0, 0)
            }

            navPanel.Controls.Add(CreateSidebarButton("F1  Billing", "Billing"))
            navPanel.Controls.Add(CreateSidebarButton("F2  Products", "Products"))
            navPanel.Controls.Add(CreateSidebarButton("F3  Customers", "Customers"))
            navPanel.Controls.Add(CreateSidebarButton("F4  Purchases", "Purchases"))
            navPanel.Controls.Add(CreateSidebarButton("F5  Refresh Metrics", "Refresh"))
            navPanel.Controls.Add(CreateSidebarButton("F6  Reports", "Reports"))

            btnSettings = CreateSidebarButton("Admin  Settings", "Settings")
            navPanel.Controls.Add(btnSettings)

            Dim logoutButton As Button = CreateSidebarButton("Ctrl+L  Logout", "Logout")
            AddHandler logoutButton.Click, AddressOf LogoutButton_Click
            navPanel.Controls.Add(logoutButton)

            Dim footer As New Label With {
                .Dock = DockStyle.Bottom,
                .Height = 84,
                .Font = New Font("Segoe UI", 9.25F, FontStyle.Regular),
                .ForeColor = Color.FromArgb(207, 219, 232),
                .Text = "Current build includes authentication, dashboard analytics, masters, billing, and live purchase stock-in workflows.",
                .TextAlign = ContentAlignment.BottomLeft
            }

            sidebar.Controls.Add(footer)
            sidebar.Controls.Add(navPanel)
            sidebar.Controls.Add(title)

            Return sidebar
        End Function

        Private Function CreateSidebarButton(text As String, tagValue As String) As Button
            Dim button As New Button With {
                .Text = text,
                .Tag = tagValue,
                .Width = 198,
                .Height = 46,
                .FlatStyle = FlatStyle.Flat,
                .TextAlign = ContentAlignment.MiddleLeft,
                .BackColor = Color.FromArgb(39, 58, 98),
                .ForeColor = Color.White,
                .Margin = New Padding(0, 0, 0, 10),
                .Cursor = Cursors.Hand
            }
            button.FlatAppearance.BorderSize = 0
            AddHandler button.Click, AddressOf ModuleButton_Click
            Return button
        End Function

        Private Function BuildContentPanel(ByRef greetingLabel As Label, ByRef roleLabel As Label, ByRef clockLabel As Label, ByRef todaySalesCard As StatCardControl, ByRef totalStockCard As StatCardControl, ByRef expiryAlertsCard As StatCardControl, ByRef lowStockAlertsCard As StatCardControl, ByRef pendingPaymentsCard As StatCardControl) As Panel
            Dim content As New Panel With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(28),
                .BackColor = ThemePalette.AppBackground
            }

            Dim headerPanel As New Panel With {
                .Dock = DockStyle.Top,
                .Height = 112,
                .BackColor = ThemePalette.AppBackground
            }

            greetingLabel = New Label With {
                .Dock = DockStyle.Top,
                .Height = 52,
                .Font = New Font("Segoe UI Semibold", 24.0F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary,
                .TextAlign = ContentAlignment.MiddleLeft
            }

            roleLabel = New Label With {
                .Dock = DockStyle.Top,
                .Height = 28,
                .Font = New Font("Segoe UI", 10.5F, FontStyle.Regular),
                .ForeColor = ThemePalette.TextMuted,
                .TextAlign = ContentAlignment.MiddleLeft
            }

            clockLabel = New Label With {
                .Dock = DockStyle.Right,
                .Width = 260,
                .Font = New Font("Segoe UI Semibold", 11.0F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary,
                .TextAlign = ContentAlignment.TopRight
            }

            headerPanel.Controls.Add(clockLabel)
            headerPanel.Controls.Add(roleLabel)
            headerPanel.Controls.Add(greetingLabel)

            Dim cardFlow As New FlowLayoutPanel With {
                .Dock = DockStyle.Top,
                .Height = 160,
                .WrapContents = True,
                .FlowDirection = FlowDirection.LeftToRight,
                .BackColor = Color.Transparent
            }

            todaySalesCard = CreateCard("Today's Sales", ThemePalette.AccentGreen, "Rs. 0.00", "Net sales billed today")
            totalStockCard = CreateCard("Total Stock", ThemePalette.AccentBlue, "0", "Saleable units on hand")
            expiryAlertsCard = CreateCard("Expiry Alerts", ThemePalette.WarningAmber, "0", "Products expiring within 60 days")
            lowStockAlertsCard = CreateCard("Low Stock", ThemePalette.DangerRed, "0", "Items at or below reorder threshold")
            pendingPaymentsCard = CreateCard("Pending Payments", ThemePalette.PurpleGray, "Rs. 0.00", "Outstanding customer balance")

            cardFlow.Controls.Add(todaySalesCard)
            cardFlow.Controls.Add(totalStockCard)
            cardFlow.Controls.Add(expiryAlertsCard)
            cardFlow.Controls.Add(lowStockAlertsCard)
            cardFlow.Controls.Add(pendingPaymentsCard)

            Dim workspaceCard As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = ThemePalette.CardBackground,
                .Padding = New Padding(26)
            }
            UiStyler.StyleCard(workspaceCard)

            Dim workspaceTitle As New Label With {
                .Dock = DockStyle.Top,
                .Height = 38,
                .Font = New Font("Segoe UI Semibold", 16.0F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary,
                .Text = "Current module status"
            }

            Dim workspaceInfo As New Label With {
                .Dock = DockStyle.Fill,
                .Font = New Font("Segoe UI", 10.25F, FontStyle.Regular),
                .ForeColor = ThemePalette.TextMuted,
                .Text = String.Join(
                    Environment.NewLine,
                    "- Authentication is active with hashed passwords and role-aware sessions.",
                    "- Product master supports search, CRUD, barcode, pricing, GST, expiry, and stock adjustments.",
                    "- Customer master supports search, CRUD, contact details, GSTINs, license numbers, and dues.",
                    "- Billing now supports customer selection, product lines, GST totals, invoice save, and Excel/PDF export with print actions.",
                    "- Purchases now support supplier management, batch stock-in, ledger posting, and payable accumulation.",
                    "- Reports now support sales, purchases, GST, stock, receivables, and profit summary views.",
                    "- Settings and deeper inventory views will follow in the next modules."),
                .TextAlign = ContentAlignment.TopLeft
            }

            workspaceCard.Controls.Add(workspaceInfo)
            workspaceCard.Controls.Add(workspaceTitle)

            content.Controls.Add(workspaceCard)
            content.Controls.Add(cardFlow)
            content.Controls.Add(headerPanel)

            Return content
        End Function

        Private Function CreateCard(title As String, accentColor As Color, value As String, subtitle As String) As StatCardControl
            Return New StatCardControl With {
                .CardTitle = title,
                .AccentColor = accentColor,
                .ValueText = value,
                .SubtitleText = subtitle
            }
        End Function

        Private Async Sub FrmDashboard_Load(sender As Object, e As EventArgs)
            UpdateClock()
            _clockTimer.Start()
            Await RefreshSummaryAsync()
        End Sub

        Private Sub ConfigureAccess()
            Dim currentUser As UserAccount = SessionManager.CurrentUser
            If currentUser Is Nothing Then
                Return
            End If

            lblGreeting.Text = $"Welcome back, {currentUser.FullName}"
            lblRole.Text = $"Signed in as {currentUser.Role.ToFriendlyText()} | Username: {currentUser.Username}"

            btnSettings.Enabled = SessionManager.IsAdmin
            btnSettings.BackColor = If(SessionManager.IsAdmin, Color.FromArgb(39, 58, 98), Color.FromArgb(76, 88, 112))
        End Sub

        Private Async Function RefreshSummaryAsync() As Task
            Try
                Dim summary As DashboardSummary = Await _dashboardService.GetSummaryAsync()
                salesCard.ValueText = $"Rs. {summary.TodaySales:N2}"
                stockCard.ValueText = summary.TotalStockUnits.ToString("N0")
                expiryCard.ValueText = summary.ExpiryAlerts.ToString("N0")
                lowStockCard.ValueText = summary.LowStockAlerts.ToString("N0")
                pendingCard.ValueText = $"Rs. {summary.PendingPayments:N2}"
            Catch ex As Exception
                AppLogger.Error("Failed to refresh dashboard metrics.", ex)
                MessageBox.Show(
                    "Dashboard metrics could not be loaded. Review the application logs for details.",
                    "Dashboard Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning)
            End Try
        End Function

        Private Sub ClockTimer_Tick(sender As Object, e As EventArgs)
            UpdateClock()
        End Sub

        Private Sub UpdateClock()
            lblClock.Text = DateTime.Now.ToString("dddd, dd MMM yyyy" & Environment.NewLine & "hh:mm:ss tt")
        End Sub

        Private Async Sub ModuleButton_Click(sender As Object, e As EventArgs)
            Dim button As Button = CType(sender, Button)
            Dim tagValue As String = Convert.ToString(button.Tag)

            Select Case tagValue
                Case "Products"
                    OpenProductsDialog()
                    Await RefreshSummaryAsync()
                Case "Customers"
                    OpenCustomersDialog()
                    Await RefreshSummaryAsync()
                Case "Billing"
                    OpenBillingDialog()
                    Await RefreshSummaryAsync()
                Case "Purchases"
                    OpenPurchasesDialog()
                    Await RefreshSummaryAsync()
                Case "Refresh"
                    Await RefreshSummaryAsync()
                Case "Reports"
                    OpenReportsDialog()
                Case "Settings"
                    If Not SessionManager.IsAdmin Then
                        MessageBox.Show("Settings access is limited to administrator accounts.", "Access Restricted", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Return
                    End If

                    ShowFutureModuleMessage("Settings")
                Case "Logout"
                    LogoutButton_Click(sender, e)
                Case Else
                    ShowFutureModuleMessage(tagValue)
            End Select
        End Sub

        Private Sub OpenProductsDialog()
            Using form As New FrmProducts(_productService)
                form.ShowDialog(Me)
            End Using
        End Sub

        Private Sub OpenBillingDialog()
            Using form As New FrmBilling(_billingService, _invoiceExportService)
                form.ShowDialog(Me)
            End Using
        End Sub

        Private Sub OpenPurchasesDialog()
            Using form As New FrmPurchases(_purchaseService, _supplierService)
                form.ShowDialog(Me)
            End Using
        End Sub

        Private Sub OpenReportsDialog()
            Using form As New FrmReports(_reportService)
                form.ShowDialog(Me)
            End Using
        End Sub

        Private Sub OpenCustomersDialog()
            Using form As New FrmCustomers(_customerService)
                form.ShowDialog(Me)
            End Using
        End Sub

        Private Sub ShowFutureModuleMessage(moduleName As String)
            MessageBox.Show(
                $"{moduleName} will be delivered in the next project module. The current build focuses on the platform foundation and authentication flow.",
                "Module Pending",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information)
        End Sub

        Private Sub LogoutButton_Click(sender As Object, e As EventArgs)
            RequestedLogout = True
            Close()
        End Sub

        Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
            If Not RequestedLogout AndAlso e.CloseReason = CloseReason.UserClosing Then
                Dim result As DialogResult = MessageBox.Show(
                    "Close the Healthwond Billing System?",
                    "Confirm Exit",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question)

                If result = DialogResult.No Then
                    e.Cancel = True
                    Return
                End If
            End If

            _clockTimer.Stop()
            MyBase.OnFormClosing(e)
        End Sub

        Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
            Select Case keyData
                Case Keys.F1
                    OpenBillingDialog()
                    Return True
                Case Keys.F2
                    OpenProductsDialog()
                    Return True
                Case Keys.F3
                    OpenCustomersDialog()
                    Return True
                Case Keys.F4
                    OpenPurchasesDialog()
                    Return True
                Case Keys.F5
                    Dim refreshTask As Task = RefreshSummaryAsync()
                    Return True
                Case (Keys.Control Or Keys.L)
                    LogoutButton_Click(Me, EventArgs.Empty)
                    Return True
                Case Keys.F6
                    OpenReportsDialog()
                    Return True
            End Select

            Return MyBase.ProcessCmdKey(msg, keyData)
        End Function

    End Class

End Namespace
