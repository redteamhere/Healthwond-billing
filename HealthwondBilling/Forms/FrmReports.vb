Imports HealthwondBilling.Models
Imports HealthwondBilling.Services
Imports HealthwondBilling.Utilities
Imports System.Diagnostics
Imports System.IO

Namespace Forms

    Public Class FrmReports
        Inherits Form

        Private ReadOnly _reportService As ReportService

        Private ReadOnly dtpFromDate As New DateTimePicker()
        Private ReadOnly dtpToDate As New DateTimePicker()
        Private ReadOnly btnRefresh As New Button()
        Private ReadOnly btnExport As New Button()
        Private ReadOnly btnOpenFolder As New Button()
        Private ReadOnly btnClose As New Button()
        Private ReadOnly tabs As New TabControl()
        Private ReadOnly lblStatus As New Label()

        Private ReadOnly dgvSales As New DataGridView()
        Private ReadOnly dgvPurchases As New DataGridView()
        Private ReadOnly dgvGst As New DataGridView()
        Private ReadOnly dgvStock As New DataGridView()
        Private ReadOnly dgvOutstanding As New DataGridView()
        Private ReadOnly dgvCustomerAging As New DataGridView()
        Private ReadOnly dgvSupplierAging As New DataGridView()

        Private ReadOnly overviewLabels As New Dictionary(Of String, Label)()
        Private ReadOnly profitLabels As New Dictionary(Of String, Label)()
        Private _overview As ReportOverview
        Private _profitLossReport As ProfitLossReport
        Private _isBusy As Boolean

        Public Sub New(reportService As ReportService)
            _reportService = reportService

            Text = "Healthwond Billing System - Reports"
            StartPosition = FormStartPosition.CenterParent
            WindowState = FormWindowState.Maximized
            MinimumSize = New Size(1380, 860)
            BackColor = ThemePalette.AppBackground
            Font = New Font("Segoe UI", 10.0F, FontStyle.Regular)
            KeyPreview = True

            BuildLayout()
            ConfigureControls()
            WireEvents()
        End Sub

        Private Sub BuildLayout()
            Dim root As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 4,
                .Padding = New Padding(22),
                .BackColor = ThemePalette.AppBackground
            }
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 78))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 116))
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 36))

            root.Controls.Add(BuildHeaderPanel(), 0, 0)
            root.Controls.Add(BuildFilterPanel(), 0, 1)
            root.Controls.Add(BuildTabPanel(), 0, 2)

            lblStatus.Dock = DockStyle.Fill
            lblStatus.Font = New Font("Segoe UI Semibold", 9.5F, FontStyle.Bold)
            lblStatus.TextAlign = ContentAlignment.MiddleLeft
            root.Controls.Add(lblStatus, 0, 3)

            Controls.Add(root)
        End Sub

        Private Function BuildHeaderPanel() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = ThemePalette.AppBackground}

            Dim title As New Label With {
                .Dock = DockStyle.Top,
                .Height = 40,
                .Text = "Reports",
                .Font = New Font("Segoe UI Semibold", 24.0F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary
            }

            Dim subtitle As New Label With {
                .Dock = DockStyle.Fill,
                .Text = "Review operational metrics, aging balances, sales, purchases, GST, stock, receivables, and gross-profit indicators from the live SQLite data set.",
                .Font = New Font("Segoe UI", 10.5F, FontStyle.Regular),
                .ForeColor = ThemePalette.TextMuted
            }

            panel.Controls.Add(subtitle)
            panel.Controls.Add(title)
            Return panel
        End Function

        Private Function BuildFilterPanel() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)

            Dim layout As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 6
            }
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 180))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 180))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 110))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 120))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 120))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))

            layout.Controls.Add(CreateInputHost("From Date", dtpFromDate), 0, 0)
            layout.Controls.Add(CreateInputHost("To Date", dtpToDate), 1, 0)

            UiStyler.StylePrimaryButton(btnRefresh)
            btnRefresh.Text = "Refresh"
            btnRefresh.Width = 100
            btnRefresh.Margin = New Padding(0, 24, 12, 0)

            UiStyler.StyleSecondaryButton(btnExport)
            btnExport.Text = "Export"
            btnExport.Width = 100
            btnExport.Margin = New Padding(0, 24, 12, 0)

            UiStyler.StyleSecondaryButton(btnOpenFolder)
            btnOpenFolder.Text = "Open Folder"
            btnOpenFolder.Width = 108
            btnOpenFolder.Margin = New Padding(0, 24, 12, 0)

            UiStyler.StyleSecondaryButton(btnClose)
            btnClose.Text = "Close"
            btnClose.Width = 90
            btnClose.Margin = New Padding(0, 24, 0, 0)

            layout.Controls.Add(btnRefresh, 2, 0)
            layout.Controls.Add(btnExport, 3, 0)
            layout.Controls.Add(btnOpenFolder, 4, 0)
            layout.Controls.Add(btnClose, 5, 0)

            panel.Controls.Add(layout)
            Return panel
        End Function

        Private Function BuildTabPanel() As Control
            tabs.Dock = DockStyle.Fill
            tabs.Appearance = TabAppearance.Normal
            tabs.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)

            tabs.TabPages.Add(CreateOverviewTab())
            tabs.TabPages.Add(CreateGridTab("Sales", dgvSales))
            tabs.TabPages.Add(CreateGridTab("Purchases", dgvPurchases))
            tabs.TabPages.Add(CreateGridTab("GST", dgvGst))
            tabs.TabPages.Add(CreateGridTab("Stock", dgvStock))
            tabs.TabPages.Add(CreateGridTab("Outstanding", dgvOutstanding))
            tabs.TabPages.Add(CreateGridTab("Customer Aging", dgvCustomerAging))
            tabs.TabPages.Add(CreateGridTab("Supplier Aging", dgvSupplierAging))
            tabs.TabPages.Add(CreateProfitTab())

            Return tabs
        End Function

        Private Function CreateOverviewTab() As TabPage
            Dim page As New TabPage("Overview") With {.BackColor = ThemePalette.AppBackground}
            Dim host As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(host)

            Dim table As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2
            }
            table.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 55.0F))
            table.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 45.0F))

            Dim metrics As String() = {
                "Sales Invoice Count",
                "Purchase Bill Count",
                "Sales Units",
                "Purchase Units",
                "Average Sale Bill",
                "Average Purchase Bill",
                "Collections Received",
                "Supplier Payments",
                "Collection Efficiency %",
                "Supplier Payment Coverage %",
                "Inventory SKU Count",
                "Inventory Stock Value At PTR",
                "Outstanding Receivables",
                "Outstanding Payables",
                "Net Cash Movement"
            }

            For rowIndex As Integer = 0 To metrics.Length - 1
                table.RowStyles.Add(New RowStyle(SizeType.Absolute, 34))

                Dim caption As New Label With {
                    .Dock = DockStyle.Fill,
                    .Text = metrics(rowIndex),
                    .ForeColor = ThemePalette.TextPrimary,
                    .Font = New Font("Segoe UI", 10.0F, FontStyle.Regular),
                    .TextAlign = ContentAlignment.MiddleLeft
                }

                Dim valueLabel As New Label With {
                    .Dock = DockStyle.Fill,
                    .Text = "0.00",
                    .ForeColor = ThemePalette.TextPrimary,
                    .Font = New Font("Segoe UI Semibold", 10.5F, FontStyle.Bold),
                    .TextAlign = ContentAlignment.MiddleRight
                }

                overviewLabels(metrics(rowIndex)) = valueLabel
                table.Controls.Add(caption, 0, rowIndex)
                table.Controls.Add(valueLabel, 1, rowIndex)
            Next

            host.Controls.Add(table)
            page.Controls.Add(host)
            Return page
        End Function

        Private Function CreateGridTab(title As String, grid As DataGridView) As TabPage
            Dim page As New TabPage(title) With {.BackColor = ThemePalette.AppBackground}
            Dim host As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(host)
            grid.Dock = DockStyle.Fill
            host.Controls.Add(grid)
            page.Controls.Add(host)
            Return page
        End Function

        Private Function CreateProfitTab() As TabPage
            Dim page As New TabPage("Profit & Loss") With {.BackColor = ThemePalette.AppBackground}
            Dim host As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(host)

            Dim table As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2
            }
            table.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 55.0F))
            table.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 45.0F))

            Dim metrics As String() = {
                "Sales Taxable Amount",
                "Sales Net Amount",
                "Purchase Taxable Amount",
                "Purchase Net Amount",
                "Estimated Cost Of Goods Sold",
                "Estimated Gross Profit",
                "Gross Margin Percentage",
                "Outstanding Receivables",
                "Outstanding Payables"
            }

            For rowIndex As Integer = 0 To metrics.Length - 1
                table.RowStyles.Add(New RowStyle(SizeType.Absolute, 34))

                Dim caption As New Label With {
                    .Dock = DockStyle.Fill,
                    .Text = metrics(rowIndex),
                    .ForeColor = ThemePalette.TextPrimary,
                    .Font = New Font("Segoe UI", 10.0F, FontStyle.Regular),
                    .TextAlign = ContentAlignment.MiddleLeft
                }

                Dim valueLabel As New Label With {
                    .Dock = DockStyle.Fill,
                    .Text = "0.00",
                    .ForeColor = ThemePalette.TextPrimary,
                    .Font = New Font("Segoe UI Semibold", 10.5F, FontStyle.Bold),
                    .TextAlign = ContentAlignment.MiddleRight
                }

                profitLabels(metrics(rowIndex)) = valueLabel
                table.Controls.Add(caption, 0, rowIndex)
                table.Controls.Add(valueLabel, 1, rowIndex)
            Next

            host.Controls.Add(table)
            page.Controls.Add(host)
            Return page
        End Function

        Private Function CreateInputHost(labelText As String, editorControl As Control) As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .Padding = New Padding(0, 0, 12, 10)}

            Dim labelControl As New Label With {
                .Dock = DockStyle.Top,
                .Height = 22,
                .Text = labelText,
                .Font = New Font("Segoe UI Semibold", 9.5F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary
            }

            editorControl.Dock = DockStyle.Fill
            panel.Controls.Add(editorControl)
            panel.Controls.Add(labelControl)
            Return panel
        End Function

        Private Sub ConfigureControls()
            dtpFromDate.Format = DateTimePickerFormat.Custom
            dtpFromDate.CustomFormat = "dd-MMM-yyyy"
            dtpFromDate.Value = New DateTime(DateTime.Today.Year, DateTime.Today.Month, 1)

            dtpToDate.Format = DateTimePickerFormat.Custom
            dtpToDate.CustomFormat = "dd-MMM-yyyy"
            dtpToDate.Value = DateTime.Today

            ConfigureGrid(dgvSales)
            ConfigureGrid(dgvPurchases)
            ConfigureGrid(dgvGst)
            ConfigureGrid(dgvStock)
            ConfigureGrid(dgvOutstanding)
            ConfigureGrid(dgvCustomerAging)
            ConfigureGrid(dgvSupplierAging)
        End Sub

        Private Sub ConfigureGrid(grid As DataGridView)
            UiStyler.StyleDataGrid(grid)
            grid.AutoGenerateColumns = True
            grid.ReadOnly = True
        End Sub

        Private Sub WireEvents()
            AddHandler Load, AddressOf FrmReports_Load
            AddHandler btnRefresh.Click, AddressOf btnRefresh_Click
            AddHandler btnExport.Click, AddressOf btnExport_Click
            AddHandler btnOpenFolder.Click, AddressOf btnOpenFolder_Click
            AddHandler btnClose.Click, AddressOf btnClose_Click
        End Sub

        Private Async Sub FrmReports_Load(sender As Object, e As EventArgs)
            Await LoadReportsAsync()
        End Sub

        Private Async Sub btnRefresh_Click(sender As Object, e As EventArgs)
            Await LoadReportsAsync()
        End Sub

        Private Async Function LoadReportsAsync() As Task
            If dtpFromDate.Value.Date > dtpToDate.Value.Date Then
                ShowStatus("From date cannot be later than To date.", True)
                Return
            End If

            SetBusy(True, "Loading reports...")

            Try
                Dim fromDate As DateTime = dtpFromDate.Value.Date
                Dim toDate As DateTime = dtpToDate.Value.Date

                Dim salesTask As Task(Of List(Of SalesReportRow)) = _reportService.GetSalesReportAsync(fromDate, toDate)
                Dim purchaseTask As Task(Of List(Of PurchaseReportRow)) = _reportService.GetPurchaseReportAsync(fromDate, toDate)
                Dim gstTask As Task(Of List(Of GstReportRow)) = _reportService.GetGstReportAsync(fromDate, toDate)
                Dim stockTask As Task(Of List(Of StockReportRow)) = _reportService.GetStockReportAsync()
                Dim outstandingTask As Task(Of List(Of CustomerOutstandingReportRow)) = _reportService.GetCustomerOutstandingReportAsync()
                Dim customerAgingTask As Task(Of List(Of AgingReportRow)) = _reportService.GetCustomerAgingReportAsync(toDate)
                Dim supplierAgingTask As Task(Of List(Of AgingReportRow)) = _reportService.GetSupplierAgingReportAsync(toDate)
                Dim overviewTask As Task(Of ReportOverview) = _reportService.GetReportOverviewAsync(fromDate, toDate)
                Dim profitTask As Task(Of ProfitLossReport) = _reportService.GetProfitLossReportAsync(fromDate, toDate)

                Await Task.WhenAll(salesTask, purchaseTask, gstTask, stockTask, outstandingTask, customerAgingTask, supplierAgingTask, overviewTask, profitTask)

                BindOverview(overviewTask.Result)
                BindGrid(dgvSales, salesTask.Result)
                BindGrid(dgvPurchases, purchaseTask.Result)
                BindGrid(dgvGst, gstTask.Result)
                BindGrid(dgvStock, stockTask.Result)
                BindGrid(dgvOutstanding, outstandingTask.Result)
                BindGrid(dgvCustomerAging, customerAgingTask.Result)
                BindGrid(dgvSupplierAging, supplierAgingTask.Result)
                BindProfit(profitTask.Result)

                ShowStatus("Reports loaded successfully.", False)
            Catch ex As Exception
                AppLogger.Error("Report loading failed.", ex)
                ShowStatus("Reports could not be loaded.", True)
            Finally
                SetBusy(False)
            End Try
        End Function

        Private Sub BindGrid(Of T)(grid As DataGridView, rows As List(Of T))
            grid.DataSource = Nothing
            grid.DataSource = rows
            ApplyGridFormatting(grid)
        End Sub

        Private Sub ApplyGridFormatting(grid As DataGridView)
            For Each column As DataGridViewColumn In grid.Columns
                Dim propertyName As String = column.DataPropertyName
                column.HeaderText = ResolveHeaderText(propertyName)
                If propertyName.EndsWith("Date", StringComparison.OrdinalIgnoreCase) Then
                    column.DefaultCellStyle.Format = "dd-MMM-yyyy"
                ElseIf propertyName.IndexOf("Amount", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                       propertyName.IndexOf("Net", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                       propertyName.IndexOf("Value", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                       propertyName.IndexOf("PTR", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                       propertyName.IndexOf("PTS", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                       propertyName.IndexOf("MRP", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                       propertyName.IndexOf("Margin", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                       propertyName.IndexOf("Percentage", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    column.DefaultCellStyle.Format = "N2"
                End If
            Next
        End Sub

        Private Function ResolveHeaderText(propertyName As String) As String
            Select Case propertyName
                Case NameOf(AgingReportRow.PartyName)
                    Return "Party"
                Case NameOf(AgingReportRow.DrugLicenseNumber)
                    Return "Drug License"
                Case NameOf(AgingReportRow.OpenDocumentCount)
                    Return "Open Docs"
                Case NameOf(AgingReportRow.OldestOpenDate)
                    Return "Oldest Open Date"
                Case NameOf(AgingReportRow.AgeInDays)
                    Return "Age (Days)"
                Case NameOf(AgingReportRow.Days0To30Amount)
                    Return "0-30 Days"
                Case NameOf(AgingReportRow.Days31To60Amount)
                    Return "31-60 Days"
                Case NameOf(AgingReportRow.Days61To90Amount)
                    Return "61-90 Days"
                Case NameOf(AgingReportRow.DaysAbove90Amount)
                    Return ">90 Days"
                Case NameOf(AgingReportRow.UnallocatedAmount)
                    Return "Unallocated"
                Case NameOf(CustomerOutstandingReportRow.DrugLicenseNumber)
                    Return "Drug License"
                Case NameOf(CustomerOutstandingReportRow.OutstandingBalance)
                    Return "Outstanding"
                Case Else
                    Return InsertSpaces(propertyName)
            End Select
        End Function

        Private Function InsertSpaces(value As String) As String
            If String.IsNullOrWhiteSpace(value) Then
                Return String.Empty
            End If

            Dim builder As New System.Text.StringBuilder()
            For index As Integer = 0 To value.Length - 1
                Dim currentChar As Char = value(index)
                If index > 0 AndAlso Char.IsUpper(currentChar) AndAlso (Char.IsLower(value(index - 1)) OrElse Char.IsDigit(value(index - 1))) Then
                    builder.Append(" "c)
                End If
                builder.Append(currentChar)
            Next

            Return builder.ToString()
        End Function

        Private Sub BindOverview(overview As ReportOverview)
            _overview = overview
            SetOverviewValue("Sales Invoice Count", overview.SalesInvoiceCount)
            SetOverviewValue("Purchase Bill Count", overview.PurchaseBillCount)
            SetOverviewValue("Sales Units", overview.SalesUnits)
            SetOverviewValue("Purchase Units", overview.PurchaseUnits)
            SetOverviewValue("Average Sale Bill", overview.AverageSaleBillValue)
            SetOverviewValue("Average Purchase Bill", overview.AveragePurchaseBillValue)
            SetOverviewValue("Collections Received", overview.CustomerCollectionsAmount)
            SetOverviewValue("Supplier Payments", overview.SupplierPaymentsAmount)
            SetOverviewValue("Collection Efficiency %", overview.CollectionEfficiencyPercentage, True)
            SetOverviewValue("Supplier Payment Coverage %", overview.SupplierPaymentCoveragePercentage, True)
            SetOverviewValue("Inventory SKU Count", overview.InventorySkuCount)
            SetOverviewValue("Inventory Stock Value At PTR", overview.InventoryStockValueAtPTR)
            SetOverviewValue("Outstanding Receivables", overview.OutstandingReceivables)
            SetOverviewValue("Outstanding Payables", overview.OutstandingPayables)
            SetOverviewValue("Net Cash Movement", overview.NetCashMovement)
        End Sub

        Private Sub SetOverviewValue(metricName As String, value As Integer)
            If Not overviewLabels.ContainsKey(metricName) Then
                Return
            End If

            overviewLabels(metricName).Text = value.ToString("N0")
        End Sub

        Private Sub SetOverviewValue(metricName As String, value As Decimal, Optional isPercent As Boolean = False)
            If Not overviewLabels.ContainsKey(metricName) Then
                Return
            End If

            overviewLabels(metricName).Text = If(isPercent, $"{value:N2}%", value.ToString("N2"))
        End Sub

        Private Sub BindProfit(report As ProfitLossReport)
            _profitLossReport = report
            SetProfitValue("Sales Taxable Amount", report.SalesTaxableAmount)
            SetProfitValue("Sales Net Amount", report.SalesNetAmount)
            SetProfitValue("Purchase Taxable Amount", report.PurchaseTaxableAmount)
            SetProfitValue("Purchase Net Amount", report.PurchaseNetAmount)
            SetProfitValue("Estimated Cost Of Goods Sold", report.EstimatedCostOfGoodsSold)
            SetProfitValue("Estimated Gross Profit", report.EstimatedGrossProfit)
            SetProfitValue("Gross Margin Percentage", report.GrossMarginPercentage, True)
            SetProfitValue("Outstanding Receivables", report.OutstandingReceivables)
            SetProfitValue("Outstanding Payables", report.OutstandingPayables)
        End Sub

        Private Sub SetProfitValue(metricName As String, value As Decimal, Optional isPercent As Boolean = False)
            If Not profitLabels.ContainsKey(metricName) Then
                Return
            End If

            profitLabels(metricName).Text = If(isPercent, $"{value:N2}%", value.ToString("N2"))
        End Sub

        Private Sub btnExport_Click(sender As Object, e As EventArgs)
            Try
                Dim filePath As String
                Select Case tabs.SelectedTab.Text
                    Case "Overview"
                        If _overview Is Nothing Then
                            ShowStatus("Load reports before exporting.", True)
                            Return
                        End If
                        filePath = ReportExportHelper.ExportOverview(_overview)
                    Case "Sales"
                        filePath = ReportExportHelper.ExportGrid("SalesReport", dgvSales)
                    Case "Purchases"
                        filePath = ReportExportHelper.ExportGrid("PurchaseReport", dgvPurchases)
                    Case "GST"
                        filePath = ReportExportHelper.ExportGrid("GstReport", dgvGst)
                    Case "Stock"
                        filePath = ReportExportHelper.ExportGrid("StockReport", dgvStock)
                    Case "Outstanding"
                        filePath = ReportExportHelper.ExportGrid("CustomerOutstandingReport", dgvOutstanding)
                    Case "Customer Aging"
                        filePath = ReportExportHelper.ExportGrid("CustomerAgingReport", dgvCustomerAging)
                    Case "Supplier Aging"
                        filePath = ReportExportHelper.ExportGrid("SupplierAgingReport", dgvSupplierAging)
                    Case Else
                        If _profitLossReport Is Nothing Then
                            ShowStatus("Load reports before exporting.", True)
                            Return
                        End If
                        filePath = ReportExportHelper.ExportProfitLoss(_profitLossReport)
                End Select

                ShowStatus($"Report exported to {Path.GetFileName(filePath)}.", False)
            Catch ex As Exception
                AppLogger.Error("Report export failed.", ex)
                ShowStatus("The current report could not be exported.", True)
            End Try
        End Sub

        Private Sub btnOpenFolder_Click(sender As Object, e As EventArgs)
            Try
                Process.Start(New ProcessStartInfo With {
                    .FileName = AppPaths.ReportsDirectory,
                    .UseShellExecute = True
                })
                ShowStatus("Opened reports folder.", False)
            Catch ex As Exception
                AppLogger.Error("Reports folder could not be opened.", ex)
                ShowStatus("The reports folder could not be opened.", True)
            End Try
        End Sub

        Private Sub btnClose_Click(sender As Object, e As EventArgs)
            Close()
        End Sub

        Private Sub SetBusy(isBusy As Boolean, Optional message As String = "")
            _isBusy = isBusy
            dtpFromDate.Enabled = Not isBusy
            dtpToDate.Enabled = Not isBusy
            btnRefresh.Enabled = Not isBusy
            btnExport.Enabled = Not isBusy
            btnOpenFolder.Enabled = Not isBusy
            btnClose.Enabled = Not isBusy
            tabs.Enabled = Not isBusy

            If isBusy Then
                lblStatus.ForeColor = ThemePalette.TextMuted
                lblStatus.Text = message
            End If
        End Sub

        Private Sub ShowStatus(message As String, isError As Boolean)
            lblStatus.ForeColor = If(isError, ThemePalette.DangerRed, ThemePalette.AccentGreen)
            lblStatus.Text = message
        End Sub

        Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
            Select Case keyData
                Case Keys.F5
                    If Not _isBusy Then
                        Dim loadTask As Task = LoadReportsAsync()
                    End If
                    Return True
                Case Keys.Control Or Keys.E
                    If btnExport.Enabled Then
                        btnExport.PerformClick()
                    End If
                    Return True
                Case Keys.Escape
                    Close()
                    Return True
            End Select

            Return MyBase.ProcessCmdKey(msg, keyData)
        End Function

    End Class

End Namespace
