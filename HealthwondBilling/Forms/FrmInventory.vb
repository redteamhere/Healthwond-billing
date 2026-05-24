Imports HealthwondBilling.Controls
Imports HealthwondBilling.Models
Imports HealthwondBilling.Services
Imports HealthwondBilling.Utilities
Imports System.Diagnostics
Imports System.IO

Namespace Forms

    Public Class FrmInventory
        Inherits Form

        Private ReadOnly _inventoryService As InventoryService

        Private ReadOnly txtSearch As New TextBox()
        Private ReadOnly nudExpiryWindow As New NumericUpDown()
        Private ReadOnly dtpLedgerFromDate As New DateTimePicker()
        Private ReadOnly dtpLedgerToDate As New DateTimePicker()
        Private ReadOnly btnRefresh As New Button()
        Private ReadOnly btnExport As New Button()
        Private ReadOnly btnOpenFolder As New Button()
        Private ReadOnly btnClose As New Button()
        Private ReadOnly tabs As New TabControl()
        Private ReadOnly lblStatus As New Label()

        Private ReadOnly dgvCurrentStock As New DataGridView()
        Private ReadOnly dgvBatchStock As New DataGridView()
        Private ReadOnly dgvExpiryStock As New DataGridView()
        Private ReadOnly dgvLowStock As New DataGridView()
        Private ReadOnly dgvLedger As New DataGridView()

        Private totalProductsCard As StatCardControl
        Private batchCountCard As StatCardControl
        Private totalUnitsCard As StatCardControl
        Private expiringSoonCard As StatCardControl
        Private lowStockCard As StatCardControl

        Private _isBusy As Boolean

        Public Sub New(inventoryService As InventoryService)
            _inventoryService = inventoryService

            Text = "Healthwond Billing System - Inventory"
            StartPosition = FormStartPosition.CenterParent
            WindowState = FormWindowState.Maximized
            MinimumSize = New Size(1400, 860)
            BackColor = ThemePalette.AppBackground
            Font = New Font("Segoe UI", 10.0F, FontStyle.Regular)
            KeyPreview = True

            BuildLayout()
            ConfigureControls()
            ConfigureGrids()
            WireEvents()
        End Sub

        Private Sub BuildLayout()
            Dim root As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 5,
                .Padding = New Padding(22),
                .BackColor = ThemePalette.AppBackground
            }
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 78))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 154))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 92))
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 36))

            root.Controls.Add(BuildHeaderPanel(), 0, 0)
            root.Controls.Add(BuildSummaryPanel(), 0, 1)
            root.Controls.Add(BuildFilterPanel(), 0, 2)
            root.Controls.Add(BuildTabsPanel(), 0, 3)

            lblStatus.Dock = DockStyle.Fill
            lblStatus.TextAlign = ContentAlignment.MiddleLeft
            lblStatus.Font = New Font("Segoe UI Semibold", 9.5F, FontStyle.Bold)
            root.Controls.Add(lblStatus, 0, 4)

            Controls.Add(root)
        End Sub

        Private Function BuildHeaderPanel() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = ThemePalette.AppBackground}

            Dim title As New Label With {
                .Dock = DockStyle.Top,
                .Height = 40,
                .Text = "Inventory",
                .Font = New Font("Segoe UI Semibold", 24.0F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary
            }

            Dim subtitle As New Label With {
                .Dock = DockStyle.Fill,
                .Text = "Track current stock, batch positions, expiry risk, low-stock alerts, and stock-ledger history from the live SQLite data set.",
                .Font = New Font("Segoe UI", 10.5F, FontStyle.Regular),
                .ForeColor = ThemePalette.TextMuted
            }

            panel.Controls.Add(subtitle)
            panel.Controls.Add(title)
            Return panel
        End Function

        Private Function BuildSummaryPanel() As Control
            Dim host As New Panel With {.Dock = DockStyle.Fill, .BackColor = ThemePalette.AppBackground}

            Dim flow As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .WrapContents = True,
                .FlowDirection = FlowDirection.LeftToRight,
                .BackColor = Color.Transparent
            }

            totalProductsCard = CreateSummaryCard("Products", ThemePalette.AccentBlue, "0", "Distinct saleable product lines")
            batchCountCard = CreateSummaryCard("Batches", ThemePalette.BrandBlue, "0", "Tracked inventory batches")
            totalUnitsCard = CreateSummaryCard("Total Units", ThemePalette.AccentGreen, "0", "Available units across all stock")
            expiringSoonCard = CreateSummaryCard("Expiry Risk", ThemePalette.WarningAmber, "0", "Batches within the chosen expiry window")
            lowStockCard = CreateSummaryCard("Low Stock", ThemePalette.DangerRed, "0", "Batches at or below reorder threshold")

            flow.Controls.Add(totalProductsCard)
            flow.Controls.Add(batchCountCard)
            flow.Controls.Add(totalUnitsCard)
            flow.Controls.Add(expiringSoonCard)
            flow.Controls.Add(lowStockCard)

            host.Controls.Add(flow)
            Return host
        End Function

        Private Function CreateSummaryCard(title As String, accentColor As Color, valueText As String, subtitle As String) As StatCardControl
            Return New StatCardControl With {
                .CardTitle = title,
                .AccentColor = accentColor,
                .ValueText = valueText,
                .SubtitleText = subtitle
            }
        End Function

        Private Function BuildFilterPanel() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)

            Dim layout As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 8
            }
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 300))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 130))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 170))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 170))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 110))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 110))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 130))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))

            layout.Controls.Add(CreateInputHost("Search", txtSearch), 0, 0)
            layout.Controls.Add(CreateInputHost("Expiry Window (Days)", nudExpiryWindow), 1, 0)
            layout.Controls.Add(CreateInputHost("Ledger From", dtpLedgerFromDate), 2, 0)
            layout.Controls.Add(CreateInputHost("Ledger To", dtpLedgerToDate), 3, 0)

            UiStyler.StylePrimaryButton(btnRefresh)
            btnRefresh.Text = "Refresh"
            btnRefresh.Width = 100
            btnRefresh.Margin = New Padding(0, 24, 10, 0)

            UiStyler.StyleSecondaryButton(btnExport)
            btnExport.Text = "Export"
            btnExport.Width = 100
            btnExport.Margin = New Padding(0, 24, 10, 0)

            UiStyler.StyleSecondaryButton(btnOpenFolder)
            btnOpenFolder.Text = "Open Folder"
            btnOpenFolder.Width = 118
            btnOpenFolder.Margin = New Padding(0, 24, 10, 0)

            UiStyler.StyleSecondaryButton(btnClose)
            btnClose.Text = "Close"
            btnClose.Width = 90
            btnClose.Margin = New Padding(0, 24, 0, 0)

            layout.Controls.Add(btnRefresh, 4, 0)
            layout.Controls.Add(btnExport, 5, 0)
            layout.Controls.Add(btnOpenFolder, 6, 0)
            layout.Controls.Add(btnClose, 7, 0)

            panel.Controls.Add(layout)
            Return panel
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

        Private Function BuildTabsPanel() As Control
            tabs.Dock = DockStyle.Fill
            tabs.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)

            tabs.TabPages.Add(CreateGridTab("Current Stock", dgvCurrentStock))
            tabs.TabPages.Add(CreateGridTab("Batch Stock", dgvBatchStock))
            tabs.TabPages.Add(CreateGridTab("Expiry Stock", dgvExpiryStock))
            tabs.TabPages.Add(CreateGridTab("Low Stock", dgvLowStock))
            tabs.TabPages.Add(CreateGridTab("Stock Ledger", dgvLedger))

            Return tabs
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

        Private Sub ConfigureControls()
            txtSearch.BorderStyle = BorderStyle.FixedSingle
            UiStyler.StyleInput(txtSearch)

            nudExpiryWindow.Minimum = 0D
            nudExpiryWindow.Maximum = 365D
            nudExpiryWindow.Value = 60D
            nudExpiryWindow.Font = New Font("Segoe UI", 10.5F, FontStyle.Regular)
            nudExpiryWindow.ThousandsSeparator = False

            dtpLedgerFromDate.Format = DateTimePickerFormat.Custom
            dtpLedgerFromDate.CustomFormat = "dd-MMM-yyyy"
            dtpLedgerFromDate.Value = DateTime.Today.AddDays(-30)

            dtpLedgerToDate.Format = DateTimePickerFormat.Custom
            dtpLedgerToDate.CustomFormat = "dd-MMM-yyyy"
            dtpLedgerToDate.Value = DateTime.Today
        End Sub

        Private Sub ConfigureGrids()
            ConfigureGrid(dgvCurrentStock)
            ConfigureCurrentStockColumns()

            ConfigureGrid(dgvBatchStock)
            ConfigureBatchStockColumns()

            ConfigureGrid(dgvExpiryStock)
            ConfigureExpiryColumns()

            ConfigureGrid(dgvLowStock)
            ConfigureLowStockColumns()

            ConfigureGrid(dgvLedger)
            ConfigureLedgerColumns()
        End Sub

        Private Sub ConfigureGrid(grid As DataGridView)
            UiStyler.StyleDataGrid(grid)
            grid.AutoGenerateColumns = False
            grid.ReadOnly = True
        End Sub

        Private Sub ConfigureCurrentStockColumns()
            dgvCurrentStock.Columns.Add(CreateTextColumn("ProductName", "Product", 150))
            dgvCurrentStock.Columns.Add(CreateTextColumn("CompanyName", "Company", 130))
            dgvCurrentStock.Columns.Add(CreateTextColumn("Packing", "Packing", 90))
            dgvCurrentStock.Columns.Add(CreateTextColumn("Composition", "Composition", 180))
            dgvCurrentStock.Columns.Add(CreateIntegerColumn("BatchCount", "Batches", 70))
            dgvCurrentStock.Columns.Add(CreateIntegerColumn("TotalStock", "Total Stock", 85))
            dgvCurrentStock.Columns.Add(CreateDateColumn("EarliestExpiryDate", "Earliest Expiry", 90))
            dgvCurrentStock.Columns.Add(CreateDateColumn("LatestExpiryDate", "Latest Expiry", 90))
            dgvCurrentStock.Columns.Add(CreateDecimalColumn("StockValueAtPTR", "Value @ PTR", 90))
            dgvCurrentStock.Columns.Add(CreateDecimalColumn("StockValueAtPTS", "Value @ PTS", 90))
            dgvCurrentStock.Columns.Add(CreateTextColumn("StockStatus", "Status", 90))
        End Sub

        Private Sub ConfigureBatchStockColumns()
            dgvBatchStock.Columns.Add(CreateTextColumn("ProductName", "Product", 150))
            dgvBatchStock.Columns.Add(CreateTextColumn("BatchNumber", "Batch", 90))
            dgvBatchStock.Columns.Add(CreateDateColumn("ExpiryDate", "Expiry", 85))
            dgvBatchStock.Columns.Add(CreateTextColumn("CompanyName", "Company", 120))
            dgvBatchStock.Columns.Add(CreateTextColumn("Packing", "Packing", 80))
            dgvBatchStock.Columns.Add(CreateIntegerColumn("CurrentStock", "Stock", 70))
            dgvBatchStock.Columns.Add(CreateDecimalColumn("GstPercentage", "GST %", 70))
            dgvBatchStock.Columns.Add(CreateDecimalColumn("MRP", "MRP", 75))
            dgvBatchStock.Columns.Add(CreateDecimalColumn("PTR", "PTR", 75))
            dgvBatchStock.Columns.Add(CreateDecimalColumn("PTS", "PTS", 75))
            dgvBatchStock.Columns.Add(CreateTextColumn("Barcode", "Barcode", 105))
            dgvBatchStock.Columns.Add(CreateTextColumn("StockStatus", "Status", 90))
        End Sub

        Private Sub ConfigureExpiryColumns()
            dgvExpiryStock.Columns.Add(CreateTextColumn("ProductName", "Product", 155))
            dgvExpiryStock.Columns.Add(CreateTextColumn("BatchNumber", "Batch", 95))
            dgvExpiryStock.Columns.Add(CreateDateColumn("ExpiryDate", "Expiry", 90))
            dgvExpiryStock.Columns.Add(CreateIntegerColumn("DaysToExpiry", "Days Left", 80))
            dgvExpiryStock.Columns.Add(CreateTextColumn("CompanyName", "Company", 130))
            dgvExpiryStock.Columns.Add(CreateIntegerColumn("CurrentStock", "Stock", 75))
            dgvExpiryStock.Columns.Add(CreateDecimalColumn("MRP", "MRP", 80))
            dgvExpiryStock.Columns.Add(CreateDecimalColumn("PTR", "PTR", 80))
            dgvExpiryStock.Columns.Add(CreateTextColumn("StockStatus", "Status", 95))
        End Sub

        Private Sub ConfigureLowStockColumns()
            dgvLowStock.Columns.Add(CreateTextColumn("ProductName", "Product", 155))
            dgvLowStock.Columns.Add(CreateTextColumn("BatchNumber", "Batch", 95))
            dgvLowStock.Columns.Add(CreateTextColumn("CompanyName", "Company", 135))
            dgvLowStock.Columns.Add(CreateIntegerColumn("CurrentStock", "Current", 75))
            dgvLowStock.Columns.Add(CreateIntegerColumn("ReorderThreshold", "Threshold", 80))
            dgvLowStock.Columns.Add(CreateIntegerColumn("ShortageUnits", "Shortage", 75))
            dgvLowStock.Columns.Add(CreateDateColumn("ExpiryDate", "Expiry", 90))
            dgvLowStock.Columns.Add(CreateTextColumn("StockStatus", "Status", 95))
        End Sub

        Private Sub ConfigureLedgerColumns()
            Dim dateColumn As DataGridViewColumn = CreateDateTimeColumn("TransactionDate", "Date/Time", 120)
            dgvLedger.Columns.Add(dateColumn)
            dgvLedger.Columns.Add(CreateTextColumn("ProductName", "Product", 145))
            dgvLedger.Columns.Add(CreateTextColumn("BatchNumber", "Batch", 85))
            dgvLedger.Columns.Add(CreateTextColumn("TransactionType", "Type", 85))
            dgvLedger.Columns.Add(CreateTextColumn("ReferenceType", "Ref Type", 90))
            dgvLedger.Columns.Add(CreateIntegerColumn("ReferenceId", "Ref Id", 70))
            dgvLedger.Columns.Add(CreateIntegerColumn("QuantityIn", "Qty In", 70))
            dgvLedger.Columns.Add(CreateIntegerColumn("QuantityOut", "Qty Out", 70))
            dgvLedger.Columns.Add(CreateIntegerColumn("BalanceQuantity", "Balance", 75))
            dgvLedger.Columns.Add(CreateDecimalColumn("UnitCost", "Unit Cost", 80))
            dgvLedger.Columns.Add(CreateTextColumn("Remarks", "Remarks", 170))
        End Sub

        Private Function CreateTextColumn(propertyName As String, headerText As String, fillWeight As Single) As DataGridViewColumn
            Return New DataGridViewTextBoxColumn With {
                .DataPropertyName = propertyName,
                .HeaderText = headerText,
                .FillWeight = fillWeight,
                .Name = propertyName
            }
        End Function

        Private Function CreateIntegerColumn(propertyName As String, headerText As String, fillWeight As Single) As DataGridViewColumn
            Return New DataGridViewTextBoxColumn With {
                .DataPropertyName = propertyName,
                .HeaderText = headerText,
                .FillWeight = fillWeight,
                .Name = propertyName,
                .DefaultCellStyle = New DataGridViewCellStyle With {.Format = "N0"}
            }
        End Function

        Private Function CreateDecimalColumn(propertyName As String, headerText As String, fillWeight As Single) As DataGridViewColumn
            Return New DataGridViewTextBoxColumn With {
                .DataPropertyName = propertyName,
                .HeaderText = headerText,
                .FillWeight = fillWeight,
                .Name = propertyName,
                .DefaultCellStyle = New DataGridViewCellStyle With {.Format = "N2"}
            }
        End Function

        Private Function CreateDateColumn(propertyName As String, headerText As String, fillWeight As Single) As DataGridViewColumn
            Return New DataGridViewTextBoxColumn With {
                .DataPropertyName = propertyName,
                .HeaderText = headerText,
                .FillWeight = fillWeight,
                .Name = propertyName,
                .DefaultCellStyle = New DataGridViewCellStyle With {.Format = "dd-MMM-yyyy"}
            }
        End Function

        Private Function CreateDateTimeColumn(propertyName As String, headerText As String, fillWeight As Single) As DataGridViewColumn
            Return New DataGridViewTextBoxColumn With {
                .DataPropertyName = propertyName,
                .HeaderText = headerText,
                .FillWeight = fillWeight,
                .Name = propertyName,
                .DefaultCellStyle = New DataGridViewCellStyle With {.Format = "dd-MMM-yyyy hh:mm tt"}
            }
        End Function

        Private Sub WireEvents()
            AddHandler Load, AddressOf FrmInventory_Load
            AddHandler btnRefresh.Click, AddressOf btnRefresh_Click
            AddHandler btnExport.Click, AddressOf btnExport_Click
            AddHandler btnOpenFolder.Click, AddressOf btnOpenFolder_Click
            AddHandler btnClose.Click, AddressOf btnClose_Click
            AddHandler txtSearch.KeyDown, AddressOf txtSearch_KeyDown
        End Sub

        Private Async Sub FrmInventory_Load(sender As Object, e As EventArgs)
            Await LoadInventoryAsync()
        End Sub

        Private Async Sub btnRefresh_Click(sender As Object, e As EventArgs)
            Await LoadInventoryAsync()
        End Sub

        Private Async Function LoadInventoryAsync() As Task
            If dtpLedgerFromDate.Value.Date > dtpLedgerToDate.Value.Date Then
                ShowStatus("Ledger From date cannot be later than Ledger To date.", True)
                Return
            End If

            SetBusy(True, "Loading inventory views...")

            Try
                Dim searchTerm As String = txtSearch.Text.Trim()
                Dim expiryWindowInDays As Integer = Decimal.ToInt32(nudExpiryWindow.Value)
                Dim fromDate As DateTime = dtpLedgerFromDate.Value.Date
                Dim toDate As DateTime = dtpLedgerToDate.Value.Date

                Dim summaryTask As Task(Of InventorySummary) = _inventoryService.GetInventorySummaryAsync(expiryWindowInDays)
                Dim currentStockTask As Task(Of List(Of InventoryCurrentStockRow)) = _inventoryService.GetCurrentStockAsync(searchTerm)
                Dim batchStockTask As Task(Of List(Of InventoryBatchStockRow)) = _inventoryService.GetBatchStockAsync(searchTerm)
                Dim expiryTask As Task(Of List(Of InventoryExpiryRow)) = _inventoryService.GetExpiryStockAsync(searchTerm, expiryWindowInDays)
                Dim lowStockTask As Task(Of List(Of InventoryLowStockRow)) = _inventoryService.GetLowStockAsync(searchTerm)
                Dim ledgerTask As Task(Of List(Of InventoryLedgerRow)) = _inventoryService.GetStockLedgerAsync(searchTerm, fromDate, toDate)

                Await Task.WhenAll(New Task() {summaryTask, currentStockTask, batchStockTask, expiryTask, lowStockTask, ledgerTask})

                BindSummary(summaryTask.Result, expiryWindowInDays)
                BindGrid(dgvCurrentStock, currentStockTask.Result)
                BindGrid(dgvBatchStock, batchStockTask.Result)
                BindGrid(dgvExpiryStock, expiryTask.Result)
                BindGrid(dgvLowStock, lowStockTask.Result)
                BindGrid(dgvLedger, ledgerTask.Result)

                ShowStatus("Inventory views loaded successfully.", False)
            Catch ex As Exception
                AppLogger.Error("Inventory loading failed.", ex)
                ShowStatus("Inventory views could not be loaded.", True)
            Finally
                SetBusy(False)
            End Try
        End Function

        Private Sub BindSummary(summary As InventorySummary, expiryWindowInDays As Integer)
            totalProductsCard.ValueText = summary.DistinctProducts.ToString("N0")
            batchCountCard.ValueText = summary.BatchCount.ToString("N0")
            totalUnitsCard.ValueText = summary.TotalUnits.ToString("N0")
            expiringSoonCard.ValueText = summary.ExpiringSoonCount.ToString("N0")
            expiringSoonCard.SubtitleText = $"Batches within {expiryWindowInDays} days"
            lowStockCard.ValueText = summary.LowStockCount.ToString("N0")
        End Sub

        Private Sub BindGrid(Of T)(grid As DataGridView, rows As List(Of T))
            grid.DataSource = Nothing
            grid.DataSource = rows
            ApplyStatusStyling(grid)
        End Sub

        Private Sub ApplyStatusStyling(grid As DataGridView)
            If Not grid.Columns.Contains("StockStatus") Then
                Return
            End If

            For Each row As DataGridViewRow In grid.Rows
                If row.IsNewRow Then
                    Continue For
                End If

                Dim statusText As String = Convert.ToString(row.Cells("StockStatus").Value)
                Dim statusCell As DataGridViewCell = row.Cells("StockStatus")

                Select Case statusText
                    Case "Expired", "Out of Stock"
                        statusCell.Style.ForeColor = ThemePalette.DangerRed
                    Case "Low Stock"
                        statusCell.Style.ForeColor = ThemePalette.WarningAmber
                    Case "Expiring Soon"
                        statusCell.Style.ForeColor = ThemePalette.WarningAmber
                    Case Else
                        statusCell.Style.ForeColor = ThemePalette.AccentGreen
                End Select
            Next
        End Sub

        Private Sub txtSearch_KeyDown(sender As Object, e As KeyEventArgs)
            If e.KeyCode = Keys.Enter AndAlso Not _isBusy Then
                e.SuppressKeyPress = True
                Dim loadTask As Task = LoadInventoryAsync()
            End If
        End Sub

        Private Sub btnExport_Click(sender As Object, e As EventArgs)
            Try
                Dim filePath As String

                Select Case tabs.SelectedTab.Text
                    Case "Current Stock"
                        filePath = ReportExportHelper.ExportGrid("InventoryCurrentStock", dgvCurrentStock)
                    Case "Batch Stock"
                        filePath = ReportExportHelper.ExportGrid("InventoryBatchStock", dgvBatchStock)
                    Case "Expiry Stock"
                        filePath = ReportExportHelper.ExportGrid("InventoryExpiryStock", dgvExpiryStock)
                    Case "Low Stock"
                        filePath = ReportExportHelper.ExportGrid("InventoryLowStock", dgvLowStock)
                    Case Else
                        filePath = ReportExportHelper.ExportGrid("InventoryStockLedger", dgvLedger)
                End Select

                ShowStatus($"Inventory export created: {Path.GetFileName(filePath)}.", False)
            Catch ex As Exception
                AppLogger.Error("Inventory export failed.", ex)
                ShowStatus("The current inventory view could not be exported.", True)
            End Try
        End Sub

        Private Sub btnOpenFolder_Click(sender As Object, e As EventArgs)
            Try
                Process.Start(New ProcessStartInfo With {
                    .FileName = AppPaths.ReportsDirectory,
                    .UseShellExecute = True
                })
                ShowStatus("Opened the reports folder for exported inventory files.", False)
            Catch ex As Exception
                AppLogger.Error("Inventory export folder could not be opened.", ex)
                ShowStatus("The export folder could not be opened.", True)
            End Try
        End Sub

        Private Sub btnClose_Click(sender As Object, e As EventArgs)
            Close()
        End Sub

        Private Sub SetBusy(isBusy As Boolean, Optional message As String = "")
            _isBusy = isBusy
            txtSearch.Enabled = Not isBusy
            nudExpiryWindow.Enabled = Not isBusy
            dtpLedgerFromDate.Enabled = Not isBusy
            dtpLedgerToDate.Enabled = Not isBusy
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
                        Dim loadTask As Task = LoadInventoryAsync()
                    End If
                    Return True
                Case Keys.Control Or Keys.E
                    If btnExport.Enabled Then
                        btnExport.PerformClick()
                    End If
                    Return True
                Case Keys.Control Or Keys.F
                    txtSearch.Focus()
                    txtSearch.SelectAll()
                    Return True
                Case Keys.Escape
                    Close()
                    Return True
            End Select

            Return MyBase.ProcessCmdKey(msg, keyData)
        End Function

    End Class

End Namespace
