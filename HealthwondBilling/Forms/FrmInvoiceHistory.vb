Imports HealthwondBilling.Models
Imports HealthwondBilling.Services
Imports HealthwondBilling.Utilities
Imports System.ComponentModel
Imports System.IO

Namespace Forms

    Public Class FrmInvoiceHistory
        Inherits Form

        Private ReadOnly _billingService As BillingService
        Private ReadOnly _invoiceExportService As InvoiceExportService
        Private ReadOnly _invoiceRows As New BindingList(Of InvoiceHistoryRow)()
        Private ReadOnly _invoiceItemRows As New BindingList(Of InvoiceDocumentItem)()

        Private ReadOnly txtSearch As New TextBox()
        Private ReadOnly dtpFromDate As New DateTimePicker()
        Private ReadOnly dtpToDate As New DateTimePicker()
        Private ReadOnly btnRefresh As New Button()
        Private ReadOnly btnEdit As New Button()
        Private ReadOnly btnExport As New Button()
        Private ReadOnly btnPrintPreview As New Button()
        Private ReadOnly btnPrint As New Button()
        Private ReadOnly btnOpenExcel As New Button()
        Private ReadOnly btnOpenPdf As New Button()
        Private ReadOnly btnOpenFolder As New Button()
        Private ReadOnly btnClose As New Button()
        Private ReadOnly dgvInvoices As New DataGridView()
        Private ReadOnly dgvInvoiceItems As New DataGridView()
        Private ReadOnly lblDetailHeader As New Label()
        Private ReadOnly lblDocumentStatus As New Label()
        Private ReadOnly lblStatus As New Label()

        Private _isBusy As Boolean

        Public Sub New(billingService As BillingService, invoiceExportService As InvoiceExportService)
            _billingService = billingService
            _invoiceExportService = invoiceExportService

            Text = "Healthwond Billing System - Invoice History"
            StartPosition = FormStartPosition.CenterParent
            WindowState = FormWindowState.Maximized
            MinimumSize = New Size(1440, 900)
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
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 88))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 68))
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 36))

            root.Controls.Add(BuildHeaderPanel(), 0, 0)
            root.Controls.Add(BuildFilterPanel(), 0, 1)
            root.Controls.Add(BuildActionPanel(), 0, 2)
            root.Controls.Add(BuildContentPanel(), 0, 3)

            lblStatus.Dock = DockStyle.Fill
            lblStatus.Font = New Font("Segoe UI Semibold", 9.5F, FontStyle.Bold)
            lblStatus.TextAlign = ContentAlignment.MiddleLeft
            root.Controls.Add(lblStatus, 0, 4)

            Controls.Add(root)
        End Sub

        Private Function BuildHeaderPanel() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = ThemePalette.AppBackground}

            Dim title As New Label With {
                .Dock = DockStyle.Top,
                .Height = 40,
                .Text = "Invoice History",
                .Font = New Font("Segoe UI Semibold", 24.0F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary
            }

            Dim subtitle As New Label With {
                .Dock = DockStyle.Fill,
                .Text = "Search saved invoices, reopen them for edit, regenerate Excel/PDF outputs, preview prints, and manage the generated files.",
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
                .ColumnCount = 4
            }
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 45.0F))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 180))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 180))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 55.0F))

            layout.Controls.Add(CreateInputHost("Search", txtSearch), 0, 0)
            layout.Controls.Add(CreateInputHost("From Date", dtpFromDate), 1, 0)
            layout.Controls.Add(CreateInputHost("To Date", dtpToDate), 2, 0)
            layout.Controls.Add(CreateHelpLabel(), 3, 0)

            panel.Controls.Add(layout)
            Return panel
        End Function

        Private Function CreateHelpLabel() As Control
            Return New Label With {
                .Dock = DockStyle.Fill,
                .Text = "Search by invoice number, customer, payment mode, or notes.",
                .ForeColor = ThemePalette.TextMuted,
                .Font = New Font("Segoe UI", 9.75F, FontStyle.Italic),
                .TextAlign = ContentAlignment.MiddleLeft
            }
        End Function

        Private Function BuildActionPanel() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)

            Dim flow As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.LeftToRight,
                .WrapContents = False,
                .BackColor = Color.Transparent,
                .Padding = New Padding(0, 6, 0, 0)
            }

            ConfigureActionButton(btnRefresh, "Refresh", AddressOf btnRefresh_Click, True, 100)
            ConfigureActionButton(btnEdit, "Edit Invoice", AddressOf btnEdit_Click, False, 118)
            ConfigureActionButton(btnExport, "Export Excel/PDF", AddressOf btnExport_Click, False, 142)
            ConfigureActionButton(btnPrintPreview, "Print Preview", AddressOf btnPrintPreview_Click, False, 122)
            ConfigureActionButton(btnPrint, "Instant Print", AddressOf btnPrint_Click, False, 112)
            ConfigureActionButton(btnOpenExcel, "Open Excel", AddressOf btnOpenExcel_Click, False, 108)
            ConfigureActionButton(btnOpenPdf, "Open PDF", AddressOf btnOpenPdf_Click, False, 100)
            ConfigureActionButton(btnOpenFolder, "Open Folder", AddressOf btnOpenFolder_Click, False, 110)
            ConfigureActionButton(btnClose, "Close", AddressOf btnClose_Click, False, 90)

            flow.Controls.AddRange(New Control() {btnRefresh, btnEdit, btnExport, btnPrintPreview, btnPrint, btnOpenExcel, btnOpenPdf, btnOpenFolder, btnClose})
            panel.Controls.Add(flow)
            Return panel
        End Function

        Private Function BuildContentPanel() As Control
            Dim split As New SplitContainer With {
                .Dock = DockStyle.Fill,
                .Orientation = Orientation.Horizontal,
                .SplitterDistance = 360,
                .BackColor = ThemePalette.AppBackground
            }

            split.Panel1.Controls.Add(BuildInvoiceGridPanel())
            split.Panel2.Controls.Add(BuildDetailPanel())
            Return split
        End Function

        Private Function BuildInvoiceGridPanel() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)
            dgvInvoices.Dock = DockStyle.Fill
            panel.Controls.Add(dgvInvoices)
            Return panel
        End Function

        Private Function BuildDetailPanel() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)

            Dim root As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 3
            }
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 28))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 28))
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            lblDetailHeader.Dock = DockStyle.Fill
            lblDetailHeader.Font = New Font("Segoe UI Semibold", 13.0F, FontStyle.Bold)
            lblDetailHeader.ForeColor = ThemePalette.TextPrimary
            lblDetailHeader.Text = "Invoice details"

            lblDocumentStatus.Dock = DockStyle.Fill
            lblDocumentStatus.Font = New Font("Segoe UI", 9.5F, FontStyle.Regular)
            lblDocumentStatus.ForeColor = ThemePalette.TextMuted
            lblDocumentStatus.Text = "Select an invoice to load line items and file availability."

            dgvInvoiceItems.Dock = DockStyle.Fill

            root.Controls.Add(lblDetailHeader, 0, 0)
            root.Controls.Add(lblDocumentStatus, 0, 1)
            root.Controls.Add(dgvInvoiceItems, 0, 2)

            panel.Controls.Add(root)
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

        Private Sub ConfigureActionButton(button As Button, text As String, handler As EventHandler, isPrimary As Boolean, width As Integer)
            If isPrimary Then
                UiStyler.StylePrimaryButton(button)
            Else
                UiStyler.StyleSecondaryButton(button)
            End If

            button.Text = text
            button.Width = width
            AddHandler button.Click, handler
        End Sub

        Private Sub ConfigureControls()
            txtSearch.BorderStyle = BorderStyle.FixedSingle
            UiStyler.StyleInput(txtSearch)

            dtpFromDate.Format = DateTimePickerFormat.Custom
            dtpFromDate.CustomFormat = "dd-MMM-yyyy"
            dtpFromDate.Value = New DateTime(DateTime.Today.Year, DateTime.Today.Month, 1)

            dtpToDate.Format = DateTimePickerFormat.Custom
            dtpToDate.CustomFormat = "dd-MMM-yyyy"
            dtpToDate.Value = DateTime.Today
        End Sub

        Private Sub ConfigureGrids()
            ConfigureGrid(dgvInvoices)
            ConfigureGrid(dgvInvoiceItems)

            dgvInvoices.AutoGenerateColumns = False
            dgvInvoices.Columns.Add(CreateDateColumn("InvoiceDate", "Date", 85))
            dgvInvoices.Columns.Add(CreateTextColumn("InvoiceNumber", "Invoice Number", 110))
            dgvInvoices.Columns.Add(CreateTextColumn("CustomerName", "Customer", 150))
            dgvInvoices.Columns.Add(CreateTextColumn("PaymentMode", "Payment", 80))
            dgvInvoices.Columns.Add(CreateIntegerColumn("LineCount", "Lines", 60))
            dgvInvoices.Columns.Add(CreateIntegerColumn("TotalUnits", "Units", 65))
            dgvInvoices.Columns.Add(CreateDecimalColumn("NetAmount", "Net Amount", 85))
            dgvInvoices.Columns.Add(CreateDecimalColumn("AmountPaid", "Paid", 80))
            dgvInvoices.Columns.Add(CreateDecimalColumn("BalanceAmount", "Balance", 80))
            dgvInvoices.Columns.Add(CreateDateTimeColumn("UpdatedAt", "Updated", 110))
            dgvInvoices.DataSource = _invoiceRows

            dgvInvoiceItems.AutoGenerateColumns = False
            dgvInvoiceItems.Columns.Add(CreateIntegerColumn("LineNumber", "#", 40))
            dgvInvoiceItems.Columns.Add(CreateTextColumn("ProductName", "Product", 160))
            dgvInvoiceItems.Columns.Add(CreateTextColumn("BatchNumber", "Batch", 80))
            dgvInvoiceItems.Columns.Add(CreateDateColumn("ExpiryDate", "Expiry", 80))
            dgvInvoiceItems.Columns.Add(CreateIntegerColumn("Quantity", "Qty", 60))
            dgvInvoiceItems.Columns.Add(CreateIntegerColumn("FreeQuantity", "Free", 60))
            dgvInvoiceItems.Columns.Add(CreateDecimalColumn("Rate", "Rate", 70))
            dgvInvoiceItems.Columns.Add(CreateDecimalColumn("DiscountPercentage", "Disc %", 70))
            dgvInvoiceItems.Columns.Add(CreateTextColumn("SchemeDescription", "Scheme", 110))
            dgvInvoiceItems.Columns.Add(CreateDecimalColumn("GstPercentage", "GST %", 65))
            dgvInvoiceItems.Columns.Add(CreateDecimalColumn("TaxableAmount", "Taxable", 85))
            dgvInvoiceItems.Columns.Add(CreateDecimalColumn("GstAmount", "GST", 70))
            dgvInvoiceItems.Columns.Add(CreateDecimalColumn("LineTotal", "Total", 80))
            dgvInvoiceItems.DataSource = _invoiceItemRows
        End Sub

        Private Sub ConfigureGrid(grid As DataGridView)
            UiStyler.StyleDataGrid(grid)
            grid.AutoGenerateColumns = False
            grid.ReadOnly = True
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
            AddHandler Load, AddressOf FrmInvoiceHistory_Load
            AddHandler dgvInvoices.SelectionChanged, AddressOf dgvInvoices_SelectionChanged
            AddHandler txtSearch.KeyDown, AddressOf txtSearch_KeyDown
        End Sub

        Private Async Sub FrmInvoiceHistory_Load(sender As Object, e As EventArgs)
            Await LoadInvoicesAsync()
        End Sub

        Private Async Sub btnRefresh_Click(sender As Object, e As EventArgs)
            Await LoadInvoicesAsync(GetSelectedInvoiceId())
        End Sub

        Private Async Function LoadInvoicesAsync(Optional invoiceIdToSelect As Integer = 0) As Task
            Dim shouldLoadDetails As Boolean = False

            If dtpFromDate.Value.Date > dtpToDate.Value.Date Then
                ShowStatus("From date cannot be later than To date.", True)
                Return
            End If

            SetBusy(True, "Loading invoice history...")

            Try
                Dim rows As List(Of InvoiceHistoryRow) = Await _billingService.LoadInvoiceHistoryAsync(dtpFromDate.Value.Date, dtpToDate.Value.Date, txtSearch.Text.Trim())

                _invoiceRows.Clear()
                For Each row As InvoiceHistoryRow In rows
                    _invoiceRows.Add(row)
                Next

                If _invoiceRows.Count = 0 Then
                    _invoiceItemRows.Clear()
                    lblDetailHeader.Text = "Invoice details"
                    lblDocumentStatus.Text = "No invoices matched the current filters."
                    ShowStatus("No invoices matched the current filters.", False)
                Else
                    SelectInvoiceRow(invoiceIdToSelect)
                    shouldLoadDetails = True
                    ShowStatus("Invoice history loaded successfully.", False)
                End If
            Catch ex As Exception
                AppLogger.Error("Invoice history could not be loaded.", ex)
                ShowStatus("Invoice history could not be loaded.", True)
            Finally
                SetBusy(False)
            End Try

            If shouldLoadDetails Then
                Await LoadSelectedInvoiceDetailsAsync()
            End If
        End Function

        Private Async Sub dgvInvoices_SelectionChanged(sender As Object, e As EventArgs)
            If _isBusy Then
                Return
            End If

            Await LoadSelectedInvoiceDetailsAsync()
        End Sub

        Private Async Function LoadSelectedInvoiceDetailsAsync() As Task
            Dim selectedRow As InvoiceHistoryRow = GetSelectedInvoice()
            UpdateActionState()

            If selectedRow Is Nothing Then
                _invoiceItemRows.Clear()
                lblDetailHeader.Text = "Invoice details"
                lblDocumentStatus.Text = "Select an invoice to load line items and file availability."
                Return
            End If

            SetBusy(True, $"Loading details for {selectedRow.InvoiceNumber}...")

            Try
                Dim document As InvoiceDocument = Await _billingService.GetInvoiceDocumentAsync(selectedRow.InvoiceId)
                _invoiceItemRows.Clear()
                For Each item As InvoiceDocumentItem In document.Items
                    _invoiceItemRows.Add(item)
                Next

                lblDetailHeader.Text = $"Invoice {document.InvoiceNumber} | {document.CustomerName} | Net {document.NetAmount:N2}"
                lblDocumentStatus.Text = BuildDocumentStatusText(document.InvoiceNumber)
                ShowStatus($"Loaded invoice {document.InvoiceNumber}.", False)
            Catch ex As Exception
                AppLogger.Error($"Invoice details could not be loaded for Id {selectedRow.InvoiceId}.", ex)
                _invoiceItemRows.Clear()
                lblDetailHeader.Text = "Invoice details"
                lblDocumentStatus.Text = "Invoice details could not be loaded."
                ShowStatus("Invoice details could not be loaded.", True)
            Finally
                SetBusy(False)
            End Try
        End Function

        Private Function BuildDocumentStatusText(invoiceNumber As String) As String
            Dim excelPath As String = _invoiceExportService.GetExcelFilePath(invoiceNumber)
            Dim pdfPath As String = _invoiceExportService.GetPdfFilePath(invoiceNumber)
            Dim excelStatus As String = If(File.Exists(excelPath), $"Excel ready: {Path.GetFileName(excelPath)}", "Excel not generated")
            Dim pdfStatus As String = If(File.Exists(pdfPath), $"PDF ready: {Path.GetFileName(pdfPath)}", "PDF not generated")
            Return $"{excelStatus} | {pdfStatus}"
        End Function

        Private Async Sub btnEdit_Click(sender As Object, e As EventArgs)
            Dim selectedRow As InvoiceHistoryRow = GetSelectedInvoice()
            If selectedRow Is Nothing Then
                ShowStatus("Select an invoice to edit.", True)
                Return
            End If

            Using form As New FrmBilling(_billingService, _invoiceExportService, selectedRow.InvoiceId)
                form.ShowDialog(Me)
            End Using

            Await LoadInvoicesAsync(selectedRow.InvoiceId)
        End Sub

        Private Async Sub btnExport_Click(sender As Object, e As EventArgs)
            Dim selectedRow As InvoiceHistoryRow = GetSelectedInvoice()
            If selectedRow Is Nothing Then
                ShowStatus("Select an invoice to export.", True)
                Return
            End If

            SetBusy(True, $"Generating files for {selectedRow.InvoiceNumber}...")
            Dim result As InvoiceExportResult = Await _invoiceExportService.GenerateInvoiceFilesAsync(selectedRow.InvoiceId)
            SetBusy(False)

            If result.IsSuccess Then
                lblDocumentStatus.Text = BuildDocumentStatusText(selectedRow.InvoiceNumber)
            End If

            ShowStatus(result.Message, Not result.IsSuccess)
        End Sub

        Private Sub btnPrintPreview_Click(sender As Object, e As EventArgs)
            Dim selectedRow As InvoiceHistoryRow = GetSelectedInvoice()
            If selectedRow Is Nothing Then
                ShowStatus("Select an invoice to preview.", True)
                Return
            End If

            Try
                _invoiceExportService.ShowPrintPreview(selectedRow.InvoiceId)
                ShowStatus($"Preview opened for invoice {selectedRow.InvoiceNumber}.", False)
            Catch ex As Exception
                AppLogger.Error($"Invoice preview failed for Id {selectedRow.InvoiceId}.", ex)
                ShowStatus("Print preview could not be opened.", True)
            End Try
        End Sub

        Private Sub btnPrint_Click(sender As Object, e As EventArgs)
            Dim selectedRow As InvoiceHistoryRow = GetSelectedInvoice()
            If selectedRow Is Nothing Then
                ShowStatus("Select an invoice to print.", True)
                Return
            End If

            Try
                _invoiceExportService.PrintInvoice(selectedRow.InvoiceId)
                ShowStatus($"Invoice {selectedRow.InvoiceNumber} sent to the default printer.", False)
            Catch ex As Exception
                AppLogger.Error($"Invoice print failed for Id {selectedRow.InvoiceId}.", ex)
                ShowStatus("Invoice could not be printed.", True)
            End Try
        End Sub

        Private Sub btnOpenExcel_Click(sender As Object, e As EventArgs)
            OpenSelectedInvoiceDocument(False)
        End Sub

        Private Sub btnOpenPdf_Click(sender As Object, e As EventArgs)
            OpenSelectedInvoiceDocument(True)
        End Sub

        Private Sub OpenSelectedInvoiceDocument(openPdf As Boolean)
            Dim selectedRow As InvoiceHistoryRow = GetSelectedInvoice()
            If selectedRow Is Nothing Then
                ShowStatus("Select an invoice first.", True)
                Return
            End If

            Try
                _invoiceExportService.OpenGeneratedInvoiceFile(selectedRow.InvoiceNumber, openPdf)
                ShowStatus($"Opened {If(openPdf, "PDF", "Excel")} for invoice {selectedRow.InvoiceNumber}.", False)
            Catch ex As FileNotFoundException
                ShowStatus($"The {If(openPdf, "PDF", "Excel")} file is not available. Export the invoice first.", True)
            Catch ex As Exception
                AppLogger.Error($"Invoice file open failed for {selectedRow.InvoiceNumber}.", ex)
                ShowStatus("The invoice file could not be opened.", True)
            End Try
        End Sub

        Private Sub btnOpenFolder_Click(sender As Object, e As EventArgs)
            Try
                _invoiceExportService.OpenInvoiceFolder()
                ShowStatus("Opened generated invoices folder.", False)
            Catch ex As Exception
                AppLogger.Error("Generated invoices folder could not be opened.", ex)
                ShowStatus("The invoice folder could not be opened.", True)
            End Try
        End Sub

        Private Sub btnClose_Click(sender As Object, e As EventArgs)
            Close()
        End Sub

        Private Sub SelectInvoiceRow(invoiceIdToSelect As Integer)
            If dgvInvoices.Rows.Count = 0 Then
                Return
            End If

            If invoiceIdToSelect > 0 Then
                For Each row As DataGridViewRow In dgvInvoices.Rows
                    Dim rowValue As InvoiceHistoryRow = TryCast(row.DataBoundItem, InvoiceHistoryRow)
                    If rowValue IsNot Nothing AndAlso rowValue.InvoiceId = invoiceIdToSelect Then
                        row.Selected = True
                        dgvInvoices.CurrentCell = row.Cells(0)
                        Return
                    End If
                Next
            End If

            dgvInvoices.Rows(0).Selected = True
            dgvInvoices.CurrentCell = dgvInvoices.Rows(0).Cells(0)
        End Sub

        Private Function GetSelectedInvoice() As InvoiceHistoryRow
            If dgvInvoices.CurrentRow Is Nothing Then
                Return Nothing
            End If

            Return TryCast(dgvInvoices.CurrentRow.DataBoundItem, InvoiceHistoryRow)
        End Function

        Private Function GetSelectedInvoiceId() As Integer
            Dim selectedRow As InvoiceHistoryRow = GetSelectedInvoice()
            If selectedRow Is Nothing Then
                Return 0
            End If

            Return selectedRow.InvoiceId
        End Function

        Private Sub txtSearch_KeyDown(sender As Object, e As KeyEventArgs)
            If e.KeyCode = Keys.Enter AndAlso Not _isBusy Then
                e.SuppressKeyPress = True
                Dim loadTask As Task = LoadInvoicesAsync(GetSelectedInvoiceId())
            End If
        End Sub

        Private Sub SetBusy(isBusy As Boolean, Optional message As String = "")
            _isBusy = isBusy
            txtSearch.Enabled = Not isBusy
            dtpFromDate.Enabled = Not isBusy
            dtpToDate.Enabled = Not isBusy
            btnRefresh.Enabled = Not isBusy
            dgvInvoices.Enabled = Not isBusy
            dgvInvoiceItems.Enabled = Not isBusy

            If isBusy Then
                lblStatus.ForeColor = ThemePalette.TextMuted
                lblStatus.Text = message
            End If

            UpdateActionState()
        End Sub

        Private Sub UpdateActionState()
            Dim hasSelection As Boolean = GetSelectedInvoice() IsNot Nothing
            btnEdit.Enabled = hasSelection AndAlso Not _isBusy
            btnExport.Enabled = hasSelection AndAlso Not _isBusy
            btnPrintPreview.Enabled = hasSelection AndAlso Not _isBusy
            btnPrint.Enabled = hasSelection AndAlso Not _isBusy
            btnOpenExcel.Enabled = hasSelection AndAlso Not _isBusy
            btnOpenPdf.Enabled = hasSelection AndAlso Not _isBusy
            btnOpenFolder.Enabled = Not _isBusy
            btnClose.Enabled = Not _isBusy
        End Sub

        Private Sub ShowStatus(message As String, isError As Boolean)
            lblStatus.ForeColor = If(isError, ThemePalette.DangerRed, ThemePalette.AccentGreen)
            lblStatus.Text = message
        End Sub

        Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
            Select Case keyData
                Case Keys.F5
                    If Not _isBusy Then
                        Dim loadTask As Task = LoadInvoicesAsync(GetSelectedInvoiceId())
                    End If
                    Return True
                Case Keys.Control Or Keys.F
                    txtSearch.Focus()
                    txtSearch.SelectAll()
                    Return True
                Case Keys.Control Or Keys.E
                    If btnExport.Enabled Then
                        btnExport.PerformClick()
                    End If
                    Return True
                Case Keys.Control Or Keys.P
                    If btnPrint.Enabled Then
                        btnPrint.PerformClick()
                    End If
                    Return True
                Case Keys.Control Or Keys.Shift Or Keys.P
                    If btnPrintPreview.Enabled Then
                        btnPrintPreview.PerformClick()
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
