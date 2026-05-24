Imports HealthwondBilling.Models
Imports HealthwondBilling.Services
Imports HealthwondBilling.Utilities
Imports System.ComponentModel
Imports System.Linq

Namespace Forms

    Public Class FrmBilling
        Inherits Form

        Private ReadOnly _billingService As BillingService
        Private ReadOnly _invoiceExportService As InvoiceExportService
        Private ReadOnly _invoiceItems As New BindingList(Of BillingLineItem)()
        Private _customers As New List(Of CustomerRecord)()
        Private _products As New List(Of ProductRecord)()
        Private _isBusy As Boolean
        Private _lastSavedInvoiceId As Integer
        Private _lastSavedInvoiceNumber As String = String.Empty

        Private ReadOnly txtInvoiceNumber As New TextBox()
        Private ReadOnly dtpInvoiceDate As New DateTimePicker()
        Private ReadOnly cboCustomer As New ComboBox()
        Private ReadOnly cboPaymentMode As New ComboBox()
        Private ReadOnly nudAmountPaid As New NumericUpDown()
        Private ReadOnly txtNotes As New TextBox()

        Private ReadOnly cboProduct As New ComboBox()
        Private ReadOnly nudQuantity As New NumericUpDown()
        Private ReadOnly nudFreeQuantity As New NumericUpDown()
        Private ReadOnly nudDiscount As New NumericUpDown()
        Private ReadOnly txtScheme As New TextBox()
        Private ReadOnly btnAddItem As New Button()
        Private ReadOnly btnRemoveItem As New Button()

        Private ReadOnly dgvItems As New DataGridView()
        Private ReadOnly btnNewInvoice As New Button()
        Private ReadOnly btnSaveInvoice As New Button()
        Private ReadOnly btnExportInvoice As New Button()
        Private ReadOnly btnPrintPreview As New Button()
        Private ReadOnly btnPrintInvoice As New Button()
        Private ReadOnly btnOpenInvoiceFolder As New Button()
        Private ReadOnly btnClose As New Button()
        Private ReadOnly lblStatus As New Label()

        Private ReadOnly lblSubTotalValue As New Label()
        Private ReadOnly lblDiscountValue As New Label()
        Private ReadOnly lblGstValue As New Label()
        Private ReadOnly lblRoundOffValue As New Label()
        Private ReadOnly lblNetAmountValue As New Label()
        Private ReadOnly lblBalanceValue As New Label()
        Private ReadOnly txtTaxSummary As New TextBox()

        Public Sub New(billingService As BillingService, invoiceExportService As InvoiceExportService)
            _billingService = billingService
            _invoiceExportService = invoiceExportService

            Text = "Healthwond Billing System - Billing"
            StartPosition = FormStartPosition.CenterParent
            WindowState = FormWindowState.Maximized
            MinimumSize = New Size(1380, 860)
            BackColor = ThemePalette.AppBackground
            Font = New Font("Segoe UI", 10.0F, FontStyle.Regular)
            KeyPreview = True

            BuildLayout()
            ConfigureEditors()
            ConfigureGrid()
            WireEvents()
        End Sub

        Private Sub BuildLayout()
            Dim root As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 6,
                .Padding = New Padding(22),
                .BackColor = ThemePalette.AppBackground
            }
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 80))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 170))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 122))
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 190))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 36))

            root.Controls.Add(BuildHeaderPanel(), 0, 0)
            root.Controls.Add(BuildInvoicePanel(), 0, 1)
            root.Controls.Add(BuildAddItemPanel(), 0, 2)
            root.Controls.Add(BuildGridPanel(), 0, 3)
            root.Controls.Add(BuildFooterPanel(), 0, 4)

            lblStatus.Dock = DockStyle.Fill
            lblStatus.TextAlign = ContentAlignment.MiddleLeft
            lblStatus.Font = New Font("Segoe UI Semibold", 9.5F, FontStyle.Bold)
            root.Controls.Add(lblStatus, 0, 5)

            Controls.Add(root)
        End Sub

        Private Function BuildHeaderPanel() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = ThemePalette.AppBackground}

            Dim title As New Label With {
                .Dock = DockStyle.Top,
                .Height = 40,
                .Text = "Billing",
                .Font = New Font("Segoe UI Semibold", 24.0F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary
            }

            Dim subtitle As New Label With {
                .Dock = DockStyle.Fill,
                .Text = "Create GST invoices with customer selection, product lines, discounts, free quantity schemes, live tax totals, and stock deduction on save.",
                .Font = New Font("Segoe UI", 10.5F, FontStyle.Regular),
                .ForeColor = ThemePalette.TextMuted
            }

            panel.Controls.Add(subtitle)
            panel.Controls.Add(title)
            Return panel
        End Function

        Private Function BuildInvoicePanel() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)

            Dim grid As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 4,
                .Padding = New Padding(0, 10, 0, 0)
            }
            For index As Integer = 0 To 3
                grid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25.0F))
            Next
            grid.RowStyles.Add(New RowStyle(SizeType.Absolute, 72))
            grid.RowStyles.Add(New RowStyle(SizeType.Absolute, 72))

            grid.Controls.Add(CreateInputHost("Invoice Number", txtInvoiceNumber), 0, 0)
            grid.Controls.Add(CreateInputHost("Invoice Date", dtpInvoiceDate), 1, 0)
            grid.Controls.Add(CreateInputHost("Customer", cboCustomer), 2, 0)
            grid.Controls.Add(CreateInputHost("Payment Mode", cboPaymentMode), 3, 0)
            grid.Controls.Add(CreateInputHost("Amount Paid", nudAmountPaid), 0, 1)

            txtNotes.Multiline = True
            txtNotes.ScrollBars = ScrollBars.Vertical
            Dim notesHost As Control = CreateInputHost("Notes", txtNotes)
            grid.Controls.Add(notesHost, 1, 1)
            grid.SetColumnSpan(notesHost, 3)

            panel.Controls.Add(grid)
            Return panel
        End Function

        Private Function BuildAddItemPanel() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)

            Dim grid As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 6,
                .Padding = New Padding(0, 10, 0, 0)
            }
            grid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 40.0F))
            grid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 10.0F))
            grid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 10.0F))
            grid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 10.0F))
            grid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 15.0F))
            grid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 15.0F))
            grid.RowStyles.Add(New RowStyle(SizeType.Absolute, 72))
            grid.RowStyles.Add(New RowStyle(SizeType.Absolute, 50))

            grid.Controls.Add(CreateInputHost("Product", cboProduct), 0, 0)
            grid.Controls.Add(CreateInputHost("Qty", nudQuantity), 1, 0)
            grid.Controls.Add(CreateInputHost("Free Qty", nudFreeQuantity), 2, 0)
            grid.Controls.Add(CreateInputHost("Discount %", nudDiscount), 3, 0)
            grid.Controls.Add(CreateInputHost("Scheme", txtScheme), 4, 0)

            Dim buttonHost As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.LeftToRight,
                .WrapContents = False,
                .Padding = New Padding(0, 22, 0, 0),
                .BackColor = Color.Transparent
            }

            UiStyler.StylePrimaryButton(btnAddItem)
            btnAddItem.Text = "Add Item"
            btnAddItem.Width = 110

            UiStyler.StyleDangerButton(btnRemoveItem)
            btnRemoveItem.Text = "Remove Line"
            btnRemoveItem.Width = 120

            buttonHost.Controls.Add(btnAddItem)
            buttonHost.Controls.Add(btnRemoveItem)
            grid.Controls.Add(buttonHost, 5, 0)

            panel.Controls.Add(grid)
            Return panel
        End Function

        Private Function BuildGridPanel() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)
            dgvItems.Dock = DockStyle.Fill
            panel.Controls.Add(dgvItems)
            Return panel
        End Function

        Private Function BuildFooterPanel() As Control
            Dim split As New SplitContainer With {
                .Dock = DockStyle.Fill,
                .SplitterDistance = 520,
                .BackColor = ThemePalette.AppBackground
            }

            split.Panel1.Controls.Add(BuildTaxSummaryPanel())
            split.Panel2.Controls.Add(BuildTotalsPanel())
            Return split
        End Function

        Private Function BuildTaxSummaryPanel() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)

            Dim title As New Label With {
                .Dock = DockStyle.Top,
                .Height = 28,
                .Text = "Tax summary",
                .Font = New Font("Segoe UI Semibold", 13.0F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary
            }

            txtTaxSummary.Dock = DockStyle.Fill
            txtTaxSummary.Multiline = True
            txtTaxSummary.ReadOnly = True
            txtTaxSummary.BackColor = Color.White
            txtTaxSummary.BorderStyle = BorderStyle.FixedSingle
            txtTaxSummary.Font = New Font("Consolas", 10.0F, FontStyle.Regular)

            panel.Controls.Add(txtTaxSummary)
            panel.Controls.Add(title)
            Return panel
        End Function

        Private Function BuildTotalsPanel() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)

            Dim root As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2
            }
            root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 42.0F))
            root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 58.0F))

            Dim totalsTable As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2
            }
            totalsTable.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 52.0F))
            totalsTable.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 48.0F))
            For index As Integer = 0 To 5
                totalsTable.RowStyles.Add(New RowStyle(SizeType.Absolute, 24))
            Next

            AddTotalRow(totalsTable, 0, "Subtotal", lblSubTotalValue)
            AddTotalRow(totalsTable, 1, "Discount", lblDiscountValue)
            AddTotalRow(totalsTable, 2, "GST", lblGstValue)
            AddTotalRow(totalsTable, 3, "Round Off", lblRoundOffValue)
            AddTotalRow(totalsTable, 4, "Net Amount", lblNetAmountValue, True)
            AddTotalRow(totalsTable, 5, "Balance", lblBalanceValue, True)

            Dim buttonFlow As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.LeftToRight,
                .WrapContents = True,
                .Padding = New Padding(0, 34, 0, 0),
                .BackColor = Color.Transparent
            }

            UiStyler.StyleSecondaryButton(btnNewInvoice)
            btnNewInvoice.Text = "New Invoice"
            btnNewInvoice.Width = 120

            UiStyler.StylePrimaryButton(btnSaveInvoice)
            btnSaveInvoice.Text = "Save Invoice"
            btnSaveInvoice.Width = 120

            UiStyler.StyleSecondaryButton(btnExportInvoice)
            btnExportInvoice.Text = "Export Excel/PDF"
            btnExportInvoice.Width = 140

            UiStyler.StyleSecondaryButton(btnPrintPreview)
            btnPrintPreview.Text = "Print Preview"
            btnPrintPreview.Width = 125

            UiStyler.StyleSecondaryButton(btnPrintInvoice)
            btnPrintInvoice.Text = "Instant Print"
            btnPrintInvoice.Width = 115

            UiStyler.StyleSecondaryButton(btnOpenInvoiceFolder)
            btnOpenInvoiceFolder.Text = "Open Invoices"
            btnOpenInvoiceFolder.Width = 120

            UiStyler.StyleSecondaryButton(btnClose)
            btnClose.Text = "Close"
            btnClose.Width = 90

            buttonFlow.Controls.Add(btnNewInvoice)
            buttonFlow.Controls.Add(btnSaveInvoice)
            buttonFlow.Controls.Add(btnExportInvoice)
            buttonFlow.Controls.Add(btnPrintPreview)
            buttonFlow.Controls.Add(btnPrintInvoice)
            buttonFlow.Controls.Add(btnOpenInvoiceFolder)
            buttonFlow.Controls.Add(btnClose)

            root.Controls.Add(totalsTable, 0, 0)
            root.Controls.Add(buttonFlow, 1, 0)
            panel.Controls.Add(root)
            Return panel
        End Function

        Private Sub AddTotalRow(table As TableLayoutPanel, rowIndex As Integer, labelText As String, valueLabel As Label, Optional emphasize As Boolean = False)
            Dim caption As New Label With {
                .Dock = DockStyle.Fill,
                .Text = labelText,
                .ForeColor = ThemePalette.TextPrimary,
                .Font = New Font("Segoe UI", If(emphasize, 10.5F, 9.75F), If(emphasize, FontStyle.Bold, FontStyle.Regular)),
                .TextAlign = ContentAlignment.MiddleLeft
            }

            valueLabel.Dock = DockStyle.Fill
            valueLabel.ForeColor = ThemePalette.TextPrimary
            valueLabel.Font = New Font("Segoe UI Semibold", If(emphasize, 11.0F, 10.0F), FontStyle.Bold)
            valueLabel.TextAlign = ContentAlignment.MiddleRight
            valueLabel.Text = "0.00"

            table.Controls.Add(caption, 0, rowIndex)
            table.Controls.Add(valueLabel, 1, rowIndex)
        End Sub

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

        Private Sub ConfigureEditors()
            txtInvoiceNumber.ReadOnly = True
            txtInvoiceNumber.BorderStyle = BorderStyle.FixedSingle
            UiStyler.StyleInput(txtInvoiceNumber)

            dtpInvoiceDate.Format = DateTimePickerFormat.Custom
            dtpInvoiceDate.CustomFormat = "dd-MMM-yyyy"

            ConfigureCombo(cboCustomer)
            ConfigureCombo(cboPaymentMode)
            ConfigureCombo(cboProduct)

            cboPaymentMode.Items.AddRange(New Object() {"Cash", "Credit", "UPI", "Card", "Bank Transfer"})
            cboPaymentMode.SelectedIndex = 0

            txtNotes.BorderStyle = BorderStyle.FixedSingle
            UiStyler.StyleInput(txtNotes)
            txtScheme.BorderStyle = BorderStyle.FixedSingle
            UiStyler.StyleInput(txtScheme)

            ConfigureNumeric(nudAmountPaid, 2, 100000000D, 50D)
            ConfigureNumeric(nudQuantity, 0, 1000000D, 1D)
            ConfigureNumeric(nudFreeQuantity, 0, 1000000D, 1D)
            ConfigureNumeric(nudDiscount, 2, 100D, 1D)
            nudQuantity.Value = 1D
        End Sub

        Private Sub ConfigureCombo(combo As ComboBox)
            combo.DropDownStyle = ComboBoxStyle.DropDown
            combo.AutoCompleteMode = AutoCompleteMode.SuggestAppend
            combo.AutoCompleteSource = AutoCompleteSource.ListItems
            combo.Font = New Font("Segoe UI", 10.5F, FontStyle.Regular)
        End Sub

        Private Sub ConfigureNumeric(control As NumericUpDown, decimalPlaces As Integer, maximum As Decimal, increment As Decimal)
            control.DecimalPlaces = decimalPlaces
            control.Maximum = maximum
            control.Minimum = 0D
            control.Increment = increment
            control.ThousandsSeparator = True
            control.Font = New Font("Segoe UI", 10.5F, FontStyle.Regular)
        End Sub

        Private Sub ConfigureGrid()
            UiStyler.StyleDataGrid(dgvItems)
            dgvItems.DataSource = _invoiceItems
            dgvItems.Columns.Clear()

            dgvItems.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "LineNumber", .HeaderText = "#", .FillWeight = 35.0F, .ReadOnly = True})
            dgvItems.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "ProductName", .HeaderText = "Product", .FillWeight = 180.0F, .ReadOnly = True})
            dgvItems.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "BatchNumber", .HeaderText = "Batch", .FillWeight = 75.0F, .ReadOnly = True})
            dgvItems.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "ExpiryDate", .HeaderText = "Expiry", .FillWeight = 70.0F, .ReadOnly = True, .DefaultCellStyle = New DataGridViewCellStyle With {.Format = "dd-MMM-yyyy"}})
            dgvItems.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "AvailableStock", .HeaderText = "Stock", .FillWeight = 55.0F, .ReadOnly = True})
            dgvItems.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "Quantity", .HeaderText = "Qty", .FillWeight = 55.0F})
            dgvItems.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "FreeQuantity", .HeaderText = "Free", .FillWeight = 55.0F})
            dgvItems.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "Rate", .HeaderText = "Rate", .FillWeight = 70.0F, .DefaultCellStyle = New DataGridViewCellStyle With {.Format = "N2"}})
            dgvItems.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "DiscountPercentage", .HeaderText = "Disc %", .FillWeight = 60.0F, .DefaultCellStyle = New DataGridViewCellStyle With {.Format = "N2"}})
            dgvItems.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "SchemeDescription", .HeaderText = "Scheme", .FillWeight = 100.0F})
            dgvItems.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "GstPercentage", .HeaderText = "GST %", .FillWeight = 60.0F, .DefaultCellStyle = New DataGridViewCellStyle With {.Format = "N2"}})
            dgvItems.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "TaxableAmount", .HeaderText = "Taxable", .FillWeight = 80.0F, .ReadOnly = True, .DefaultCellStyle = New DataGridViewCellStyle With {.Format = "N2"}})
            dgvItems.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "GstAmount", .HeaderText = "GST", .FillWeight = 70.0F, .ReadOnly = True, .DefaultCellStyle = New DataGridViewCellStyle With {.Format = "N2"}})
            dgvItems.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "LineTotal", .HeaderText = "Total", .FillWeight = 80.0F, .ReadOnly = True, .DefaultCellStyle = New DataGridViewCellStyle With {.Format = "N2"}})
        End Sub

        Private Sub WireEvents()
            AddHandler Load, AddressOf FrmBilling_Load
            AddHandler btnAddItem.Click, AddressOf btnAddItem_Click
            AddHandler btnRemoveItem.Click, AddressOf btnRemoveItem_Click
            AddHandler btnNewInvoice.Click, AddressOf btnNewInvoice_Click
            AddHandler btnSaveInvoice.Click, AddressOf btnSaveInvoice_Click
            AddHandler btnExportInvoice.Click, AddressOf btnExportInvoice_Click
            AddHandler btnPrintPreview.Click, AddressOf btnPrintPreview_Click
            AddHandler btnPrintInvoice.Click, AddressOf btnPrintInvoice_Click
            AddHandler btnOpenInvoiceFolder.Click, AddressOf btnOpenInvoiceFolder_Click
            AddHandler btnClose.Click, AddressOf btnClose_Click
            AddHandler dgvItems.CellEndEdit, AddressOf dgvItems_CellEndEdit
            AddHandler dgvItems.DataError, AddressOf dgvItems_DataError
            AddHandler nudAmountPaid.ValueChanged, AddressOf nudAmountPaid_ValueChanged
            AddHandler dtpInvoiceDate.ValueChanged, AddressOf dtpInvoiceDate_ValueChanged
        End Sub

        Private Async Sub FrmBilling_Load(sender As Object, e As EventArgs)
            Await LoadLookupsAsync()
            Await PrepareNewInvoiceAsync()
        End Sub

        Private Async Function LoadLookupsAsync() As Task
            SetBusy(True, "Loading billing lookups...")

            Try
                _customers = Await _billingService.LoadCustomersAsync()
                _products = Await _billingService.LoadProductsAsync()

                cboCustomer.DataSource = Nothing
                cboCustomer.DataSource = _customers
                cboCustomer.SelectedIndex = -1

                cboProduct.DataSource = Nothing
                cboProduct.DataSource = _products.Where(Function(product) product.CurrentStock > 0).ToList()
                cboProduct.SelectedIndex = -1

                ShowStatus("Billing lookups loaded.", False)
            Catch ex As Exception
                AppLogger.Error("Billing lookups failed to load.", ex)
                ShowStatus("Billing lookups could not be loaded.", True)
            Finally
                SetBusy(False)
            End Try
        End Function

        Private Async Function PrepareNewInvoiceAsync() As Task
            _invoiceItems.Clear()
            txtNotes.Clear()
            nudAmountPaid.Value = 0D
            nudQuantity.Value = 1D
            nudFreeQuantity.Value = 0D
            nudDiscount.Value = 0D
            txtScheme.Clear()
            cboProduct.SelectedIndex = -1

            txtInvoiceNumber.Text = Await _billingService.GenerateNextInvoiceNumberAsync(dtpInvoiceDate.Value.Date)
            ReindexLines()
            RefreshTotals()
            UpdateExportActionState()
            ShowStatus("Ready to create a new invoice.", False)
        End Function

        Private Sub btnAddItem_Click(sender As Object, e As EventArgs)
            Dim product As ProductRecord = TryCast(cboProduct.SelectedItem, ProductRecord)
            If product Is Nothing Then
                ShowStatus("Select a product to add.", True)
                Return
            End If

            If product.CurrentStock <= 0 Then
                ShowStatus("This product has no available stock.", True)
                Return
            End If

            Dim existingLine As BillingLineItem = _invoiceItems.FirstOrDefault(Function(line) line.ProductId = product.Id AndAlso line.BatchNumber = product.BatchNumber)
            If existingLine IsNot Nothing Then
                existingLine.Quantity += Decimal.ToInt32(nudQuantity.Value)
                existingLine.FreeQuantity += Decimal.ToInt32(nudFreeQuantity.Value)
                existingLine.DiscountPercentage = nudDiscount.Value
                existingLine.SchemeDescription = txtScheme.Text.Trim()
                _billingService.RecalculateLine(existingLine)
                dgvItems.Refresh()
            Else
                Dim line As BillingLineItem = _billingService.CreateLineFromProduct(product)
                line.Quantity = Decimal.ToInt32(nudQuantity.Value)
                line.FreeQuantity = Decimal.ToInt32(nudFreeQuantity.Value)
                line.DiscountPercentage = nudDiscount.Value
                line.SchemeDescription = txtScheme.Text.Trim()
                _billingService.RecalculateLine(line)
                _invoiceItems.Add(line)
            End If

            ReindexLines()
            RefreshTotals()
            nudQuantity.Value = 1D
            nudFreeQuantity.Value = 0D
            nudDiscount.Value = 0D
            txtScheme.Clear()
            ShowStatus($"Added '{product.ProductName}' to the invoice.", False)
        End Sub

        Private Sub btnRemoveItem_Click(sender As Object, e As EventArgs)
            If dgvItems.CurrentRow Is Nothing Then
                ShowStatus("Select a line to remove.", True)
                Return
            End If

            Dim line As BillingLineItem = TryCast(dgvItems.CurrentRow.DataBoundItem, BillingLineItem)
            If line Is Nothing Then
                Return
            End If

            _invoiceItems.Remove(line)
            ReindexLines()
            RefreshTotals()
            ShowStatus("Invoice line removed.", False)
        End Sub

        Private Async Sub btnNewInvoice_Click(sender As Object, e As EventArgs)
            Await PrepareNewInvoiceAsync()
        End Sub

        Private Async Sub btnSaveInvoice_Click(sender As Object, e As EventArgs)
            Dim customer As CustomerRecord = TryCast(cboCustomer.SelectedItem, CustomerRecord)
            Dim draft As New BillingInvoiceDraft With {
                .InvoiceNumber = txtInvoiceNumber.Text,
                .InvoiceDate = dtpInvoiceDate.Value.Date,
                .CustomerId = If(customer Is Nothing, 0, customer.Id),
                .CustomerName = If(customer Is Nothing, String.Empty, customer.CustomerName),
                .PaymentMode = If(cboPaymentMode.SelectedItem, String.Empty).ToString(),
                .AmountPaid = nudAmountPaid.Value,
                .Notes = txtNotes.Text
            }

            For Each line As BillingLineItem In _invoiceItems
                draft.Items.Add(line)
            Next

            draft.Summary = _billingService.CalculateTotals(draft.Items, draft.AmountPaid)

            SetBusy(True, "Saving invoice...")
            Dim result As InvoiceSaveResult = Await _billingService.SaveInvoiceAsync(draft, If(SessionManager.CurrentUser Is Nothing, 0, SessionManager.CurrentUser.Id))
            If Not result.IsSuccess Then
                SetBusy(False)
                ShowStatus(result.Message, True)
                Return
            End If

            _lastSavedInvoiceId = result.InvoiceId
            _lastSavedInvoiceNumber = result.InvoiceNumber

            Dim exportResult As InvoiceExportResult = Await _invoiceExportService.GenerateInvoiceFilesAsync(result.InvoiceId)
            SetBusy(False)

            Await LoadLookupsAsync()
            Await PrepareNewInvoiceAsync()
            UpdateExportActionState()

            If exportResult.IsSuccess Then
                ShowStatus($"{result.Message} Excel and PDF generated.", False)
            Else
                ShowStatus($"{result.Message} Export is pending. Use Export Excel/PDF to retry.", True)
            End If
        End Sub

        Private Async Sub btnExportInvoice_Click(sender As Object, e As EventArgs)
            If Not EnsureSavedInvoiceAvailable() Then
                Return
            End If

            SetBusy(True, "Generating invoice files...")
            Dim result As InvoiceExportResult = Await _invoiceExportService.GenerateInvoiceFilesAsync(_lastSavedInvoiceId)
            SetBusy(False)
            ShowStatus(result.Message, Not result.IsSuccess)
        End Sub

        Private Sub btnPrintPreview_Click(sender As Object, e As EventArgs)
            If Not EnsureSavedInvoiceAvailable() Then
                Return
            End If

            Try
                _invoiceExportService.ShowPrintPreview(_lastSavedInvoiceId)
                ShowStatus($"Preview opened for invoice {_lastSavedInvoiceNumber}.", False)
            Catch ex As Exception
                AppLogger.Error($"Invoice preview failed for invoice Id {_lastSavedInvoiceId}.", ex)
                ShowStatus("Print preview could not be opened.", True)
            End Try
        End Sub

        Private Sub btnPrintInvoice_Click(sender As Object, e As EventArgs)
            If Not EnsureSavedInvoiceAvailable() Then
                Return
            End If

            Try
                _invoiceExportService.PrintInvoice(_lastSavedInvoiceId)
                ShowStatus($"Invoice {_lastSavedInvoiceNumber} sent to the default printer.", False)
            Catch ex As Exception
                AppLogger.Error($"Invoice print failed for invoice Id {_lastSavedInvoiceId}.", ex)
                ShowStatus("Invoice could not be printed.", True)
            End Try
        End Sub

        Private Sub btnOpenInvoiceFolder_Click(sender As Object, e As EventArgs)
            Try
                _invoiceExportService.OpenInvoiceFolder()
                ShowStatus("Opened generated invoices folder.", False)
            Catch ex As Exception
                AppLogger.Error("Generated invoices folder could not be opened.", ex)
                ShowStatus("Invoice folder could not be opened.", True)
            End Try
        End Sub

        Private Sub btnClose_Click(sender As Object, e As EventArgs)
            Close()
        End Sub

        Private Sub dgvItems_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs)
            If e.RowIndex < 0 OrElse e.RowIndex >= _invoiceItems.Count Then
                Return
            End If

            Dim line As BillingLineItem = _invoiceItems(e.RowIndex)
            If line.Quantity + line.FreeQuantity > line.AvailableStock Then
                ShowStatus($"Stock exceeded for '{line.ProductName}'.", True)
                line.Quantity = Math.Max(line.AvailableStock - line.FreeQuantity, 0)
            End If

            _billingService.RecalculateLine(line)
            dgvItems.Refresh()
            RefreshTotals()
        End Sub

        Private Sub dgvItems_DataError(sender As Object, e As DataGridViewDataErrorEventArgs)
            e.Cancel = True
            ShowStatus("Enter a valid numeric value in the invoice grid.", True)
        End Sub

        Private Sub nudAmountPaid_ValueChanged(sender As Object, e As EventArgs)
            RefreshTotals()
        End Sub

        Private Async Sub dtpInvoiceDate_ValueChanged(sender As Object, e As EventArgs)
            If _invoiceItems.Count = 0 AndAlso Not _isBusy Then
                txtInvoiceNumber.Text = Await _billingService.GenerateNextInvoiceNumberAsync(dtpInvoiceDate.Value.Date)
            End If
        End Sub

        Private Sub ReindexLines()
            For index As Integer = 0 To _invoiceItems.Count - 1
                _invoiceItems(index).LineNumber = index + 1
            Next
            dgvItems.Refresh()
        End Sub

        Private Sub RefreshTotals()
            Dim summary As BillingTotalsSummary = _billingService.CalculateTotals(_invoiceItems.ToList(), nudAmountPaid.Value)
            lblSubTotalValue.Text = summary.SubTotal.ToString("N2")
            lblDiscountValue.Text = summary.DiscountAmount.ToString("N2")
            lblGstValue.Text = summary.GstAmount.ToString("N2")
            lblRoundOffValue.Text = summary.RoundOffAmount.ToString("N2")
            lblNetAmountValue.Text = summary.NetAmount.ToString("N2")
            lblBalanceValue.Text = summary.BalanceAmount.ToString("N2")
            txtTaxSummary.Text = summary.TaxSummaryText
        End Sub

        Private Sub SetBusy(isBusy As Boolean, Optional message As String = "")
            _isBusy = isBusy

            cboCustomer.Enabled = Not isBusy
            cboProduct.Enabled = Not isBusy
            cboPaymentMode.Enabled = Not isBusy
            nudAmountPaid.Enabled = Not isBusy
            txtNotes.Enabled = Not isBusy
            nudQuantity.Enabled = Not isBusy
            nudFreeQuantity.Enabled = Not isBusy
            nudDiscount.Enabled = Not isBusy
            txtScheme.Enabled = Not isBusy
            btnAddItem.Enabled = Not isBusy
            btnRemoveItem.Enabled = Not isBusy
            btnNewInvoice.Enabled = Not isBusy
            btnSaveInvoice.Enabled = Not isBusy
            btnExportInvoice.Enabled = Not isBusy AndAlso _lastSavedInvoiceId > 0
            btnPrintPreview.Enabled = Not isBusy AndAlso _lastSavedInvoiceId > 0
            btnPrintInvoice.Enabled = Not isBusy AndAlso _lastSavedInvoiceId > 0
            btnOpenInvoiceFolder.Enabled = Not isBusy
            btnClose.Enabled = Not isBusy
            dgvItems.Enabled = Not isBusy
            dtpInvoiceDate.Enabled = Not isBusy

            If isBusy Then
                lblStatus.ForeColor = ThemePalette.TextMuted
                lblStatus.Text = message
            End If
        End Sub

        Private Sub ShowStatus(message As String, isError As Boolean)
            lblStatus.ForeColor = If(isError, ThemePalette.DangerRed, ThemePalette.AccentGreen)
            lblStatus.Text = message
        End Sub

        Private Function EnsureSavedInvoiceAvailable() As Boolean
            If _lastSavedInvoiceId > 0 Then
                Return True
            End If

            ShowStatus("Save an invoice first to export, preview, or print it.", True)
            Return False
        End Function

        Private Sub UpdateExportActionState()
            btnExportInvoice.Enabled = _lastSavedInvoiceId > 0 AndAlso Not _isBusy
            btnPrintPreview.Enabled = _lastSavedInvoiceId > 0 AndAlso Not _isBusy
            btnPrintInvoice.Enabled = _lastSavedInvoiceId > 0 AndAlso Not _isBusy
            btnOpenInvoiceFolder.Enabled = Not _isBusy
        End Sub

        Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
            Select Case keyData
                Case Keys.F2
                    cboCustomer.Focus()
                    Return True
                Case Keys.F3
                    cboProduct.Focus()
                    Return True
                Case Keys.Control Or Keys.S
                    btnSaveInvoice.PerformClick()
                    Return True
                Case Keys.Control Or Keys.N
                    btnNewInvoice.PerformClick()
                    Return True
                Case Keys.Control Or Keys.E
                    btnExportInvoice.PerformClick()
                    Return True
                Case Keys.Control Or Keys.P
                    btnPrintInvoice.PerformClick()
                    Return True
                Case Keys.Control Or Keys.Shift Or Keys.P
                    btnPrintPreview.PerformClick()
                    Return True
                Case Keys.Delete
                    btnRemoveItem.PerformClick()
                    Return True
                Case Keys.Escape
                    Close()
                    Return True
            End Select

            Return MyBase.ProcessCmdKey(msg, keyData)
        End Function

    End Class

End Namespace
