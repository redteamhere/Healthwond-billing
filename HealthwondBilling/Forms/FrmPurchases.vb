Imports HealthwondBilling.Models
Imports HealthwondBilling.Services
Imports HealthwondBilling.Utilities
Imports System.ComponentModel
Imports System.Linq

Namespace Forms

    Public Class FrmPurchases
        Inherits Form

        Private ReadOnly _purchaseService As PurchaseService
        Private ReadOnly _supplierService As SupplierService
        Private ReadOnly _purchasePrintService As PurchasePrintService
        Private ReadOnly _purchaseItems As New BindingList(Of PurchaseLineItem)()

        Private _suppliers As New List(Of SupplierRecord)()
        Private _products As New List(Of ProductRecord)()
        Private _currentSupplierId As Integer
        Private _isBusy As Boolean
        Private _lastSavedPurchaseId As Integer

        Private ReadOnly txtPurchaseNumber As New TextBox()
        Private ReadOnly dtpPurchaseDate As New DateTimePicker()
        Private ReadOnly cboSupplier As New ComboBox()
        Private ReadOnly txtSupplierInvoiceNumber As New TextBox()
        Private ReadOnly dtpSupplierInvoiceDate As New DateTimePicker()
        Private ReadOnly txtPurchaseOrderNumber As New TextBox()
        Private ReadOnly dtpPurchaseOrderDate As New DateTimePicker()
        Private ReadOnly txtPlaceOfSupply As New TextBox()
        Private ReadOnly nudCaseCount As New NumericUpDown()
        Private ReadOnly txtTransportName As New TextBox()
        Private ReadOnly txtEwayBillNumber As New TextBox()
        Private ReadOnly txtNotes As New TextBox()

        Private ReadOnly txtSupplierName As New TextBox()
        Private ReadOnly txtSupplierPhone As New TextBox()
        Private ReadOnly txtSupplierGstin As New TextBox()
        Private ReadOnly txtSupplierEmail As New TextBox()
        Private ReadOnly txtSupplierLicense As New TextBox()
        Private ReadOnly txtSupplierAddress As New TextBox()
        Private ReadOnly nudSupplierOutstanding As New NumericUpDown()
        Private ReadOnly btnNewSupplier As New Button()
        Private ReadOnly btnSaveSupplier As New Button()
        Private ReadOnly btnDeleteSupplier As New Button()

        Private ReadOnly cboProduct As New ComboBox()
        Private ReadOnly txtPacking As New TextBox()
        Private ReadOnly txtHsnCode As New TextBox()
        Private ReadOnly txtBatchNumber As New TextBox()
        Private ReadOnly dtpExpiryDate As New DateTimePicker()
        Private ReadOnly nudQuantity As New NumericUpDown()
        Private ReadOnly nudFreeQuantity As New NumericUpDown()
        Private ReadOnly nudPTR As New NumericUpDown()
        Private ReadOnly nudPTS As New NumericUpDown()
        Private ReadOnly nudMRP As New NumericUpDown()
        Private ReadOnly nudGST As New NumericUpDown()
        Private ReadOnly btnAddItem As New Button()
        Private ReadOnly btnRemoveItem As New Button()

        Private ReadOnly dgvItems As New DataGridView()
        Private ReadOnly btnNewPurchase As New Button()
        Private ReadOnly btnSavePurchase As New Button()
        Private ReadOnly btnPrintPreview As New Button()
        Private ReadOnly btnPrint As New Button()
        Private ReadOnly btnClose As New Button()
        Private ReadOnly lblStatus As New Label()

        Private ReadOnly lblSubTotalValue As New Label()
        Private ReadOnly lblGstValue As New Label()
        Private ReadOnly lblRoundOffValue As New Label()
        Private ReadOnly lblNetAmountValue As New Label()
        Private ReadOnly txtTaxSummary As New TextBox()

        Public Sub New(purchaseService As PurchaseService, supplierService As SupplierService, purchasePrintService As PurchasePrintService)
            _purchaseService = purchaseService
            _supplierService = supplierService
            _purchasePrintService = purchasePrintService

            Text = "Healthwond Billing System - Purchases"
            StartPosition = FormStartPosition.CenterParent
            WindowState = FormWindowState.Maximized
            MinimumSize = New Size(1440, 900)
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
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 250))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 154))
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 170))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 36))

            root.Controls.Add(BuildHeaderPanel(), 0, 0)
            root.Controls.Add(BuildTopSection(), 0, 1)
            root.Controls.Add(BuildItemEntryPanel(), 0, 2)
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
                .Text = "Purchases",
                .Font = New Font("Segoe UI Semibold", 24.0F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary
            }

            Dim subtitle As New Label With {
                .Dock = DockStyle.Fill,
                .Text = "Record supplier purchases, update inventory batch stock, maintain supplier masters, and post stock ledger entries through one controlled workflow.",
                .Font = New Font("Segoe UI", 10.5F, FontStyle.Regular),
                .ForeColor = ThemePalette.TextMuted
            }

            panel.Controls.Add(subtitle)
            panel.Controls.Add(title)
            Return panel
        End Function

        Private Function BuildTopSection() As Control
            Dim split As New SplitContainer With {
                .Dock = DockStyle.Fill,
                .SplitterDistance = 500,
                .BackColor = ThemePalette.AppBackground
            }

            split.Panel1.Controls.Add(BuildSupplierPanel())
            split.Panel2.Controls.Add(BuildPurchaseHeaderPanel())
            Return split
        End Function

        Private Function BuildSupplierPanel() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)

            Dim title As New Label With {
                .Dock = DockStyle.Top,
                .Height = 30,
                .Text = "Supplier master",
                .Font = New Font("Segoe UI Semibold", 14.0F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary
            }

            Dim actionFlow As New FlowLayoutPanel With {
                .Dock = DockStyle.Bottom,
                .Height = 58,
                .FlowDirection = FlowDirection.LeftToRight,
                .WrapContents = False,
                .BackColor = Color.Transparent,
                .Padding = New Padding(0, 10, 0, 0)
            }

            UiStyler.StyleSecondaryButton(btnNewSupplier)
            btnNewSupplier.Text = "New Supplier"
            btnNewSupplier.Width = 120

            UiStyler.StylePrimaryButton(btnSaveSupplier)
            btnSaveSupplier.Text = "Save Supplier"
            btnSaveSupplier.Width = 120

            UiStyler.StyleDangerButton(btnDeleteSupplier)
            btnDeleteSupplier.Text = "Delete Supplier"
            btnDeleteSupplier.Width = 130

            actionFlow.Controls.Add(btnNewSupplier)
            actionFlow.Controls.Add(btnSaveSupplier)
            actionFlow.Controls.Add(btnDeleteSupplier)

            Dim editor As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2,
                .Padding = New Padding(0, 14, 0, 0)
            }
            editor.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0F))
            editor.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0F))
            For index As Integer = 0 To 2
                editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 72))
            Next
            editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 108))
            editor.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            editor.Controls.Add(CreateInputHost("Supplier Name", txtSupplierName), 0, 0)
            editor.Controls.Add(CreateInputHost("Phone", txtSupplierPhone), 1, 0)
            editor.Controls.Add(CreateInputHost("GSTIN", txtSupplierGstin), 0, 1)
            editor.Controls.Add(CreateInputHost("Email", txtSupplierEmail), 1, 1)
            editor.Controls.Add(CreateInputHost("Drug License Number", txtSupplierLicense), 0, 2)
            editor.Controls.Add(CreateInputHost("Outstanding Balance", nudSupplierOutstanding), 1, 2)

            txtSupplierAddress.Multiline = True
            txtSupplierAddress.ScrollBars = ScrollBars.Vertical
            Dim addressHost As Control = CreateInputHost("Address", txtSupplierAddress)
            editor.Controls.Add(addressHost, 0, 3)
            editor.SetColumnSpan(addressHost, 2)

            panel.Controls.Add(UiStyler.CreateScrollableHost(editor))
            panel.Controls.Add(actionFlow)
            panel.Controls.Add(title)
            Return panel
        End Function

        Private Function BuildPurchaseHeaderPanel() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)

            Dim title As New Label With {
                .Dock = DockStyle.Top,
                .Height = 30,
                .Text = "Purchase header",
                .Font = New Font("Segoe UI Semibold", 14.0F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary
            }

            Dim editor As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 4,
                .Padding = New Padding(0, 14, 0, 0)
            }
            For index As Integer = 0 To 3
                editor.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25.0F))
            Next
            editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 68))
            editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 68))
            editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 68))
            editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 108))
            editor.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            editor.Controls.Add(CreateInputHost("Purchase Number", txtPurchaseNumber), 0, 0)
            editor.Controls.Add(CreateInputHost("Purchase Date", dtpPurchaseDate), 1, 0)
            editor.Controls.Add(CreateInputHost("Supplier Invoice Number", txtSupplierInvoiceNumber), 2, 0)
            editor.Controls.Add(CreateInputHost("Supplier Invoice Date", dtpSupplierInvoiceDate), 3, 0)

            Dim supplierHost As Control = CreateInputHost("Supplier", cboSupplier)
            editor.Controls.Add(supplierHost, 0, 1)
            editor.SetColumnSpan(supplierHost, 2)
            editor.Controls.Add(CreateInputHost("P.O. Number", txtPurchaseOrderNumber), 2, 1)
            editor.Controls.Add(CreateInputHost("P.O. Date", dtpPurchaseOrderDate), 3, 1)

            editor.Controls.Add(CreateInputHost("Place of Supply", txtPlaceOfSupply), 0, 2)
            editor.Controls.Add(CreateInputHost("Cases", nudCaseCount), 1, 2)
            editor.Controls.Add(CreateInputHost("Transport", txtTransportName), 2, 2)
            editor.Controls.Add(CreateInputHost("E-Way Bill Number", txtEwayBillNumber), 3, 2)

            txtNotes.Multiline = True
            txtNotes.ScrollBars = ScrollBars.Vertical
            Dim notesHost As Control = CreateInputHost("Notes", txtNotes)
            editor.Controls.Add(notesHost, 0, 3)
            editor.SetColumnSpan(notesHost, 4)

            panel.Controls.Add(UiStyler.CreateScrollableHost(editor))
            panel.Controls.Add(title)
            Return panel
        End Function

        Private Function BuildItemEntryPanel() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)

            Dim grid As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 12,
                .Padding = New Padding(0, 10, 0, 0)
            }
            grid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 19.0F))
            grid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 8.0F))
            grid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 9.0F))
            grid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 10.0F))
            grid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 8.0F))
            grid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 6.0F))
            grid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 6.0F))
            grid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 7.0F))
            grid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 7.0F))
            grid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 7.0F))
            grid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 6.0F))
            grid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 17.0F))
            grid.RowStyles.Add(New RowStyle(SizeType.Absolute, 72))
            grid.RowStyles.Add(New RowStyle(SizeType.Absolute, 50))

            grid.Controls.Add(CreateInputHost("Product", cboProduct), 0, 0)
            grid.Controls.Add(CreateInputHost("Packing", txtPacking), 1, 0)
            grid.Controls.Add(CreateInputHost("HSN", txtHsnCode), 2, 0)
            grid.Controls.Add(CreateInputHost("Batch", txtBatchNumber), 3, 0)
            grid.Controls.Add(CreateInputHost("Expiry", dtpExpiryDate), 4, 0)
            grid.Controls.Add(CreateInputHost("Qty", nudQuantity), 5, 0)
            grid.Controls.Add(CreateInputHost("Free Qty", nudFreeQuantity), 6, 0)
            grid.Controls.Add(CreateInputHost("PTR", nudPTR), 7, 0)
            grid.Controls.Add(CreateInputHost("PTS", nudPTS), 8, 0)
            grid.Controls.Add(CreateInputHost("MRP", nudMRP), 9, 0)
            grid.Controls.Add(CreateInputHost("GST %", nudGST), 10, 0)

            Dim buttonHost As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.LeftToRight,
                .WrapContents = True,
                .Padding = New Padding(0, 22, 0, 0),
                .BackColor = Color.Transparent
            }

            UiStyler.StylePrimaryButton(btnAddItem)
            btnAddItem.Text = "Add Line"
            btnAddItem.Width = 100

            UiStyler.StyleDangerButton(btnRemoveItem)
            btnRemoveItem.Text = "Remove Line"
            btnRemoveItem.Width = 120

            buttonHost.Controls.Add(btnAddItem)
            buttonHost.Controls.Add(btnRemoveItem)
            grid.Controls.Add(buttonHost, 11, 0)

            panel.Controls.Add(UiStyler.CreateScrollableHost(grid))
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
                .SplitterDistance = 540,
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
            root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 52.0F))
            root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 48.0F))

            Dim totalsTable As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2
            }
            totalsTable.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 52.0F))
            totalsTable.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 48.0F))
            For index As Integer = 0 To 3
                totalsTable.RowStyles.Add(New RowStyle(SizeType.Absolute, 28))
            Next

            AddTotalRow(totalsTable, 0, "Subtotal", lblSubTotalValue)
            AddTotalRow(totalsTable, 1, "GST", lblGstValue)
            AddTotalRow(totalsTable, 2, "Round Off", lblRoundOffValue)
            AddTotalRow(totalsTable, 3, "Net Amount", lblNetAmountValue, True)

            Dim buttonFlow As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.LeftToRight,
                .WrapContents = True,
                .Padding = New Padding(0, 48, 0, 0),
                .BackColor = Color.Transparent
            }

            UiStyler.StyleSecondaryButton(btnNewPurchase)
            btnNewPurchase.Text = "New Purchase"
            btnNewPurchase.Width = 130

            UiStyler.StylePrimaryButton(btnSavePurchase)
            btnSavePurchase.Text = "Save Purchase"
            btnSavePurchase.Width = 130

            UiStyler.StyleSecondaryButton(btnPrintPreview)
            btnPrintPreview.Text = "Print Preview"
            btnPrintPreview.Width = 130

            UiStyler.StylePrimaryButton(btnPrint)
            btnPrint.Text = "Print"
            btnPrint.Width = 100

            UiStyler.StyleSecondaryButton(btnClose)
            btnClose.Text = "Close"
            btnClose.Width = 90

            buttonFlow.Controls.Add(btnNewPurchase)
            buttonFlow.Controls.Add(btnSavePurchase)
            buttonFlow.Controls.Add(btnPrintPreview)
            buttonFlow.Controls.Add(btnPrint)
            buttonFlow.Controls.Add(btnClose)

            root.Controls.Add(totalsTable, 0, 0)
            root.Controls.Add(buttonFlow, 1, 0)
            panel.Controls.Add(UiStyler.CreateScrollableHost(root))
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
            txtPurchaseNumber.ReadOnly = True
            txtPurchaseNumber.BorderStyle = BorderStyle.FixedSingle
            UiStyler.StyleInput(txtPurchaseNumber)

            ConfigureDatePicker(dtpPurchaseDate)
            ConfigureOptionalDatePicker(dtpSupplierInvoiceDate)
            ConfigureOptionalDatePicker(dtpPurchaseOrderDate)

            ConfigureCombo(cboSupplier)
            ConfigureCombo(cboProduct)

            For Each editorTextBox As TextBox In New TextBox() {txtSupplierInvoiceNumber, txtPurchaseOrderNumber, txtPlaceOfSupply, txtTransportName, txtEwayBillNumber, txtNotes, txtSupplierName, txtSupplierPhone, txtSupplierGstin, txtSupplierEmail, txtSupplierLicense, txtSupplierAddress, txtPacking, txtHsnCode, txtBatchNumber}
                editorTextBox.BorderStyle = BorderStyle.FixedSingle
                UiStyler.StyleInput(editorTextBox)
            Next

            txtPacking.ReadOnly = True
            txtHsnCode.ReadOnly = True

            ConfigureNumeric(nudSupplierOutstanding, 2, 100000000D, 50D)
            ConfigureNumeric(nudCaseCount, 0, 1000000D, 1D)
            ConfigureNumeric(nudQuantity, 0, 1000000D, 1D)
            ConfigureNumeric(nudFreeQuantity, 0, 1000000D, 1D)
            ConfigureNumeric(nudPTR, 2, 1000000D, 1D)
            ConfigureNumeric(nudPTS, 2, 1000000D, 1D)
            ConfigureNumeric(nudMRP, 2, 1000000D, 1D)
            ConfigureNumeric(nudGST, 2, 100D, 0.5D)

            nudQuantity.Value = 1D
            ConfigureDatePicker(dtpExpiryDate)
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

        Private Sub ConfigureDatePicker(control As DateTimePicker)
            control.Format = DateTimePickerFormat.Custom
            control.CustomFormat = "dd-MMM-yyyy"
        End Sub

        Private Sub ConfigureOptionalDatePicker(control As DateTimePicker)
            ConfigureDatePicker(control)
            control.ShowCheckBox = True
            control.Checked = False
        End Sub

        Private Sub ConfigureGrid()
            UiStyler.StyleDataGrid(dgvItems)
            dgvItems.ReadOnly = False
            dgvItems.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgvItems.ScrollBars = ScrollBars.Both
            dgvItems.DataSource = _purchaseItems
            dgvItems.Columns.Clear()

            dgvItems.Columns.Add(CreateGridColumn("LineNumber", "#", 40, True))
            dgvItems.Columns.Add(CreateGridColumn("ProductName", "Product", 180, True))
            dgvItems.Columns.Add(CreateGridColumn("Packing", "Pack", 70, True))
            dgvItems.Columns.Add(CreateGridColumn("HsnCode", "HSN", 90, True))
            dgvItems.Columns.Add(CreateGridColumn("BatchNumber", "Batch", 90))
            dgvItems.Columns.Add(CreateGridColumn("ExpiryDate", "Expiry", 90, True, "dd-MMM-yyyy"))
            dgvItems.Columns.Add(CreateGridColumn("ExistingStock", "Stock", 65, True))
            dgvItems.Columns.Add(CreateGridColumn("Quantity", "Qty", 55))
            dgvItems.Columns.Add(CreateGridColumn("FreeQuantity", "Free", 55))
            dgvItems.Columns.Add(CreateGridColumn("PTR", "PTR", 70, False, "N2"))
            dgvItems.Columns.Add(CreateGridColumn("PTS", "PTS", 70, False, "N2"))
            dgvItems.Columns.Add(CreateGridColumn("MRP", "MRP", 70, False, "N2"))
            dgvItems.Columns.Add(CreateGridColumn("GstPercentage", "GST %", 60, False, "N2"))
            dgvItems.Columns.Add(CreateGridColumn("TaxableAmount", "Taxable", 90, True, "N2"))
            dgvItems.Columns.Add(CreateGridColumn("GstAmount", "GST", 80, True, "N2"))
            dgvItems.Columns.Add(CreateGridColumn("LineTotal", "Total", 90, True, "N2"))
        End Sub

        Private Function CreateGridColumn(propertyName As String, headerText As String, width As Integer, Optional [readOnly] As Boolean = False, Optional format As String = Nothing) As DataGridViewTextBoxColumn
            Dim column As New DataGridViewTextBoxColumn With {
                .DataPropertyName = propertyName,
                .HeaderText = headerText,
                .Width = width,
                .ReadOnly = [readOnly]
            }

            If Not String.IsNullOrWhiteSpace(format) Then
                column.DefaultCellStyle = New DataGridViewCellStyle With {.Format = format}
            End If

            Return column
        End Function

        Private Sub WireEvents()
            AddHandler Load, AddressOf FrmPurchases_Load
            AddHandler btnNewSupplier.Click, AddressOf btnNewSupplier_Click
            AddHandler btnSaveSupplier.Click, AddressOf btnSaveSupplier_Click
            AddHandler btnDeleteSupplier.Click, AddressOf btnDeleteSupplier_Click
            AddHandler btnAddItem.Click, AddressOf btnAddItem_Click
            AddHandler btnRemoveItem.Click, AddressOf btnRemoveItem_Click
            AddHandler btnNewPurchase.Click, AddressOf btnNewPurchase_Click
            AddHandler btnSavePurchase.Click, AddressOf btnSavePurchase_Click
            AddHandler btnPrintPreview.Click, AddressOf btnPrintPreview_Click
            AddHandler btnPrint.Click, AddressOf btnPrint_Click
            AddHandler btnClose.Click, AddressOf btnClose_Click
            AddHandler cboSupplier.SelectedIndexChanged, AddressOf cboSupplier_SelectedIndexChanged
            AddHandler cboProduct.SelectedIndexChanged, AddressOf cboProduct_SelectedIndexChanged
            AddHandler dgvItems.CellEndEdit, AddressOf dgvItems_CellEndEdit
            AddHandler dgvItems.DataError, AddressOf dgvItems_DataError
            AddHandler dtpPurchaseDate.ValueChanged, AddressOf dtpPurchaseDate_ValueChanged
        End Sub

        Private Async Sub FrmPurchases_Load(sender As Object, e As EventArgs)
            PrepareNewSupplier(False)
            Await LoadLookupsAsync()
            Await PrepareNewPurchaseAsync()
        End Sub

        Private Async Function LoadLookupsAsync(Optional preferredSupplierId As Integer = 0) As Task
            SetBusy(True, "Loading purchase lookups...")

            Try
                Dim retainedSupplierId As Integer = If(preferredSupplierId > 0, preferredSupplierId, _currentSupplierId)

                _suppliers = Await _supplierService.SearchAsync(String.Empty)
                _products = Await _purchaseService.LoadProductsAsync()

                cboSupplier.DataSource = Nothing
                cboSupplier.DataSource = _suppliers
                cboSupplier.SelectedIndex = -1

                cboProduct.DataSource = Nothing
                cboProduct.DataSource = _products
                cboProduct.SelectedIndex = -1

                If _suppliers.Count > 0 Then
                    SelectSupplier(retainedSupplierId)
                Else
                    PrepareNewSupplier(False)
                End If

                ShowStatus("Purchase lookups loaded.", False)
            Catch ex As Exception
                AppLogger.Error("Purchase lookups failed to load.", ex)
                ShowStatus("Purchase lookups could not be loaded.", True)
            Finally
                SetBusy(False)
            End Try
        End Function

        Private Sub SelectSupplier(preferredSupplierId As Integer)
            If cboSupplier.Items.Count = 0 Then
                Return
            End If

            Dim targetIndex As Integer = 0
            If preferredSupplierId > 0 Then
                For index As Integer = 0 To _suppliers.Count - 1
                    If _suppliers(index).Id = preferredSupplierId Then
                        targetIndex = index
                        Exit For
                    End If
                Next
            End If

            cboSupplier.SelectedIndex = targetIndex
            BindSupplier(TryCast(cboSupplier.SelectedItem, SupplierRecord))
        End Sub

        Private Async Function PrepareNewPurchaseAsync() As Task
            _purchaseItems.Clear()
            txtSupplierInvoiceNumber.Clear()
            ResetOptionalDatePicker(dtpSupplierInvoiceDate)
            txtPurchaseOrderNumber.Clear()
            ResetOptionalDatePicker(dtpPurchaseOrderDate)
            txtPlaceOfSupply.Clear()
            nudCaseCount.Value = 0D
            txtTransportName.Clear()
            txtEwayBillNumber.Clear()
            txtNotes.Clear()
            cboProduct.SelectedIndex = -1
            ResetItemEntry()

            txtPurchaseNumber.Text = Await _purchaseService.GenerateNextPurchaseNumberAsync(dtpPurchaseDate.Value.Date)
            ReindexLines()
            RefreshTotals()
            ShowStatus("Ready to record a new purchase.", False)
        End Function

        Private Sub PrepareNewSupplier(Optional clearSelection As Boolean = True)
            _currentSupplierId = 0
            txtSupplierName.Clear()
            txtSupplierPhone.Clear()
            txtSupplierGstin.Clear()
            txtSupplierEmail.Clear()
            txtSupplierLicense.Clear()
            txtSupplierAddress.Clear()
            nudSupplierOutstanding.Value = 0D

            If clearSelection Then
                cboSupplier.SelectedIndex = -1
            End If

            UpdateSupplierActionState()
        End Sub

        Private Sub BindSupplier(supplier As SupplierRecord)
            If supplier Is Nothing Then
                PrepareNewSupplier(False)
                Return
            End If

            _currentSupplierId = supplier.Id
            txtSupplierName.Text = supplier.SupplierName
            txtSupplierPhone.Text = supplier.Phone
            txtSupplierGstin.Text = supplier.Gstin
            txtSupplierEmail.Text = supplier.Email
            txtSupplierLicense.Text = supplier.DrugLicenseNumber
            txtSupplierAddress.Text = supplier.Address
            nudSupplierOutstanding.Value = Math.Min(nudSupplierOutstanding.Maximum, supplier.OutstandingBalance)
            UpdateSupplierActionState()
        End Sub

        Private Function ReadSupplierFromForm() As SupplierRecord
            Return New SupplierRecord With {
                .Id = _currentSupplierId,
                .SupplierName = txtSupplierName.Text,
                .Phone = txtSupplierPhone.Text,
                .Gstin = txtSupplierGstin.Text,
                .Email = txtSupplierEmail.Text,
                .DrugLicenseNumber = txtSupplierLicense.Text,
                .Address = txtSupplierAddress.Text,
                .OutstandingBalance = nudSupplierOutstanding.Value
            }
        End Function

        Private Sub UpdateSupplierActionState()
            btnDeleteSupplier.Enabled = _currentSupplierId > 0
        End Sub

        Private Async Sub btnNewSupplier_Click(sender As Object, e As EventArgs)
            PrepareNewSupplier()
            ShowStatus("Ready to add a new supplier.", False)
            Await Task.CompletedTask
        End Sub

        Private Async Sub btnSaveSupplier_Click(sender As Object, e As EventArgs)
            Dim supplier As SupplierRecord = ReadSupplierFromForm()
            SetBusy(True, "Saving supplier...")
            Dim result As EntityOperationResult = Await _supplierService.SaveAsync(supplier)
            SetBusy(False)
            ShowStatus(result.Message, Not result.IsSuccess)

            If result.IsSuccess Then
                Await LoadLookupsAsync(result.EntityId)
            End If
        End Sub

        Private Async Sub btnDeleteSupplier_Click(sender As Object, e As EventArgs)
            Dim supplier As SupplierRecord = ReadSupplierFromForm()
            If supplier.Id <= 0 Then
                ShowStatus("Select a supplier to delete.", True)
                Return
            End If

            Dim confirmation As DialogResult = MessageBox.Show(
                $"Delete '{supplier.SupplierName}' from the supplier master?",
                "Confirm Delete",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question)

            If confirmation <> DialogResult.Yes Then
                Return
            End If

            SetBusy(True, "Deleting supplier...")
            Dim result As EntityOperationResult = Await _supplierService.DeleteAsync(supplier)
            SetBusy(False)
            ShowStatus(result.Message, Not result.IsSuccess)

            If result.IsSuccess Then
                PrepareNewSupplier()
                Await LoadLookupsAsync()
            End If
        End Sub

        Private Sub cboSupplier_SelectedIndexChanged(sender As Object, e As EventArgs)
            If _isBusy Then
                Return
            End If

            BindSupplier(TryCast(cboSupplier.SelectedItem, SupplierRecord))
        End Sub

        Private Sub cboProduct_SelectedIndexChanged(sender As Object, e As EventArgs)
            If _isBusy Then
                Return
            End If

            Dim product As ProductRecord = TryCast(cboProduct.SelectedItem, ProductRecord)
            If product Is Nothing Then
                txtPacking.Clear()
                txtHsnCode.Clear()
                Return
            End If

            Dim line As PurchaseLineItem = _purchaseService.CreateLineFromProduct(product)
            txtPacking.Text = line.Packing
            txtHsnCode.Text = line.HsnCode
            txtBatchNumber.Text = line.BatchNumber
            dtpExpiryDate.Value = If(line.ExpiryDate = DateTime.MinValue, DateTime.Today, line.ExpiryDate)
            nudQuantity.Value = 1D
            nudFreeQuantity.Value = 0D
            nudPTR.Value = Math.Min(nudPTR.Maximum, line.PTR)
            nudPTS.Value = Math.Min(nudPTS.Maximum, line.PTS)
            nudMRP.Value = Math.Min(nudMRP.Maximum, line.MRP)
            nudGST.Value = Math.Min(nudGST.Maximum, line.GstPercentage)
        End Sub

        Private Sub btnAddItem_Click(sender As Object, e As EventArgs)
            Dim product As ProductRecord = TryCast(cboProduct.SelectedItem, ProductRecord)
            If product Is Nothing Then
                ShowStatus("Select a product to add.", True)
                Return
            End If

            Dim batchNumber As String = txtBatchNumber.Text.Trim().ToUpperInvariant()
            If batchNumber = String.Empty Then
                ShowStatus("Enter a batch number for the purchase line.", True)
                txtBatchNumber.Focus()
                Return
            End If

            Dim existingLine As PurchaseLineItem =
                _purchaseItems.FirstOrDefault(
                    Function(line) String.Equals(line.ProductName, product.ProductName, StringComparison.OrdinalIgnoreCase) AndAlso
                                   String.Equals(line.BatchNumber, batchNumber, StringComparison.OrdinalIgnoreCase))

            If existingLine IsNot Nothing Then
                existingLine.Quantity += Decimal.ToInt32(nudQuantity.Value)
                existingLine.FreeQuantity += Decimal.ToInt32(nudFreeQuantity.Value)
                existingLine.ExpiryDate = dtpExpiryDate.Value.Date
                existingLine.PTR = nudPTR.Value
                existingLine.PTS = nudPTS.Value
                existingLine.MRP = nudMRP.Value
                existingLine.GstPercentage = nudGST.Value
                _purchaseService.RecalculateLine(existingLine)
                dgvItems.Refresh()
            Else
                Dim line As PurchaseLineItem = _purchaseService.CreateLineFromProduct(product)
                line.BatchNumber = batchNumber
                line.ExpiryDate = dtpExpiryDate.Value.Date
                line.Quantity = Decimal.ToInt32(nudQuantity.Value)
                line.FreeQuantity = Decimal.ToInt32(nudFreeQuantity.Value)
                line.PTR = nudPTR.Value
                line.PTS = nudPTS.Value
                line.MRP = nudMRP.Value
                line.GstPercentage = nudGST.Value
                _purchaseService.RecalculateLine(line)
                _purchaseItems.Add(line)
            End If

            ReindexLines()
            RefreshTotals()
            ResetItemEntry()
            ShowStatus($"Added '{product.ProductName}' to the purchase.", False)
        End Sub

        Private Sub ResetItemEntry()
            nudQuantity.Value = 1D
            nudFreeQuantity.Value = 0D
            nudPTR.Value = 0D
            nudPTS.Value = 0D
            nudMRP.Value = 0D
            nudGST.Value = 0D
            txtBatchNumber.Clear()
            dtpExpiryDate.Value = DateTime.Today
        End Sub

        Private Sub btnRemoveItem_Click(sender As Object, e As EventArgs)
            If dgvItems.CurrentRow Is Nothing Then
                ShowStatus("Select a line to remove.", True)
                Return
            End If

            Dim line As PurchaseLineItem = TryCast(dgvItems.CurrentRow.DataBoundItem, PurchaseLineItem)
            If line Is Nothing Then
                Return
            End If

            _purchaseItems.Remove(line)
            ReindexLines()
            RefreshTotals()
            ShowStatus("Purchase line removed.", False)
        End Sub

        Private Async Sub btnNewPurchase_Click(sender As Object, e As EventArgs)
            Await PrepareNewPurchaseAsync()
        End Sub

        Private Async Sub btnSavePurchase_Click(sender As Object, e As EventArgs)
            Dim supplier As SupplierRecord = TryCast(cboSupplier.SelectedItem, SupplierRecord)
            Dim draft As New PurchaseDraft With {
                .PurchaseNumber = txtPurchaseNumber.Text,
                .PurchaseDate = dtpPurchaseDate.Value.Date,
                .SupplierId = If(supplier Is Nothing, 0, supplier.Id),
                .SupplierName = If(supplier Is Nothing, String.Empty, supplier.SupplierName),
                .SupplierInvoiceNumber = txtSupplierInvoiceNumber.Text,
                .SupplierInvoiceDate = ReadOptionalDate(dtpSupplierInvoiceDate),
                .PurchaseOrderNumber = txtPurchaseOrderNumber.Text,
                .PurchaseOrderDate = ReadOptionalDate(dtpPurchaseOrderDate),
                .PlaceOfSupply = txtPlaceOfSupply.Text,
                .CaseCount = Decimal.ToInt32(nudCaseCount.Value),
                .TransportName = txtTransportName.Text,
                .EwayBillNumber = txtEwayBillNumber.Text,
                .Notes = txtNotes.Text
            }

            For Each line As PurchaseLineItem In _purchaseItems
                draft.Items.Add(line)
            Next

            draft.Summary = _purchaseService.CalculateTotals(draft.Items)

            SetBusy(True, "Saving purchase...")
            Dim result As PurchaseSaveResult = Await _purchaseService.SavePurchaseAsync(draft, If(SessionManager.CurrentUser Is Nothing, 0, SessionManager.CurrentUser.Id))
            SetBusy(False)

            If result.IsSuccess Then
                _lastSavedPurchaseId = result.PurchaseId
                UpdatePrintActionState()
                Await LoadLookupsAsync(draft.SupplierId)
                Await PrepareNewPurchaseAsync()
                ShowStatus($"{result.Message} Preview or print is ready.", False)
            Else
                ShowStatus(result.Message, True)
            End If
        End Sub

        Private Sub btnPrintPreview_Click(sender As Object, e As EventArgs)
            If _lastSavedPurchaseId <= 0 Then
                ShowStatus("Save a purchase before previewing the print layout.", True)
                Return
            End If

            Try
                _purchasePrintService.ShowPrintPreview(_lastSavedPurchaseId)
                ShowStatus("Purchase print preview opened.", False)
            Catch ex As Exception
                AppLogger.Error($"Purchase preview failed for Id {_lastSavedPurchaseId}.", ex)
                ShowStatus("Purchase preview could not be opened.", True)
            End Try
        End Sub

        Private Sub btnPrint_Click(sender As Object, e As EventArgs)
            If _lastSavedPurchaseId <= 0 Then
                ShowStatus("Save a purchase before printing.", True)
                Return
            End If

            Try
                _purchasePrintService.PrintPurchase(_lastSavedPurchaseId)
                ShowStatus("Purchase sent to printer.", False)
            Catch ex As Exception
                AppLogger.Error($"Purchase print failed for Id {_lastSavedPurchaseId}.", ex)
                ShowStatus("Purchase could not be printed.", True)
            End Try
        End Sub

        Private Sub btnClose_Click(sender As Object, e As EventArgs)
            Close()
        End Sub

        Private Sub dgvItems_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs)
            If e.RowIndex < 0 OrElse e.RowIndex >= _purchaseItems.Count Then
                Return
            End If

            Dim line As PurchaseLineItem = _purchaseItems(e.RowIndex)
            line.BatchNumber = If(line.BatchNumber, String.Empty).Trim().ToUpperInvariant()
            _purchaseService.RecalculateLine(line)
            dgvItems.Refresh()
            RefreshTotals()
        End Sub

        Private Sub dgvItems_DataError(sender As Object, e As DataGridViewDataErrorEventArgs)
            e.Cancel = True
            ShowStatus("Enter a valid numeric value in the purchase grid.", True)
        End Sub

        Private Async Sub dtpPurchaseDate_ValueChanged(sender As Object, e As EventArgs)
            If _purchaseItems.Count = 0 AndAlso Not _isBusy Then
                txtPurchaseNumber.Text = Await _purchaseService.GenerateNextPurchaseNumberAsync(dtpPurchaseDate.Value.Date)
            End If
        End Sub

        Private Sub ReindexLines()
            For index As Integer = 0 To _purchaseItems.Count - 1
                _purchaseItems(index).LineNumber = index + 1
            Next
            dgvItems.Refresh()
        End Sub

        Private Sub RefreshTotals()
            Dim summary As PurchaseTotalsSummary = _purchaseService.CalculateTotals(_purchaseItems.ToList())
            lblSubTotalValue.Text = summary.SubTotal.ToString("N2")
            lblGstValue.Text = summary.GstAmount.ToString("N2")
            lblRoundOffValue.Text = summary.RoundOffAmount.ToString("N2")
            lblNetAmountValue.Text = summary.NetAmount.ToString("N2")
            txtTaxSummary.Text = summary.TaxSummaryText
        End Sub

        Private Function ReadOptionalDate(control As DateTimePicker) As DateTime?
            If Not control.Checked Then
                Return Nothing
            End If

            Return control.Value.Date
        End Function

        Private Sub ResetOptionalDatePicker(control As DateTimePicker)
            control.Value = DateTime.Today
            control.Checked = False
        End Sub

        Private Sub SetBusy(isBusy As Boolean, Optional message As String = "")
            _isBusy = isBusy

            cboSupplier.Enabled = Not isBusy
            txtPurchaseNumber.Enabled = Not isBusy
            dtpPurchaseDate.Enabled = Not isBusy
            txtSupplierInvoiceNumber.Enabled = Not isBusy
            dtpSupplierInvoiceDate.Enabled = Not isBusy
            txtPurchaseOrderNumber.Enabled = Not isBusy
            dtpPurchaseOrderDate.Enabled = Not isBusy
            txtPlaceOfSupply.Enabled = Not isBusy
            nudCaseCount.Enabled = Not isBusy
            txtTransportName.Enabled = Not isBusy
            txtEwayBillNumber.Enabled = Not isBusy
            txtNotes.Enabled = Not isBusy

            txtSupplierName.Enabled = Not isBusy
            txtSupplierPhone.Enabled = Not isBusy
            txtSupplierGstin.Enabled = Not isBusy
            txtSupplierEmail.Enabled = Not isBusy
            txtSupplierLicense.Enabled = Not isBusy
            txtSupplierAddress.Enabled = Not isBusy
            nudSupplierOutstanding.Enabled = Not isBusy
            btnNewSupplier.Enabled = Not isBusy
            btnSaveSupplier.Enabled = Not isBusy
            btnDeleteSupplier.Enabled = Not isBusy AndAlso _currentSupplierId > 0

            cboProduct.Enabled = Not isBusy
            txtPacking.Enabled = Not isBusy
            txtHsnCode.Enabled = Not isBusy
            txtBatchNumber.Enabled = Not isBusy
            dtpExpiryDate.Enabled = Not isBusy
            nudQuantity.Enabled = Not isBusy
            nudFreeQuantity.Enabled = Not isBusy
            nudPTR.Enabled = Not isBusy
            nudPTS.Enabled = Not isBusy
            nudMRP.Enabled = Not isBusy
            nudGST.Enabled = Not isBusy
            btnAddItem.Enabled = Not isBusy
            btnRemoveItem.Enabled = Not isBusy
            dgvItems.Enabled = Not isBusy

            btnNewPurchase.Enabled = Not isBusy
            btnSavePurchase.Enabled = Not isBusy
            btnPrintPreview.Enabled = Not isBusy AndAlso _lastSavedPurchaseId > 0
            btnPrint.Enabled = Not isBusy AndAlso _lastSavedPurchaseId > 0
            btnClose.Enabled = Not isBusy

            If isBusy Then
                lblStatus.ForeColor = ThemePalette.TextMuted
                lblStatus.Text = message
            End If
        End Sub

        Private Sub UpdatePrintActionState()
            btnPrintPreview.Enabled = Not _isBusy AndAlso _lastSavedPurchaseId > 0
            btnPrint.Enabled = Not _isBusy AndAlso _lastSavedPurchaseId > 0
        End Sub

        Private Sub ShowStatus(message As String, isError As Boolean)
            lblStatus.ForeColor = If(isError, ThemePalette.DangerRed, ThemePalette.AccentGreen)
            lblStatus.Text = message
        End Sub

        Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
            Select Case keyData
                Case Keys.F2
                    cboSupplier.Focus()
                    Return True
                Case Keys.F3
                    cboProduct.Focus()
                    Return True
                Case Keys.Control Or Keys.S
                    btnSavePurchase.PerformClick()
                    Return True
                Case Keys.Control Or Keys.N
                    btnNewPurchase.PerformClick()
                    Return True
                Case Keys.Control Or Keys.P
                    btnPrintPreview.PerformClick()
                    Return True
                Case Keys.Control Or Keys.Shift Or Keys.N
                    btnNewSupplier.PerformClick()
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
