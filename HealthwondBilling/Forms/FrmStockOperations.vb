Imports HealthwondBilling.Models
Imports HealthwondBilling.Services
Imports HealthwondBilling.Utilities
Imports System.ComponentModel
Imports System.Linq

Namespace Forms

    Public Class FrmStockOperations
        Inherits Form

        Private ReadOnly _stockOperationService As StockOperationService
        Private ReadOnly _purchaseRows As New BindingList(Of PurchaseHistoryLookupRow)()
        Private ReadOnly _purchaseReturnItems As New BindingList(Of PurchaseReturnLineItem)()
        Private ReadOnly _stockAdjustmentItems As New BindingList(Of StockAdjustmentLineItem)()

        Private _products As New List(Of ProductRecord)()
        Private _isBusy As Boolean

        Private ReadOnly tabs As New TabControl()
        Private ReadOnly lblStatus As New Label()

        Private ReadOnly txtPurchaseSearch As New TextBox()
        Private ReadOnly dtpPurchaseFrom As New DateTimePicker()
        Private ReadOnly dtpPurchaseTo As New DateTimePicker()
        Private ReadOnly btnRefreshPurchases As New Button()
        Private ReadOnly dgvPurchases As New DataGridView()
        Private ReadOnly txtReturnNumber As New TextBox()
        Private ReadOnly dtpReturnDate As New DateTimePicker()
        Private ReadOnly txtReturnNotes As New TextBox()
        Private ReadOnly btnSaveReturn As New Button()
        Private ReadOnly btnResetReturn As New Button()
        Private ReadOnly lblSelectedPurchase As New Label()
        Private ReadOnly lblReturnSubTotalValue As New Label()
        Private ReadOnly lblReturnGstValue As New Label()
        Private ReadOnly lblReturnRoundOffValue As New Label()
        Private ReadOnly lblReturnNetValue As New Label()
        Private ReadOnly txtReturnTaxSummary As New TextBox()
        Private ReadOnly dgvReturnItems As New DataGridView()

        Private ReadOnly txtAdjustmentNumber As New TextBox()
        Private ReadOnly dtpAdjustmentDate As New DateTimePicker()
        Private ReadOnly txtAdjustmentNotes As New TextBox()
        Private ReadOnly cboAdjustmentProduct As New ComboBox()
        Private ReadOnly cboAdjustmentMode As New ComboBox()
        Private ReadOnly nudAdjustmentQuantity As New NumericUpDown()
        Private ReadOnly txtAdjustmentLineRemarks As New TextBox()
        Private ReadOnly btnAddAdjustmentLine As New Button()
        Private ReadOnly btnRemoveAdjustmentLine As New Button()
        Private ReadOnly btnSaveAdjustment As New Button()
        Private ReadOnly btnNewAdjustment As New Button()
        Private ReadOnly dgvAdjustmentItems As New DataGridView()
        Private ReadOnly lblAdjustmentSummary As New Label()

        Public Sub New(stockOperationService As StockOperationService)
            _stockOperationService = stockOperationService

            Text = "Healthwond Billing System - Stock Operations"
            StartPosition = FormStartPosition.CenterParent
            WindowState = FormWindowState.Maximized
            MinimumSize = New Size(1500, 920)
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
                .RowCount = 3,
                .Padding = New Padding(22),
                .BackColor = ThemePalette.AppBackground
            }
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 78))
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 36))

            root.Controls.Add(BuildHeaderPanel(), 0, 0)
            root.Controls.Add(BuildTabsPanel(), 0, 1)

            lblStatus.Dock = DockStyle.Fill
            lblStatus.Font = New Font("Segoe UI Semibold", 9.5F, FontStyle.Bold)
            lblStatus.TextAlign = ContentAlignment.MiddleLeft
            root.Controls.Add(lblStatus, 0, 2)

            Controls.Add(root)
        End Sub

        Private Function BuildHeaderPanel() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = ThemePalette.AppBackground}

            Dim title As New Label With {
                .Dock = DockStyle.Top,
                .Height = 40,
                .Text = "Stock Operations",
                .Font = New Font("Segoe UI Semibold", 24.0F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary
            }

            Dim subtitle As New Label With {
                .Dock = DockStyle.Fill,
                .Text = "Process supplier purchase returns and manual stock adjustments with auditable ledger posting and balance corrections.",
                .Font = New Font("Segoe UI", 10.5F, FontStyle.Regular),
                .ForeColor = ThemePalette.TextMuted
            }

            panel.Controls.Add(subtitle)
            panel.Controls.Add(title)
            Return panel
        End Function

        Private Function BuildTabsPanel() As Control
            tabs.Dock = DockStyle.Fill
            tabs.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
            tabs.TabPages.Add(BuildPurchaseReturnTab())
            tabs.TabPages.Add(BuildStockAdjustmentTab())
            Return tabs
        End Function

        Private Function BuildPurchaseReturnTab() As TabPage
            Dim page As New TabPage("Purchase Return") With {.BackColor = ThemePalette.AppBackground}

            Dim root As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 4,
                .Padding = New Padding(8),
                .BackColor = ThemePalette.AppBackground
            }
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 116))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 260))
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 200))

            root.Controls.Add(BuildPurchaseReturnFilterPanel(), 0, 0)
            root.Controls.Add(BuildPurchaseGridPanel(), 0, 1)
            root.Controls.Add(BuildReturnItemsPanel(), 0, 2)
            root.Controls.Add(BuildPurchaseReturnFooterPanel(), 0, 3)

            page.Controls.Add(root)
            Return page
        End Function

        Private Function BuildPurchaseReturnFilterPanel() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)

            Dim layout As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 4
            }
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 48.0F))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 180))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 180))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 120))

            layout.Controls.Add(CreateInputHost("Search Purchase", txtPurchaseSearch), 0, 0)
            layout.Controls.Add(CreateInputHost("From Date", dtpPurchaseFrom), 1, 0)
            layout.Controls.Add(CreateInputHost("To Date", dtpPurchaseTo), 2, 0)

            UiStyler.StylePrimaryButton(btnRefreshPurchases)
            btnRefreshPurchases.Text = "Refresh"
            btnRefreshPurchases.Width = 100
            btnRefreshPurchases.Margin = New Padding(0, 24, 0, 0)
            layout.Controls.Add(btnRefreshPurchases, 3, 0)

            panel.Controls.Add(layout)
            Return panel
        End Function

        Private Function BuildPurchaseGridPanel() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)

            Dim title As New Label With {
                .Dock = DockStyle.Top,
                .Height = 28,
                .Text = "Purchase search",
                .Font = New Font("Segoe UI Semibold", 13.0F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary
            }

            dgvPurchases.Dock = DockStyle.Fill
            panel.Controls.Add(dgvPurchases)
            panel.Controls.Add(title)
            Return panel
        End Function

        Private Function BuildReturnItemsPanel() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)

            Dim root As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 3
            }
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 28))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 62))
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            lblSelectedPurchase.Dock = DockStyle.Fill
            lblSelectedPurchase.Font = New Font("Segoe UI Semibold", 12.0F, FontStyle.Bold)
            lblSelectedPurchase.ForeColor = ThemePalette.TextPrimary
            lblSelectedPurchase.Text = "Select a purchase to load returnable lines."

            Dim headerLayout As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 4
            }
            headerLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 30.0F))
            headerLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 190))
            headerLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 40.0F))
            headerLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 220))

            headerLayout.Controls.Add(CreateInputHost("Return Number", txtReturnNumber), 0, 0)
            headerLayout.Controls.Add(CreateInputHost("Return Date", dtpReturnDate), 1, 0)

            txtReturnNotes.BorderStyle = BorderStyle.FixedSingle
            UiStyler.StyleInput(txtReturnNotes)
            headerLayout.Controls.Add(CreateInputHost("Return Notes", txtReturnNotes), 2, 0)

            Dim buttonFlow As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.LeftToRight,
                .WrapContents = False,
                .Padding = New Padding(0, 22, 0, 0),
                .BackColor = Color.Transparent
            }

            UiStyler.StyleSecondaryButton(btnResetReturn)
            btnResetReturn.Text = "Reset"
            btnResetReturn.Width = 90

            UiStyler.StylePrimaryButton(btnSaveReturn)
            btnSaveReturn.Text = "Save Return"
            btnSaveReturn.Width = 110

            buttonFlow.Controls.Add(btnResetReturn)
            buttonFlow.Controls.Add(btnSaveReturn)
            headerLayout.Controls.Add(buttonFlow, 3, 0)

            dgvReturnItems.Dock = DockStyle.Fill

            root.Controls.Add(lblSelectedPurchase, 0, 0)
            root.Controls.Add(headerLayout, 0, 1)
            root.Controls.Add(dgvReturnItems, 0, 2)

            panel.Controls.Add(root)
            Return panel
        End Function

        Private Function BuildPurchaseReturnFooterPanel() As Control
            Dim split As New SplitContainer With {
                .Dock = DockStyle.Fill,
                .SplitterDistance = 620,
                .BackColor = ThemePalette.AppBackground
            }

            Dim taxPanel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(taxPanel)
            Dim taxTitle As New Label With {
                .Dock = DockStyle.Top,
                .Height = 26,
                .Text = "Return tax summary",
                .Font = New Font("Segoe UI Semibold", 12.0F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary
            }
            txtReturnTaxSummary.Dock = DockStyle.Fill
            txtReturnTaxSummary.Multiline = True
            txtReturnTaxSummary.ReadOnly = True
            txtReturnTaxSummary.BorderStyle = BorderStyle.FixedSingle
            txtReturnTaxSummary.BackColor = Color.White
            txtReturnTaxSummary.Font = New Font("Consolas", 10.0F, FontStyle.Regular)
            taxPanel.Controls.Add(txtReturnTaxSummary)
            taxPanel.Controls.Add(taxTitle)

            Dim totalsPanel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(totalsPanel)
            Dim totalsTable As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2
            }
            totalsTable.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 52.0F))
            totalsTable.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 48.0F))
            For index As Integer = 0 To 3
                totalsTable.RowStyles.Add(New RowStyle(SizeType.Absolute, 30))
            Next

            AddTotalRow(totalsTable, 0, "Subtotal", lblReturnSubTotalValue)
            AddTotalRow(totalsTable, 1, "GST", lblReturnGstValue)
            AddTotalRow(totalsTable, 2, "Round Off", lblReturnRoundOffValue)
            AddTotalRow(totalsTable, 3, "Net Amount", lblReturnNetValue, True)
            totalsPanel.Controls.Add(totalsTable)

            split.Panel1.Controls.Add(taxPanel)
            split.Panel2.Controls.Add(totalsPanel)
            Return split
        End Function

        Private Function BuildStockAdjustmentTab() As TabPage
            Dim page As New TabPage("Stock Adjustment") With {.BackColor = ThemePalette.AppBackground}

            Dim root As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 4,
                .Padding = New Padding(8),
                .BackColor = ThemePalette.AppBackground
            }
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 120))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 122))
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 84))

            root.Controls.Add(BuildAdjustmentHeaderPanel(), 0, 0)
            root.Controls.Add(BuildAdjustmentEntryPanel(), 0, 1)
            root.Controls.Add(BuildAdjustmentGridPanel(), 0, 2)
            root.Controls.Add(BuildAdjustmentFooterPanel(), 0, 3)

            page.Controls.Add(root)
            Return page
        End Function

        Private Function BuildAdjustmentHeaderPanel() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)

            Dim layout As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 4
            }
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 24.0F))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 190))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 46.0F))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 220))

            layout.Controls.Add(CreateInputHost("Adjustment Number", txtAdjustmentNumber), 0, 0)
            layout.Controls.Add(CreateInputHost("Adjustment Date", dtpAdjustmentDate), 1, 0)

            txtAdjustmentNotes.BorderStyle = BorderStyle.FixedSingle
            UiStyler.StyleInput(txtAdjustmentNotes)
            layout.Controls.Add(CreateInputHost("Adjustment Notes", txtAdjustmentNotes), 2, 0)

            Dim buttonFlow As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.LeftToRight,
                .WrapContents = False,
                .Padding = New Padding(0, 22, 0, 0),
                .BackColor = Color.Transparent
            }

            UiStyler.StyleSecondaryButton(btnNewAdjustment)
            btnNewAdjustment.Text = "New"
            btnNewAdjustment.Width = 90

            UiStyler.StylePrimaryButton(btnSaveAdjustment)
            btnSaveAdjustment.Text = "Save Adjustment"
            btnSaveAdjustment.Width = 130

            buttonFlow.Controls.Add(btnNewAdjustment)
            buttonFlow.Controls.Add(btnSaveAdjustment)
            layout.Controls.Add(buttonFlow, 3, 0)

            panel.Controls.Add(layout)
            Return panel
        End Function

        Private Function BuildAdjustmentEntryPanel() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)

            Dim layout As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 6
            }
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 40.0F))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 160))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 130))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 32.0F))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 110))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 230))

            layout.Controls.Add(CreateInputHost("Product", cboAdjustmentProduct), 0, 0)
            layout.Controls.Add(CreateInputHost("Mode", cboAdjustmentMode), 1, 0)
            layout.Controls.Add(CreateInputHost("Quantity", nudAdjustmentQuantity), 2, 0)

            txtAdjustmentLineRemarks.BorderStyle = BorderStyle.FixedSingle
            UiStyler.StyleInput(txtAdjustmentLineRemarks)
            layout.Controls.Add(CreateInputHost("Line Remarks", txtAdjustmentLineRemarks), 3, 0)

            Dim buttonFlow As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.LeftToRight,
                .WrapContents = False,
                .Padding = New Padding(0, 22, 0, 0),
                .BackColor = Color.Transparent
            }

            UiStyler.StylePrimaryButton(btnAddAdjustmentLine)
            btnAddAdjustmentLine.Text = "Add Line"
            btnAddAdjustmentLine.Width = 100

            UiStyler.StyleDangerButton(btnRemoveAdjustmentLine)
            btnRemoveAdjustmentLine.Text = "Remove"
            btnRemoveAdjustmentLine.Width = 100

            buttonFlow.Controls.Add(btnAddAdjustmentLine)
            buttonFlow.Controls.Add(btnRemoveAdjustmentLine)
            layout.Controls.Add(buttonFlow, 5, 0)

            panel.Controls.Add(layout)
            Return panel
        End Function

        Private Function BuildAdjustmentGridPanel() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)
            dgvAdjustmentItems.Dock = DockStyle.Fill
            panel.Controls.Add(dgvAdjustmentItems)
            Return panel
        End Function

        Private Function BuildAdjustmentFooterPanel() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)
            lblAdjustmentSummary.Dock = DockStyle.Fill
            lblAdjustmentSummary.Font = New Font("Segoe UI Semibold", 11.0F, FontStyle.Bold)
            lblAdjustmentSummary.ForeColor = ThemePalette.TextPrimary
            lblAdjustmentSummary.TextAlign = ContentAlignment.MiddleLeft
            panel.Controls.Add(lblAdjustmentSummary)
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

        Private Sub ConfigureControls()
            txtPurchaseSearch.BorderStyle = BorderStyle.FixedSingle
            UiStyler.StyleInput(txtPurchaseSearch)

            ConfigureDatePicker(dtpPurchaseFrom, New DateTime(DateTime.Today.Year, DateTime.Today.Month, 1))
            ConfigureDatePicker(dtpPurchaseTo, DateTime.Today)
            ConfigureDatePicker(dtpReturnDate, DateTime.Today)
            ConfigureDatePicker(dtpAdjustmentDate, DateTime.Today)

            txtReturnNumber.ReadOnly = True
            txtReturnNumber.BorderStyle = BorderStyle.FixedSingle
            UiStyler.StyleInput(txtReturnNumber)

            txtReturnNotes.BorderStyle = BorderStyle.FixedSingle
            UiStyler.StyleInput(txtReturnNotes)

            txtAdjustmentNumber.ReadOnly = True
            txtAdjustmentNumber.BorderStyle = BorderStyle.FixedSingle
            UiStyler.StyleInput(txtAdjustmentNumber)

            txtAdjustmentNotes.BorderStyle = BorderStyle.FixedSingle
            UiStyler.StyleInput(txtAdjustmentNotes)

            ConfigureCombo(cboAdjustmentProduct)
            ConfigureCombo(cboAdjustmentMode)
            cboAdjustmentMode.DropDownStyle = ComboBoxStyle.DropDownList
            cboAdjustmentMode.Items.AddRange(New Object() {"Increase", "Decrease"})
            cboAdjustmentMode.SelectedIndex = 0

            ConfigureNumeric(nudAdjustmentQuantity, 0, 1000000D, 1D)
            nudAdjustmentQuantity.Value = 1D

            txtAdjustmentLineRemarks.BorderStyle = BorderStyle.FixedSingle
            UiStyler.StyleInput(txtAdjustmentLineRemarks)

            txtReturnTaxSummary.ReadOnly = True
            txtReturnTaxSummary.BackColor = Color.White

            lblAdjustmentSummary.Text = "No adjustment lines yet."
            ResetReturnTotals()
        End Sub

        Private Sub ConfigureDatePicker(control As DateTimePicker, defaultValue As DateTime)
            control.Format = DateTimePickerFormat.Custom
            control.CustomFormat = "dd-MMM-yyyy"
            control.Value = defaultValue
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

        Private Sub ConfigureGrids()
            ConfigureGrid(dgvPurchases)
            dgvPurchases.AutoGenerateColumns = False
            dgvPurchases.Columns.Add(CreateDateColumn("PurchaseDate", "Date", 82))
            dgvPurchases.Columns.Add(CreateTextColumn("PurchaseNumber", "Purchase No", 100))
            dgvPurchases.Columns.Add(CreateTextColumn("SupplierName", "Supplier", 150))
            dgvPurchases.Columns.Add(CreateTextColumn("SupplierInvoiceNumber", "Supplier Invoice", 110))
            dgvPurchases.Columns.Add(CreateIntegerColumn("LineCount", "Lines", 60))
            dgvPurchases.Columns.Add(CreateIntegerColumn("TotalUnits", "Units", 65))
            dgvPurchases.Columns.Add(CreateDecimalColumn("NetAmount", "Net Amount", 90))
            dgvPurchases.Columns.Add(CreateTextColumn("Notes", "Notes", 130))
            dgvPurchases.DataSource = _purchaseRows

            ConfigureGrid(dgvReturnItems)
            dgvReturnItems.AutoGenerateColumns = False
            dgvReturnItems.ReadOnly = False
            dgvReturnItems.Columns.Add(CreateIntegerColumn("LineNumber", "#", 42))
            dgvReturnItems.Columns.Add(CreateTextColumn("ProductName", "Product", 150))
            dgvReturnItems.Columns.Add(CreateTextColumn("BatchNumber", "Batch", 85))
            dgvReturnItems.Columns.Add(CreateDateColumn("ExpiryDate", "Expiry", 80))
            dgvReturnItems.Columns.Add(CreateIntegerColumn("PurchasedQuantity", "Pur Qty", 65))
            dgvReturnItems.Columns.Add(CreateIntegerColumn("PurchasedFreeQuantity", "Pur Free", 65))
            dgvReturnItems.Columns.Add(CreateIntegerColumn("RemainingQuantity", "Rem Qty", 65))
            dgvReturnItems.Columns.Add(CreateIntegerColumn("RemainingFreeQuantity", "Rem Free", 65))
            dgvReturnItems.Columns.Add(CreateIntegerColumn("CurrentStock", "Stock", 60))
            dgvReturnItems.Columns.Add(CreateEditableIntegerColumn("ReturnQuantity", "Return Qty", 70))
            dgvReturnItems.Columns.Add(CreateEditableIntegerColumn("ReturnFreeQuantity", "Return Free", 72))
            dgvReturnItems.Columns.Add(CreateDecimalColumn("PTR", "PTR", 70))
            dgvReturnItems.Columns.Add(CreateDecimalColumn("GstPercentage", "GST %", 65))
            dgvReturnItems.Columns.Add(CreateDecimalColumn("LineTotal", "Total", 82))
            dgvReturnItems.DataSource = _purchaseReturnItems

            ConfigureGrid(dgvAdjustmentItems)
            dgvAdjustmentItems.AutoGenerateColumns = False
            dgvAdjustmentItems.Columns.Add(CreateIntegerColumn("LineNumber", "#", 42))
            dgvAdjustmentItems.Columns.Add(CreateTextColumn("ProductName", "Product", 160))
            dgvAdjustmentItems.Columns.Add(CreateTextColumn("BatchNumber", "Batch", 85))
            dgvAdjustmentItems.Columns.Add(CreateIntegerColumn("CurrentStock", "Current", 70))
            dgvAdjustmentItems.Columns.Add(CreateTextColumn("AdjustmentMode", "Mode", 80))
            dgvAdjustmentItems.Columns.Add(CreateIntegerColumn("Quantity", "Qty", 60))
            dgvAdjustmentItems.Columns.Add(CreateIntegerColumn("ResultingStock", "Result", 70))
            dgvAdjustmentItems.Columns.Add(CreateDecimalColumn("UnitCost", "Unit Cost", 85))
            dgvAdjustmentItems.Columns.Add(CreateTextColumn("Remarks", "Remarks", 150))
            dgvAdjustmentItems.DataSource = _stockAdjustmentItems
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
                .Name = propertyName,
                .ReadOnly = True
            }
        End Function

        Private Function CreateIntegerColumn(propertyName As String, headerText As String, fillWeight As Single) As DataGridViewColumn
            Return New DataGridViewTextBoxColumn With {
                .DataPropertyName = propertyName,
                .HeaderText = headerText,
                .FillWeight = fillWeight,
                .Name = propertyName,
                .ReadOnly = True,
                .DefaultCellStyle = New DataGridViewCellStyle With {.Format = "N0"}
            }
        End Function

        Private Function CreateEditableIntegerColumn(propertyName As String, headerText As String, fillWeight As Single) As DataGridViewColumn
            Return New DataGridViewTextBoxColumn With {
                .DataPropertyName = propertyName,
                .HeaderText = headerText,
                .FillWeight = fillWeight,
                .Name = propertyName,
                .ReadOnly = False,
                .DefaultCellStyle = New DataGridViewCellStyle With {.Format = "N0"}
            }
        End Function

        Private Function CreateDecimalColumn(propertyName As String, headerText As String, fillWeight As Single) As DataGridViewColumn
            Return New DataGridViewTextBoxColumn With {
                .DataPropertyName = propertyName,
                .HeaderText = headerText,
                .FillWeight = fillWeight,
                .Name = propertyName,
                .ReadOnly = True,
                .DefaultCellStyle = New DataGridViewCellStyle With {.Format = "N2"}
            }
        End Function

        Private Function CreateDateColumn(propertyName As String, headerText As String, fillWeight As Single) As DataGridViewColumn
            Return New DataGridViewTextBoxColumn With {
                .DataPropertyName = propertyName,
                .HeaderText = headerText,
                .FillWeight = fillWeight,
                .Name = propertyName,
                .ReadOnly = True,
                .DefaultCellStyle = New DataGridViewCellStyle With {.Format = "dd-MMM-yyyy"}
            }
        End Function

        Private Sub WireEvents()
            AddHandler Load, AddressOf FrmStockOperations_Load
            AddHandler btnRefreshPurchases.Click, AddressOf btnRefreshPurchases_Click
            AddHandler dgvPurchases.SelectionChanged, AddressOf dgvPurchases_SelectionChanged
            AddHandler dgvReturnItems.CellEndEdit, AddressOf dgvReturnItems_CellEndEdit
            AddHandler dgvReturnItems.DataError, AddressOf dgvReturnItems_DataError
            AddHandler btnSaveReturn.Click, AddressOf btnSaveReturn_Click
            AddHandler btnResetReturn.Click, AddressOf btnResetReturn_Click
            AddHandler dtpReturnDate.ValueChanged, AddressOf dtpReturnDate_ValueChanged
            AddHandler txtPurchaseSearch.KeyDown, AddressOf txtPurchaseSearch_KeyDown
            AddHandler cboAdjustmentProduct.SelectedIndexChanged, AddressOf cboAdjustmentProduct_SelectedIndexChanged
            AddHandler btnAddAdjustmentLine.Click, AddressOf btnAddAdjustmentLine_Click
            AddHandler btnRemoveAdjustmentLine.Click, AddressOf btnRemoveAdjustmentLine_Click
            AddHandler btnSaveAdjustment.Click, AddressOf btnSaveAdjustment_Click
            AddHandler btnNewAdjustment.Click, AddressOf btnNewAdjustment_Click
            AddHandler dtpAdjustmentDate.ValueChanged, AddressOf dtpAdjustmentDate_ValueChanged
        End Sub

        Private Async Sub FrmStockOperations_Load(sender As Object, e As EventArgs)
            Await LoadProductsAsync()
            Await PreparePurchaseReturnAsync()
            Await PrepareStockAdjustmentAsync()
            Await LoadPurchasesAsync()
        End Sub

        Private Async Function LoadProductsAsync() As Task
            _products = Await _stockOperationService.LoadProductsAsync()
            cboAdjustmentProduct.DataSource = Nothing
            cboAdjustmentProduct.DataSource = _products
            cboAdjustmentProduct.SelectedIndex = -1
        End Function

        Private Async Function PreparePurchaseReturnAsync() As Task
            txtReturnNumber.Text = Await _stockOperationService.GenerateNextPurchaseReturnNumberAsync(dtpReturnDate.Value.Date)
            txtReturnNotes.Clear()
            _purchaseReturnItems.Clear()
            lblSelectedPurchase.Text = "Select a purchase to load returnable lines."
            ResetReturnTotals()
        End Function

        Private Async Function PrepareStockAdjustmentAsync() As Task
            txtAdjustmentNumber.Text = Await _stockOperationService.GenerateNextStockAdjustmentNumberAsync(dtpAdjustmentDate.Value.Date)
            txtAdjustmentNotes.Clear()
            txtAdjustmentLineRemarks.Clear()
            nudAdjustmentQuantity.Value = 1D
            If cboAdjustmentMode.Items.Count > 0 Then
                cboAdjustmentMode.SelectedIndex = 0
            End If
            cboAdjustmentProduct.SelectedIndex = -1
            _stockAdjustmentItems.Clear()
            RefreshAdjustmentSummary()
        End Function

        Private Async Sub btnRefreshPurchases_Click(sender As Object, e As EventArgs)
            Await LoadPurchasesAsync(GetSelectedPurchaseId())
        End Sub

        Private Async Function LoadPurchasesAsync(Optional purchaseIdToSelect As Integer = 0) As Task
            If dtpPurchaseFrom.Value.Date > dtpPurchaseTo.Value.Date Then
                ShowStatus("Purchase From date cannot be later than To date.", True)
                Return
            End If

            SetBusy(True, "Loading purchases for return...")

            Try
                Dim rows As List(Of PurchaseHistoryLookupRow) = Await _stockOperationService.SearchPurchasesForReturnAsync(dtpPurchaseFrom.Value.Date, dtpPurchaseTo.Value.Date, txtPurchaseSearch.Text.Trim())
                _purchaseRows.Clear()
                For Each row As PurchaseHistoryLookupRow In rows
                    _purchaseRows.Add(row)
                Next

                If _purchaseRows.Count = 0 Then
                    _purchaseReturnItems.Clear()
                    lblSelectedPurchase.Text = "No purchases matched the current filters."
                    ResetReturnTotals()
                    ShowStatus("No purchases matched the current filters.", False)
                Else
                    SelectPurchaseRow(purchaseIdToSelect)
                    ShowStatus("Purchase list loaded successfully.", False)
                End If
            Catch ex As Exception
                AppLogger.Error("Purchase return search failed.", ex)
                ShowStatus("Purchases could not be loaded for return.", True)
            Finally
                SetBusy(False)
            End Try

            If GetSelectedPurchase() IsNot Nothing Then
                Await LoadSelectedPurchaseLinesAsync()
            End If
        End Function

        Private Async Sub dgvPurchases_SelectionChanged(sender As Object, e As EventArgs)
            If _isBusy Then
                Return
            End If

            Await LoadSelectedPurchaseLinesAsync()
        End Sub

        Private Async Function LoadSelectedPurchaseLinesAsync() As Task
            Dim purchase As PurchaseHistoryLookupRow = GetSelectedPurchase()
            If purchase Is Nothing Then
                _purchaseReturnItems.Clear()
                lblSelectedPurchase.Text = "Select a purchase to load returnable lines."
                ResetReturnTotals()
                Return
            End If

            SetBusy(True, $"Loading returnable lines for {purchase.PurchaseNumber}...")

            Try
                Dim lines As List(Of PurchaseReturnLineItem) = Await _stockOperationService.GetPurchaseReturnLinesAsync(purchase.PurchaseId)
                _purchaseReturnItems.Clear()
                For Each line As PurchaseReturnLineItem In lines
                    _purchaseReturnItems.Add(line)
                Next
                lblSelectedPurchase.Text = $"Purchase {purchase.PurchaseNumber} | Supplier {purchase.SupplierName} | Net {purchase.NetAmount:N2}"
                RefreshReturnTotals()
                ShowStatus("Returnable purchase lines loaded.", False)
            Catch ex As Exception
                AppLogger.Error($"Purchase return lines could not be loaded for purchase Id {purchase.PurchaseId}.", ex)
                _purchaseReturnItems.Clear()
                lblSelectedPurchase.Text = "The selected purchase could not be loaded."
                ResetReturnTotals()
                ShowStatus("Returnable purchase lines could not be loaded.", True)
            Finally
                SetBusy(False)
            End Try
        End Function

        Private Sub dgvReturnItems_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs)
            If e.RowIndex < 0 OrElse e.RowIndex >= _purchaseReturnItems.Count Then
                Return
            End If

            Dim line As PurchaseReturnLineItem = _purchaseReturnItems(e.RowIndex)
            line.ReturnQuantity = Math.Min(Math.Max(line.ReturnQuantity, 0), line.RemainingQuantity)
            line.ReturnFreeQuantity = Math.Min(Math.Max(line.ReturnFreeQuantity, 0), line.RemainingFreeQuantity)
            _stockOperationService.RecalculatePurchaseReturnLine(line)
            dgvReturnItems.Refresh()
            RefreshReturnTotals()
        End Sub

        Private Sub dgvReturnItems_DataError(sender As Object, e As DataGridViewDataErrorEventArgs)
            e.Cancel = True
            ShowStatus("Enter a valid numeric value in the return grid.", True)
        End Sub

        Private Async Sub btnSaveReturn_Click(sender As Object, e As EventArgs)
            Dim purchase As PurchaseHistoryLookupRow = GetSelectedPurchase()
            If purchase Is Nothing Then
                ShowStatus("Select a purchase before saving a return.", True)
                Return
            End If

            Dim draft As New PurchaseReturnDraft With {
                .ReturnNumber = txtReturnNumber.Text,
                .PurchaseId = purchase.PurchaseId,
                .PurchaseNumber = purchase.PurchaseNumber,
                .SupplierId = purchase.SupplierId,
                .SupplierName = purchase.SupplierName,
                .ReturnDate = dtpReturnDate.Value.Date,
                .Notes = txtReturnNotes.Text
            }

            For Each line As PurchaseReturnLineItem In _purchaseReturnItems
                draft.Items.Add(line)
            Next

            draft.Summary = _stockOperationService.CalculatePurchaseReturnTotals(draft.Items)

            SetBusy(True, "Saving purchase return...")
            Dim result As PurchaseReturnSaveResult = Await _stockOperationService.SavePurchaseReturnAsync(draft, If(SessionManager.CurrentUser Is Nothing, 0, SessionManager.CurrentUser.Id))
            SetBusy(False)
            ShowStatus(result.Message, Not result.IsSuccess)

            If result.IsSuccess Then
                Await LoadProductsAsync()
                Await PreparePurchaseReturnAsync()
                Await LoadPurchasesAsync(purchase.PurchaseId)
            End If
        End Sub

        Private Async Sub btnResetReturn_Click(sender As Object, e As EventArgs)
            Await PreparePurchaseReturnAsync()
            Await LoadSelectedPurchaseLinesAsync()
            ShowStatus("Purchase return form reset.", False)
        End Sub

        Private Async Sub dtpReturnDate_ValueChanged(sender As Object, e As EventArgs)
            If Not _isBusy Then
                txtReturnNumber.Text = Await _stockOperationService.GenerateNextPurchaseReturnNumberAsync(dtpReturnDate.Value.Date)
            End If
        End Sub

        Private Sub RefreshReturnTotals()
            Dim summary As PurchaseReturnSummary = _stockOperationService.CalculatePurchaseReturnTotals(_purchaseReturnItems.ToList())
            lblReturnSubTotalValue.Text = summary.SubTotal.ToString("N2")
            lblReturnGstValue.Text = summary.GstAmount.ToString("N2")
            lblReturnRoundOffValue.Text = summary.RoundOffAmount.ToString("N2")
            lblReturnNetValue.Text = summary.NetAmount.ToString("N2")
            txtReturnTaxSummary.Text = summary.TaxSummaryText
        End Sub

        Private Sub ResetReturnTotals()
            lblReturnSubTotalValue.Text = "0.00"
            lblReturnGstValue.Text = "0.00"
            lblReturnRoundOffValue.Text = "0.00"
            lblReturnNetValue.Text = "0.00"
            txtReturnTaxSummary.Text = "No return lines selected."
        End Sub

        Private Sub cboAdjustmentProduct_SelectedIndexChanged(sender As Object, e As EventArgs)
            If _isBusy Then
                Return
            End If

            Dim product As ProductRecord = TryCast(cboAdjustmentProduct.SelectedItem, ProductRecord)
            If product Is Nothing Then
                Return
            End If

            txtAdjustmentLineRemarks.Text = $"Manual {If(cboAdjustmentMode.SelectedIndex = 1, "decrease", "increase")} for {product.ProductName}"
        End Sub

        Private Sub btnAddAdjustmentLine_Click(sender As Object, e As EventArgs)
            Dim product As ProductRecord = TryCast(cboAdjustmentProduct.SelectedItem, ProductRecord)
            If product Is Nothing Then
                ShowStatus("Select a product for the stock adjustment line.", True)
                Return
            End If

            Dim selectedMode As StockAdjustmentMode = If(cboAdjustmentMode.SelectedIndex = 1, StockAdjustmentMode.Decrease, StockAdjustmentMode.Increase)
            Dim quantity As Integer = Decimal.ToInt32(nudAdjustmentQuantity.Value)
            If quantity <= 0 Then
                ShowStatus("Adjustment quantity must be greater than zero.", True)
                Return
            End If

            Dim existingLine As StockAdjustmentLineItem =
                _stockAdjustmentItems.FirstOrDefault(
                    Function(line) line.ProductId = product.Id AndAlso line.AdjustmentMode = selectedMode AndAlso
                                   String.Equals(line.Remarks, txtAdjustmentLineRemarks.Text.Trim(), StringComparison.OrdinalIgnoreCase))

            If existingLine IsNot Nothing Then
                existingLine.Quantity += quantity
                _stockOperationService.RecalculateStockAdjustmentLine(existingLine)
                dgvAdjustmentItems.Refresh()
            Else
                Dim line As StockAdjustmentLineItem = _stockOperationService.CreateAdjustmentLineFromProduct(product, selectedMode)
                line.Quantity = quantity
                line.Remarks = txtAdjustmentLineRemarks.Text.Trim()
                _stockOperationService.RecalculateStockAdjustmentLine(line)
                _stockAdjustmentItems.Add(line)
            End If

            ReindexAdjustmentLines()
            RefreshAdjustmentSummary()
            nudAdjustmentQuantity.Value = 1D
            txtAdjustmentLineRemarks.Clear()
            ShowStatus($"Added adjustment line for '{product.ProductName}'.", False)
        End Sub

        Private Sub btnRemoveAdjustmentLine_Click(sender As Object, e As EventArgs)
            If dgvAdjustmentItems.CurrentRow Is Nothing Then
                ShowStatus("Select an adjustment line to remove.", True)
                Return
            End If

            Dim line As StockAdjustmentLineItem = TryCast(dgvAdjustmentItems.CurrentRow.DataBoundItem, StockAdjustmentLineItem)
            If line Is Nothing Then
                Return
            End If

            _stockAdjustmentItems.Remove(line)
            ReindexAdjustmentLines()
            RefreshAdjustmentSummary()
            ShowStatus("Adjustment line removed.", False)
        End Sub

        Private Async Sub btnSaveAdjustment_Click(sender As Object, e As EventArgs)
            Dim draft As New StockAdjustmentDraft With {
                .AdjustmentNumber = txtAdjustmentNumber.Text,
                .AdjustmentDate = dtpAdjustmentDate.Value.Date,
                .Notes = txtAdjustmentNotes.Text
            }

            For Each line As StockAdjustmentLineItem In _stockAdjustmentItems
                draft.Items.Add(line)
            Next

            SetBusy(True, "Saving stock adjustment...")
            Dim result As StockAdjustmentSaveResult = Await _stockOperationService.SaveStockAdjustmentAsync(draft, If(SessionManager.CurrentUser Is Nothing, 0, SessionManager.CurrentUser.Id))
            SetBusy(False)
            ShowStatus(result.Message, Not result.IsSuccess)

            If result.IsSuccess Then
                Await LoadProductsAsync()
                Await PrepareStockAdjustmentAsync()
            End If
        End Sub

        Private Async Sub btnNewAdjustment_Click(sender As Object, e As EventArgs)
            Await PrepareStockAdjustmentAsync()
            ShowStatus("Ready to create a new stock adjustment.", False)
        End Sub

        Private Async Sub dtpAdjustmentDate_ValueChanged(sender As Object, e As EventArgs)
            If _stockAdjustmentItems.Count = 0 AndAlso Not _isBusy Then
                txtAdjustmentNumber.Text = Await _stockOperationService.GenerateNextStockAdjustmentNumberAsync(dtpAdjustmentDate.Value.Date)
            End If
        End Sub

        Private Sub ReindexAdjustmentLines()
            For index As Integer = 0 To _stockAdjustmentItems.Count - 1
                _stockAdjustmentItems(index).LineNumber = index + 1
            Next
            dgvAdjustmentItems.Refresh()
        End Sub

        Private Sub RefreshAdjustmentSummary()
            Dim increaseUnits As Integer = _stockAdjustmentItems.Where(Function(line) line.AdjustmentMode = StockAdjustmentMode.Increase).Sum(Function(line) line.Quantity)
            Dim decreaseUnits As Integer = _stockAdjustmentItems.Where(Function(line) line.AdjustmentMode = StockAdjustmentMode.Decrease).Sum(Function(line) line.Quantity)
            lblAdjustmentSummary.Text = $"Lines: {_stockAdjustmentItems.Count:N0} | Increase Units: {increaseUnits:N0} | Decrease Units: {decreaseUnits:N0}"
        End Sub

        Private Function GetSelectedPurchase() As PurchaseHistoryLookupRow
            If dgvPurchases.CurrentRow Is Nothing Then
                Return Nothing
            End If

            Return TryCast(dgvPurchases.CurrentRow.DataBoundItem, PurchaseHistoryLookupRow)
        End Function

        Private Function GetSelectedPurchaseId() As Integer
            Dim purchase As PurchaseHistoryLookupRow = GetSelectedPurchase()
            If purchase Is Nothing Then
                Return 0
            End If

            Return purchase.PurchaseId
        End Function

        Private Sub SelectPurchaseRow(purchaseIdToSelect As Integer)
            If dgvPurchases.Rows.Count = 0 Then
                Return
            End If

            If purchaseIdToSelect > 0 Then
                For Each row As DataGridViewRow In dgvPurchases.Rows
                    Dim purchase As PurchaseHistoryLookupRow = TryCast(row.DataBoundItem, PurchaseHistoryLookupRow)
                    If purchase IsNot Nothing AndAlso purchase.PurchaseId = purchaseIdToSelect Then
                        row.Selected = True
                        dgvPurchases.CurrentCell = row.Cells(0)
                        Return
                    End If
                Next
            End If

            dgvPurchases.Rows(0).Selected = True
            dgvPurchases.CurrentCell = dgvPurchases.Rows(0).Cells(0)
        End Sub

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

        Private Sub txtPurchaseSearch_KeyDown(sender As Object, e As KeyEventArgs)
            If e.KeyCode = Keys.Enter AndAlso Not _isBusy Then
                e.SuppressKeyPress = True
                Dim loadTask As Task = LoadPurchasesAsync(GetSelectedPurchaseId())
            End If
        End Sub

        Private Sub SetBusy(isBusy As Boolean, Optional message As String = "")
            _isBusy = isBusy

            txtPurchaseSearch.Enabled = Not isBusy
            dtpPurchaseFrom.Enabled = Not isBusy
            dtpPurchaseTo.Enabled = Not isBusy
            btnRefreshPurchases.Enabled = Not isBusy
            dgvPurchases.Enabled = Not isBusy
            txtReturnNotes.Enabled = Not isBusy
            dtpReturnDate.Enabled = Not isBusy
            btnSaveReturn.Enabled = Not isBusy
            btnResetReturn.Enabled = Not isBusy
            dgvReturnItems.Enabled = Not isBusy

            dtpAdjustmentDate.Enabled = Not isBusy
            txtAdjustmentNotes.Enabled = Not isBusy
            cboAdjustmentProduct.Enabled = Not isBusy
            cboAdjustmentMode.Enabled = Not isBusy
            nudAdjustmentQuantity.Enabled = Not isBusy
            txtAdjustmentLineRemarks.Enabled = Not isBusy
            btnAddAdjustmentLine.Enabled = Not isBusy
            btnRemoveAdjustmentLine.Enabled = Not isBusy
            btnSaveAdjustment.Enabled = Not isBusy
            btnNewAdjustment.Enabled = Not isBusy
            dgvAdjustmentItems.Enabled = Not isBusy

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
                    If tabs.SelectedIndex = 0 Then
                        Dim loadTask As Task = LoadPurchasesAsync(GetSelectedPurchaseId())
                    Else
                        Dim prepTask As Task = PrepareStockAdjustmentAsync()
                    End If
                    Return True
                Case Keys.Control Or Keys.S
                    If tabs.SelectedIndex = 0 Then
                        btnSaveReturn.PerformClick()
                    Else
                        btnSaveAdjustment.PerformClick()
                    End If
                    Return True
                Case Keys.Control Or Keys.N
                    If tabs.SelectedIndex = 0 Then
                        btnResetReturn.PerformClick()
                    Else
                        btnNewAdjustment.PerformClick()
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
