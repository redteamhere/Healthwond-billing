Imports HealthwondBilling.Models
Imports HealthwondBilling.Services
Imports HealthwondBilling.Utilities

Namespace Forms

    Public Class FrmProducts
        Inherits Form

        Private ReadOnly _productService As ProductService

        Private ReadOnly txtSearch As New TextBox()
        Private ReadOnly btnSearch As New Button()
        Private ReadOnly btnRefresh As New Button()
        Private ReadOnly btnNew As New Button()
        Private ReadOnly btnSave As New Button()
        Private ReadOnly btnDelete As New Button()
        Private ReadOnly btnClose As New Button()
        Private ReadOnly dgvProducts As New DataGridView()
        Private ReadOnly lblRecordCount As New Label()
        Private ReadOnly lblStatus As New Label()

        Private ReadOnly txtProductName As New TextBox()
        Private ReadOnly txtCompanyName As New TextBox()
        Private ReadOnly txtPacking As New TextBox()
        Private ReadOnly txtHsnCode As New TextBox()
        Private ReadOnly txtBatchNumber As New TextBox()
        Private ReadOnly txtBarcode As New TextBox()
        Private ReadOnly txtComposition As New TextBox()
        Private ReadOnly dtpExpiryDate As New DateTimePicker()
        Private ReadOnly nudGst As New NumericUpDown()
        Private ReadOnly nudMRP As New NumericUpDown()
        Private ReadOnly nudPTR As New NumericUpDown()
        Private ReadOnly nudPTS As New NumericUpDown()
        Private ReadOnly nudCurrentStock As New NumericUpDown()

        Private _currentProductId As Integer
        Private _isBusy As Boolean

        Public Sub New(productService As ProductService)
            _productService = productService

            Text = "Healthwond Billing System - Product Master"
            StartPosition = FormStartPosition.CenterParent
            Size = New Size(1420, 860)
            MinimumSize = New Size(1280, 780)
            BackColor = ThemePalette.AppBackground
            Font = New Font("Segoe UI", 10.0F, FontStyle.Regular)
            KeyPreview = True

            BuildLayout()
            ConfigureGrid()
            ConfigureEditors()
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
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 88))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 112))
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 36))

            root.Controls.Add(BuildHeaderPanel(), 0, 0)
            root.Controls.Add(BuildToolbarPanel(), 0, 1)
            root.Controls.Add(BuildMainPanel(), 0, 2)

            lblStatus.Dock = DockStyle.Fill
            lblStatus.Font = New Font("Segoe UI Semibold", 9.5F, FontStyle.Bold)
            lblStatus.ForeColor = ThemePalette.TextMuted
            lblStatus.TextAlign = ContentAlignment.MiddleLeft
            root.Controls.Add(lblStatus, 0, 3)

            Controls.Add(root)
        End Sub

        Private Function BuildHeaderPanel() As Panel
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = ThemePalette.AppBackground}

            Dim title As New Label With {
                .Dock = DockStyle.Top,
                .Height = 44,
                .Text = "Product Master",
                .Font = New Font("Segoe UI Semibold", 24.0F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary
            }

            Dim subtitle As New Label With {
                .Dock = DockStyle.Fill,
                .Text = "Create, search, edit, and retire pharmaceutical products with batch, price, GST, barcode, and stock controls.",
                .Font = New Font("Segoe UI", 10.5F, FontStyle.Regular),
                .ForeColor = ThemePalette.TextMuted
            }

            panel.Controls.Add(subtitle)
            panel.Controls.Add(title)
            Return panel
        End Function

        Private Function BuildToolbarPanel() As Panel
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White, .Padding = New Padding(18)}
            UiStyler.StyleCard(panel)

            Dim searchPanel As New Panel With {.Dock = DockStyle.Left, .Width = 420}
            Dim searchLabel As New Label With {
                .Dock = DockStyle.Top,
                .Height = 22,
                .Text = "Search products",
                .Font = New Font("Segoe UI Semibold", 9.5F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary
            }

            txtSearch.Dock = DockStyle.Bottom
            txtSearch.Height = 40
            txtSearch.BorderStyle = BorderStyle.FixedSingle
            UiStyler.StyleInput(txtSearch)

            searchPanel.Controls.Add(txtSearch)
            searchPanel.Controls.Add(searchLabel)

            Dim actionFlow As New FlowLayoutPanel With {
                .Dock = DockStyle.Right,
                .Width = 610,
                .FlowDirection = FlowDirection.LeftToRight,
                .WrapContents = False,
                .AutoSize = False,
                .BackColor = Color.Transparent,
                .Padding = New Padding(0, 10, 0, 0)
            }

            ConfigureToolbarButton(btnSearch, "Search", AddressOf btnSearch_Click, 95)
            ConfigureToolbarButton(btnRefresh, "Refresh", AddressOf btnRefresh_Click, 95)
            ConfigureToolbarButton(btnNew, "New", AddressOf btnNew_Click, 95)
            ConfigureToolbarButton(btnSave, "Save", AddressOf btnSave_Click, 95)
            ConfigureToolbarButton(btnClose, "Close", AddressOf btnClose_Click, 95)
            UiStyler.StyleDangerButton(btnDelete)
            btnDelete.Text = "Delete"
            btnDelete.Width = 95
            AddHandler btnDelete.Click, AddressOf btnDelete_Click

            actionFlow.Controls.AddRange(New Control() {btnSearch, btnRefresh, btnNew, btnSave, btnDelete, btnClose})

            panel.Controls.Add(actionFlow)
            panel.Controls.Add(searchPanel)
            Return panel
        End Function

        Private Function BuildMainPanel() As Control
            Dim split As New SplitContainer With {
                .Dock = DockStyle.Fill,
                .SplitterDistance = 780,
                .BackColor = ThemePalette.AppBackground,
                .FixedPanel = FixedPanel.None
            }

            split.Panel1.Controls.Add(BuildGridCard())
            split.Panel2.Controls.Add(BuildEditorCard())
            Return split
        End Function

        Private Function BuildGridCard() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)

            Dim title As New Label With {
                .Dock = DockStyle.Top,
                .Height = 30,
                .Text = "Available products",
                .Font = New Font("Segoe UI Semibold", 14.0F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary
            }

            lblRecordCount.Dock = DockStyle.Top
            lblRecordCount.Height = 24
            lblRecordCount.Font = New Font("Segoe UI", 9.5F, FontStyle.Regular)
            lblRecordCount.ForeColor = ThemePalette.TextMuted

            dgvProducts.Dock = DockStyle.Fill

            panel.Controls.Add(dgvProducts)
            panel.Controls.Add(lblRecordCount)
            panel.Controls.Add(title)
            Return panel
        End Function

        Private Function BuildEditorCard() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)

            Dim title As New Label With {
                .Dock = DockStyle.Top,
                .Height = 30,
                .Text = "Product details",
                .Font = New Font("Segoe UI Semibold", 14.0F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary
            }

            Dim editor As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2,
                .Padding = New Padding(0, 14, 0, 0)
            }
            editor.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0F))
            editor.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0F))
            For index As Integer = 0 To 5
                editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 72))
            Next
            editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 110))
            editor.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            editor.Controls.Add(CreateInputHost("Product Name", txtProductName), 0, 0)
            editor.Controls.Add(CreateInputHost("Company Name", txtCompanyName), 1, 0)
            editor.Controls.Add(CreateInputHost("Packing", txtPacking), 0, 1)
            editor.Controls.Add(CreateInputHost("HSN Code", txtHsnCode), 1, 1)
            editor.Controls.Add(CreateInputHost("Batch Number", txtBatchNumber), 0, 2)
            editor.Controls.Add(CreateInputHost("Barcode", txtBarcode), 1, 2)
            editor.Controls.Add(CreateInputHost("Expiry Date", dtpExpiryDate), 0, 3)
            editor.Controls.Add(CreateInputHost("GST %", nudGst), 1, 3)
            editor.Controls.Add(CreateInputHost("MRP", nudMRP), 0, 4)
            editor.Controls.Add(CreateInputHost("PTR", nudPTR), 1, 4)
            editor.Controls.Add(CreateInputHost("PTS", nudPTS), 0, 5)
            editor.Controls.Add(CreateInputHost("Current Stock", nudCurrentStock), 1, 5)

            txtComposition.Multiline = True
            txtComposition.ScrollBars = ScrollBars.Vertical
            Dim compositionHost As Control = CreateInputHost("Composition", txtComposition)
            editor.Controls.Add(compositionHost, 0, 6)
            editor.SetColumnSpan(compositionHost, 2)

            panel.Controls.Add(UiStyler.CreateScrollableHost(editor))
            panel.Controls.Add(title)
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

        Private Sub ConfigureToolbarButton(button As Button, text As String, handler As EventHandler, width As Integer)
            UiStyler.StyleSecondaryButton(button)
            button.Text = text
            button.Width = width
            AddHandler button.Click, handler
        End Sub

        Private Sub ConfigureGrid()
            UiStyler.StyleDataGrid(dgvProducts)
            dgvProducts.Columns.Clear()

            dgvProducts.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "ProductName", .HeaderText = "Product", .FillWeight = 170.0F})
            dgvProducts.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "CompanyName", .HeaderText = "Company", .FillWeight = 130.0F})
            dgvProducts.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "BatchNumber", .HeaderText = "Batch", .FillWeight = 90.0F})
            dgvProducts.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "Packing", .HeaderText = "Packing", .FillWeight = 90.0F})
            dgvProducts.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "ExpiryDate", .HeaderText = "Expiry", .FillWeight = 85.0F, .DefaultCellStyle = New DataGridViewCellStyle With {.Format = "dd-MMM-yyyy"}})
            dgvProducts.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "CurrentStock", .HeaderText = "Stock", .FillWeight = 70.0F})
            dgvProducts.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "PTS", .HeaderText = "PTS", .FillWeight = 75.0F, .DefaultCellStyle = New DataGridViewCellStyle With {.Format = "N2"}})
            dgvProducts.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "MRP", .HeaderText = "MRP", .FillWeight = 75.0F, .DefaultCellStyle = New DataGridViewCellStyle With {.Format = "N2"}})
        End Sub

        Private Sub ConfigureEditors()
            For Each editorTextBox As TextBox In New TextBox() {txtProductName, txtCompanyName, txtPacking, txtHsnCode, txtBatchNumber, txtBarcode, txtComposition}
                editorTextBox.BorderStyle = BorderStyle.FixedSingle
                UiStyler.StyleInput(editorTextBox)
            Next

            dtpExpiryDate.Format = DateTimePickerFormat.Custom
            dtpExpiryDate.CustomFormat = "dd-MMM-yyyy"

            ConfigureNumeric(nudGst, 2, 1000D, 0.25D)
            ConfigureNumeric(nudMRP, 2, 1000000D, 1D)
            ConfigureNumeric(nudPTR, 2, 1000000D, 1D)
            ConfigureNumeric(nudPTS, 2, 1000000D, 1D)
            ConfigureNumeric(nudCurrentStock, 0, 1000000D, 1D)
        End Sub

        Private Sub ConfigureNumeric(control As NumericUpDown, decimalPlaces As Integer, maximum As Decimal, increment As Decimal)
            control.DecimalPlaces = decimalPlaces
            control.Maximum = maximum
            control.Minimum = 0D
            control.Increment = increment
            control.ThousandsSeparator = True
            control.Font = New Font("Segoe UI", 10.5F, FontStyle.Regular)
        End Sub

        Private Sub WireEvents()
            AddHandler Load, AddressOf FrmProducts_Load
            AddHandler dgvProducts.SelectionChanged, AddressOf dgvProducts_SelectionChanged
            AddHandler txtSearch.KeyDown, AddressOf txtSearch_KeyDown
        End Sub

        Private Async Sub FrmProducts_Load(sender As Object, e As EventArgs)
            PrepareNewProduct()
            Await LoadProductsAsync()
        End Sub

        Private Async Sub btnSearch_Click(sender As Object, e As EventArgs)
            Await LoadProductsAsync()
        End Sub

        Private Async Sub btnRefresh_Click(sender As Object, e As EventArgs)
            Await LoadProductsAsync(txtSearch.Text, _currentProductId)
        End Sub

        Private Sub btnNew_Click(sender As Object, e As EventArgs)
            PrepareNewProduct()
            ShowStatus("Ready to add a new product.", False)
        End Sub

        Private Async Sub btnSave_Click(sender As Object, e As EventArgs)
            Await SaveCurrentAsync()
        End Sub

        Private Async Sub btnDelete_Click(sender As Object, e As EventArgs)
            Dim currentProduct As ProductRecord = GetSelectedProduct()
            If currentProduct Is Nothing Then
                ShowStatus("Select a product to delete.", True)
                Return
            End If

            Dim confirmation As DialogResult = MessageBox.Show(
                $"Delete '{currentProduct.ProductName}' from the product master?",
                "Confirm Delete",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question)

            If confirmation <> DialogResult.Yes Then
                Return
            End If

            SetBusy(True, "Deleting product...")
            Dim result As EntityOperationResult = Await _productService.DeleteAsync(currentProduct)
            SetBusy(False)
            ShowStatus(result.Message, Not result.IsSuccess)

            If result.IsSuccess Then
                Await LoadProductsAsync(txtSearch.Text)
            End If
        End Sub

        Private Sub btnClose_Click(sender As Object, e As EventArgs)
            Close()
        End Sub

        Private Async Function SaveCurrentAsync() As Task
            Dim product As ProductRecord = ReadProductFromForm()
            SetBusy(True, "Saving product...")
            Dim result As EntityOperationResult = Await _productService.SaveAsync(product)
            SetBusy(False)
            ShowStatus(result.Message, Not result.IsSuccess)

            If result.IsSuccess Then
                Await LoadProductsAsync(txtSearch.Text, result.EntityId)
            End If
        End Function

        Private Async Function LoadProductsAsync(Optional searchTerm As String = Nothing, Optional preferredProductId As Integer = 0) As Task
            SetBusy(True, "Loading products...")

            Try
                Dim products As List(Of ProductRecord) = Await _productService.SearchAsync(If(searchTerm, txtSearch.Text))
                dgvProducts.DataSource = Nothing
                dgvProducts.DataSource = products
                lblRecordCount.Text = $"{products.Count:N0} product(s)"

                If products.Count = 0 Then
                    PrepareNewProduct()
                    ShowStatus("No products matched the current search.", False)
                    Return
                End If

                SelectProductInGrid(preferredProductId)
                If preferredProductId = 0 Then
                    ShowStatus("Products loaded successfully.", False)
                End If
            Catch ex As Exception
                AppLogger.Error("Product list load failed.", ex)
                ShowStatus("Products could not be loaded.", True)
            Finally
                SetBusy(False)
            End Try
        End Function

        Private Sub SelectProductInGrid(preferredProductId As Integer)
            Dim targetIndex As Integer = 0

            If preferredProductId > 0 Then
                For index As Integer = 0 To dgvProducts.Rows.Count - 1
                    Dim rowProduct As ProductRecord = TryCast(dgvProducts.Rows(index).DataBoundItem, ProductRecord)
                    If rowProduct IsNot Nothing AndAlso rowProduct.Id = preferredProductId Then
                        targetIndex = index
                        Exit For
                    End If
                Next
            End If

            If dgvProducts.Rows.Count = 0 Then
                Return
            End If

            dgvProducts.ClearSelection()
            dgvProducts.Rows(targetIndex).Selected = True
            dgvProducts.CurrentCell = dgvProducts.Rows(targetIndex).Cells(0)
            BindProduct(TryCast(dgvProducts.Rows(targetIndex).DataBoundItem, ProductRecord))
        End Sub

        Private Sub dgvProducts_SelectionChanged(sender As Object, e As EventArgs)
            If _isBusy Then
                Return
            End If

            BindProduct(GetSelectedProduct())
        End Sub

        Private Function GetSelectedProduct() As ProductRecord
            If dgvProducts.CurrentRow Is Nothing Then
                Return Nothing
            End If

            Return TryCast(dgvProducts.CurrentRow.DataBoundItem, ProductRecord)
        End Function

        Private Sub BindProduct(product As ProductRecord)
            If product Is Nothing Then
                Return
            End If

            _currentProductId = product.Id
            txtProductName.Text = product.ProductName
            txtCompanyName.Text = product.CompanyName
            txtPacking.Text = product.Packing
            txtHsnCode.Text = product.HsnCode
            txtBatchNumber.Text = product.BatchNumber
            txtBarcode.Text = product.Barcode
            txtComposition.Text = product.Composition
            dtpExpiryDate.Value = If(product.ExpiryDate < dtpExpiryDate.MinDate, dtpExpiryDate.MinDate, product.ExpiryDate)
            nudGst.Value = Math.Min(nudGst.Maximum, product.GstPercentage)
            nudMRP.Value = Math.Min(nudMRP.Maximum, product.MRP)
            nudPTR.Value = Math.Min(nudPTR.Maximum, product.PTR)
            nudPTS.Value = Math.Min(nudPTS.Maximum, product.PTS)
            nudCurrentStock.Value = Math.Min(nudCurrentStock.Maximum, product.CurrentStock)
            UpdateActionState()
        End Sub

        Private Function ReadProductFromForm() As ProductRecord
            Return New ProductRecord With {
                .Id = _currentProductId,
                .ProductName = txtProductName.Text,
                .CompanyName = txtCompanyName.Text,
                .Packing = txtPacking.Text,
                .HsnCode = txtHsnCode.Text,
                .BatchNumber = txtBatchNumber.Text,
                .Barcode = txtBarcode.Text,
                .Composition = txtComposition.Text,
                .ExpiryDate = dtpExpiryDate.Value.Date,
                .GstPercentage = nudGst.Value,
                .MRP = nudMRP.Value,
                .PTR = nudPTR.Value,
                .PTS = nudPTS.Value,
                .CurrentStock = Decimal.ToInt32(nudCurrentStock.Value)
            }
        End Function

        Private Sub PrepareNewProduct()
            _currentProductId = 0
            txtProductName.Clear()
            txtCompanyName.Clear()
            txtPacking.Clear()
            txtHsnCode.Clear()
            txtBatchNumber.Clear()
            txtBarcode.Clear()
            txtComposition.Clear()
            dtpExpiryDate.Value = DateTime.Today.AddYears(1)
            nudGst.Value = 0D
            nudMRP.Value = 0D
            nudPTR.Value = 0D
            nudPTS.Value = 0D
            nudCurrentStock.Value = 0D
            dgvProducts.ClearSelection()
            UpdateActionState()
            txtProductName.Focus()
        End Sub

        Private Sub UpdateActionState()
            btnDelete.Enabled = _currentProductId > 0
        End Sub

        Private Sub SetBusy(isBusy As Boolean, Optional statusMessage As String = "")
            _isBusy = isBusy
            txtSearch.Enabled = Not isBusy
            btnSearch.Enabled = Not isBusy
            btnRefresh.Enabled = Not isBusy
            btnNew.Enabled = Not isBusy
            btnSave.Enabled = Not isBusy
            btnDelete.Enabled = Not isBusy AndAlso _currentProductId > 0
            btnClose.Enabled = Not isBusy
            dgvProducts.Enabled = Not isBusy

            If isBusy Then
                lblStatus.ForeColor = ThemePalette.TextMuted
                lblStatus.Text = statusMessage
            End If
        End Sub

        Private Sub ShowStatus(message As String, isError As Boolean)
            lblStatus.ForeColor = If(isError, ThemePalette.DangerRed, ThemePalette.AccentGreen)
            lblStatus.Text = message
        End Sub

        Private Async Sub txtSearch_KeyDown(sender As Object, e As KeyEventArgs)
            If e.KeyCode = Keys.Enter Then
                e.SuppressKeyPress = True
                Await LoadProductsAsync()
            End If
        End Sub

        Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
            Select Case keyData
                Case Keys.Control Or Keys.N
                    PrepareNewProduct()
                    Return True
                Case Keys.Control Or Keys.S
                    Dim saveTask As Task = SaveCurrentAsync()
                    Return True
                Case Keys.F5
                    Dim loadTask As Task = LoadProductsAsync(txtSearch.Text, _currentProductId)
                    Return True
                Case Keys.Delete
                    If btnDelete.Enabled Then
                        btnDelete.PerformClick()
                        Return True
                    End If
                Case Keys.Escape
                    Close()
                    Return True
            End Select

            Return MyBase.ProcessCmdKey(msg, keyData)
        End Function

    End Class

End Namespace
