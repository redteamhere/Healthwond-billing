Imports HealthwondBilling.Models
Imports HealthwondBilling.Services
Imports HealthwondBilling.Utilities

Namespace Forms

    Public Class FrmCustomers
        Inherits Form

        Private ReadOnly _customerService As CustomerService

        Private ReadOnly txtSearch As New TextBox()
        Private ReadOnly btnSearch As New Button()
        Private ReadOnly btnRefresh As New Button()
        Private ReadOnly btnNew As New Button()
        Private ReadOnly btnSave As New Button()
        Private ReadOnly btnDelete As New Button()
        Private ReadOnly btnClose As New Button()
        Private ReadOnly dgvCustomers As New DataGridView()
        Private ReadOnly lblRecordCount As New Label()
        Private ReadOnly lblStatus As New Label()

        Private ReadOnly txtCustomerName As New TextBox()
        Private ReadOnly txtPhone As New TextBox()
        Private ReadOnly txtGstin As New TextBox()
        Private ReadOnly txtEmail As New TextBox()
        Private ReadOnly txtDrugLicense As New TextBox()
        Private ReadOnly txtAddress As New TextBox()
        Private ReadOnly nudOutstandingBalance As New NumericUpDown()

        Private _currentCustomerId As Integer
        Private _isBusy As Boolean

        Public Sub New(customerService As CustomerService)
            _customerService = customerService

            Text = "Healthwond Billing System - Customer Master"
            StartPosition = FormStartPosition.CenterParent
            Size = New Size(1340, 820)
            MinimumSize = New Size(1200, 740)
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
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 78))
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
                .Text = "Customer Master",
                .Font = New Font("Segoe UI Semibold", 24.0F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary
            }

            Dim subtitle As New Label With {
                .Dock = DockStyle.Fill,
                .Text = "Maintain trade customers, GSTINs, drug license data, contact details, and outstanding balances for fast billing selection.",
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
                .Text = "Search customers",
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
                .SplitterDistance = 700,
                .BackColor = ThemePalette.AppBackground
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
                .Text = "Available customers",
                .Font = New Font("Segoe UI Semibold", 14.0F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary
            }

            lblRecordCount.Dock = DockStyle.Top
            lblRecordCount.Height = 24
            lblRecordCount.Font = New Font("Segoe UI", 9.5F, FontStyle.Regular)
            lblRecordCount.ForeColor = ThemePalette.TextMuted

            dgvCustomers.Dock = DockStyle.Fill

            panel.Controls.Add(dgvCustomers)
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
                .Text = "Customer details",
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
            For index As Integer = 0 To 2
                editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 72))
            Next
            editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 126))
            editor.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            editor.Controls.Add(CreateInputHost("Customer Name", txtCustomerName), 0, 0)
            editor.Controls.Add(CreateInputHost("Phone", txtPhone), 1, 0)
            editor.Controls.Add(CreateInputHost("GSTIN", txtGstin), 0, 1)
            editor.Controls.Add(CreateInputHost("Email", txtEmail), 1, 1)
            editor.Controls.Add(CreateInputHost("Drug License Number", txtDrugLicense), 0, 2)
            editor.Controls.Add(CreateInputHost("Outstanding Balance", nudOutstandingBalance), 1, 2)

            txtAddress.Multiline = True
            txtAddress.ScrollBars = ScrollBars.Vertical
            Dim addressHost As Control = CreateInputHost("Address", txtAddress)
            editor.Controls.Add(addressHost, 0, 3)
            editor.SetColumnSpan(addressHost, 2)

            panel.Controls.Add(editor)
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
            UiStyler.StyleDataGrid(dgvCustomers)
            dgvCustomers.Columns.Clear()
            dgvCustomers.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "CustomerName", .HeaderText = "Customer", .FillWeight = 170.0F})
            dgvCustomers.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "Gstin", .HeaderText = "GSTIN", .FillWeight = 120.0F})
            dgvCustomers.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "DrugLicenseNumber", .HeaderText = "Drug License", .FillWeight = 100.0F})
            dgvCustomers.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "Phone", .HeaderText = "Phone", .FillWeight = 90.0F})
            dgvCustomers.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "Email", .HeaderText = "Email", .FillWeight = 140.0F})
            dgvCustomers.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "OutstandingBalance", .HeaderText = "Outstanding", .FillWeight = 85.0F, .DefaultCellStyle = New DataGridViewCellStyle With {.Format = "N2"}})
        End Sub

        Private Sub ConfigureEditors()
            For Each editorTextBox As TextBox In New TextBox() {txtCustomerName, txtPhone, txtGstin, txtEmail, txtDrugLicense, txtAddress}
                editorTextBox.BorderStyle = BorderStyle.FixedSingle
                UiStyler.StyleInput(editorTextBox)
            Next

            nudOutstandingBalance.DecimalPlaces = 2
            nudOutstandingBalance.Maximum = 100000000D
            nudOutstandingBalance.Minimum = 0D
            nudOutstandingBalance.Increment = 50D
            nudOutstandingBalance.ThousandsSeparator = True
            nudOutstandingBalance.Font = New Font("Segoe UI", 10.5F, FontStyle.Regular)
        End Sub

        Private Sub WireEvents()
            AddHandler Load, AddressOf FrmCustomers_Load
            AddHandler dgvCustomers.SelectionChanged, AddressOf dgvCustomers_SelectionChanged
            AddHandler txtSearch.KeyDown, AddressOf txtSearch_KeyDown
        End Sub

        Private Async Sub FrmCustomers_Load(sender As Object, e As EventArgs)
            PrepareNewCustomer()
            Await LoadCustomersAsync()
        End Sub

        Private Async Sub btnSearch_Click(sender As Object, e As EventArgs)
            Await LoadCustomersAsync()
        End Sub

        Private Async Sub btnRefresh_Click(sender As Object, e As EventArgs)
            Await LoadCustomersAsync(txtSearch.Text, _currentCustomerId)
        End Sub

        Private Sub btnNew_Click(sender As Object, e As EventArgs)
            PrepareNewCustomer()
            ShowStatus("Ready to add a new customer.", False)
        End Sub

        Private Async Sub btnSave_Click(sender As Object, e As EventArgs)
            Await SaveCurrentAsync()
        End Sub

        Private Async Sub btnDelete_Click(sender As Object, e As EventArgs)
            Dim currentCustomer As CustomerRecord = GetSelectedCustomer()
            If currentCustomer Is Nothing Then
                ShowStatus("Select a customer to delete.", True)
                Return
            End If

            Dim confirmation As DialogResult = MessageBox.Show(
                $"Delete '{currentCustomer.CustomerName}' from the customer master?",
                "Confirm Delete",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question)

            If confirmation <> DialogResult.Yes Then
                Return
            End If

            SetBusy(True, "Deleting customer...")
            Dim result As EntityOperationResult = Await _customerService.DeleteAsync(currentCustomer)
            SetBusy(False)
            ShowStatus(result.Message, Not result.IsSuccess)

            If result.IsSuccess Then
                Await LoadCustomersAsync(txtSearch.Text)
            End If
        End Sub

        Private Sub btnClose_Click(sender As Object, e As EventArgs)
            Close()
        End Sub

        Private Async Function SaveCurrentAsync() As Task
            Dim customer As CustomerRecord = ReadCustomerFromForm()
            SetBusy(True, "Saving customer...")
            Dim result As EntityOperationResult = Await _customerService.SaveAsync(customer)
            SetBusy(False)
            ShowStatus(result.Message, Not result.IsSuccess)

            If result.IsSuccess Then
                Await LoadCustomersAsync(txtSearch.Text, result.EntityId)
            End If
        End Function

        Private Async Function LoadCustomersAsync(Optional searchTerm As String = Nothing, Optional preferredCustomerId As Integer = 0) As Task
            SetBusy(True, "Loading customers...")

            Try
                Dim customers As List(Of CustomerRecord) = Await _customerService.SearchAsync(If(searchTerm, txtSearch.Text))
                dgvCustomers.DataSource = Nothing
                dgvCustomers.DataSource = customers
                lblRecordCount.Text = $"{customers.Count:N0} customer(s)"

                If customers.Count = 0 Then
                    PrepareNewCustomer()
                    ShowStatus("No customers matched the current search.", False)
                    Return
                End If

                SelectCustomerInGrid(preferredCustomerId)
                If preferredCustomerId = 0 Then
                    ShowStatus("Customers loaded successfully.", False)
                End If
            Catch ex As Exception
                AppLogger.Error("Customer list load failed.", ex)
                ShowStatus("Customers could not be loaded.", True)
            Finally
                SetBusy(False)
            End Try
        End Function

        Private Sub SelectCustomerInGrid(preferredCustomerId As Integer)
            Dim targetIndex As Integer = 0

            If preferredCustomerId > 0 Then
                For index As Integer = 0 To dgvCustomers.Rows.Count - 1
                    Dim rowCustomer As CustomerRecord = TryCast(dgvCustomers.Rows(index).DataBoundItem, CustomerRecord)
                    If rowCustomer IsNot Nothing AndAlso rowCustomer.Id = preferredCustomerId Then
                        targetIndex = index
                        Exit For
                    End If
                Next
            End If

            If dgvCustomers.Rows.Count = 0 Then
                Return
            End If

            dgvCustomers.ClearSelection()
            dgvCustomers.Rows(targetIndex).Selected = True
            dgvCustomers.CurrentCell = dgvCustomers.Rows(targetIndex).Cells(0)
            BindCustomer(TryCast(dgvCustomers.Rows(targetIndex).DataBoundItem, CustomerRecord))
        End Sub

        Private Sub dgvCustomers_SelectionChanged(sender As Object, e As EventArgs)
            If _isBusy Then
                Return
            End If

            BindCustomer(GetSelectedCustomer())
        End Sub

        Private Function GetSelectedCustomer() As CustomerRecord
            If dgvCustomers.CurrentRow Is Nothing Then
                Return Nothing
            End If

            Return TryCast(dgvCustomers.CurrentRow.DataBoundItem, CustomerRecord)
        End Function

        Private Sub BindCustomer(customer As CustomerRecord)
            If customer Is Nothing Then
                Return
            End If

            _currentCustomerId = customer.Id
            txtCustomerName.Text = customer.CustomerName
            txtPhone.Text = customer.Phone
            txtGstin.Text = customer.Gstin
            txtEmail.Text = customer.Email
            txtDrugLicense.Text = customer.DrugLicenseNumber
            txtAddress.Text = customer.Address
            nudOutstandingBalance.Value = Math.Min(nudOutstandingBalance.Maximum, customer.OutstandingBalance)
            UpdateActionState()
        End Sub

        Private Function ReadCustomerFromForm() As CustomerRecord
            Return New CustomerRecord With {
                .Id = _currentCustomerId,
                .CustomerName = txtCustomerName.Text,
                .Phone = txtPhone.Text,
                .Gstin = txtGstin.Text,
                .Email = txtEmail.Text,
                .DrugLicenseNumber = txtDrugLicense.Text,
                .Address = txtAddress.Text,
                .OutstandingBalance = nudOutstandingBalance.Value
            }
        End Function

        Private Sub PrepareNewCustomer()
            _currentCustomerId = 0
            txtCustomerName.Clear()
            txtPhone.Clear()
            txtGstin.Clear()
            txtEmail.Clear()
            txtDrugLicense.Clear()
            txtAddress.Clear()
            nudOutstandingBalance.Value = 0D
            dgvCustomers.ClearSelection()
            UpdateActionState()
            txtCustomerName.Focus()
        End Sub

        Private Sub UpdateActionState()
            btnDelete.Enabled = _currentCustomerId > 0
        End Sub

        Private Sub SetBusy(isBusy As Boolean, Optional statusMessage As String = "")
            _isBusy = isBusy
            txtSearch.Enabled = Not isBusy
            btnSearch.Enabled = Not isBusy
            btnRefresh.Enabled = Not isBusy
            btnNew.Enabled = Not isBusy
            btnSave.Enabled = Not isBusy
            btnDelete.Enabled = Not isBusy AndAlso _currentCustomerId > 0
            btnClose.Enabled = Not isBusy
            dgvCustomers.Enabled = Not isBusy

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
                Await LoadCustomersAsync()
            End If
        End Sub

        Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
            Select Case keyData
                Case Keys.Control Or Keys.N
                    PrepareNewCustomer()
                    Return True
                Case Keys.Control Or Keys.S
                    Dim saveTask As Task = SaveCurrentAsync()
                    Return True
                Case Keys.F5
                    Dim loadTask As Task = LoadCustomersAsync(txtSearch.Text, _currentCustomerId)
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
