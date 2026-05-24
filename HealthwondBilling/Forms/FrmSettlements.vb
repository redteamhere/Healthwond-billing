Imports HealthwondBilling.Models
Imports HealthwondBilling.Services
Imports HealthwondBilling.Utilities
Imports System.ComponentModel
Imports System.Linq

Namespace Forms

    Public Class FrmSettlements
        Inherits Form

        Private ReadOnly _settlementService As SettlementService
        Private ReadOnly _customerService As CustomerService
        Private ReadOnly _supplierService As SupplierService

        Private ReadOnly _customerRows As New BindingList(Of CustomerRecord)()
        Private ReadOnly _supplierRows As New BindingList(Of SupplierRecord)()
        Private ReadOnly _customerPaymentRows As New BindingList(Of CustomerPaymentHistoryRow)()
        Private ReadOnly _supplierPaymentRows As New BindingList(Of SupplierPaymentHistoryRow)()

        Private ReadOnly tabs As New TabControl()
        Private ReadOnly lblStatus As New Label()

        Private ReadOnly txtCustomerSearch As New TextBox()
        Private ReadOnly btnRefreshCustomers As New Button()
        Private ReadOnly dgvCustomers As New DataGridView()
        Private ReadOnly lblSelectedCustomer As New Label()
        Private ReadOnly lblCustomerOutstandingValue As New Label()
        Private ReadOnly lblCustomerBalanceAfterValue As New Label()
        Private ReadOnly txtReceiptNumber As New TextBox()
        Private ReadOnly dtpCustomerPaymentDate As New DateTimePicker()
        Private ReadOnly cboCustomerPaymentMode As New ComboBox()
        Private ReadOnly txtCustomerReference As New TextBox()
        Private ReadOnly nudCustomerPaymentAmount As New NumericUpDown()
        Private ReadOnly txtCustomerNotes As New TextBox()
        Private ReadOnly btnCustomerFillDue As New Button()
        Private ReadOnly btnCustomerNew As New Button()
        Private ReadOnly btnCustomerSave As New Button()
        Private ReadOnly dtpCustomerHistoryFrom As New DateTimePicker()
        Private ReadOnly dtpCustomerHistoryTo As New DateTimePicker()
        Private ReadOnly txtCustomerHistorySearch As New TextBox()
        Private ReadOnly btnRefreshCustomerHistory As New Button()
        Private ReadOnly dgvCustomerHistory As New DataGridView()
        Private ReadOnly lblCustomerHistoryCaption As New Label()

        Private ReadOnly txtSupplierSearch As New TextBox()
        Private ReadOnly btnRefreshSuppliers As New Button()
        Private ReadOnly dgvSuppliers As New DataGridView()
        Private ReadOnly lblSelectedSupplier As New Label()
        Private ReadOnly lblSupplierOutstandingValue As New Label()
        Private ReadOnly lblSupplierBalanceAfterValue As New Label()
        Private ReadOnly txtSupplierPaymentNumber As New TextBox()
        Private ReadOnly dtpSupplierPaymentDate As New DateTimePicker()
        Private ReadOnly cboSupplierPaymentMode As New ComboBox()
        Private ReadOnly txtSupplierReference As New TextBox()
        Private ReadOnly nudSupplierPaymentAmount As New NumericUpDown()
        Private ReadOnly txtSupplierNotes As New TextBox()
        Private ReadOnly btnSupplierFillDue As New Button()
        Private ReadOnly btnSupplierNew As New Button()
        Private ReadOnly btnSupplierSave As New Button()
        Private ReadOnly dtpSupplierHistoryFrom As New DateTimePicker()
        Private ReadOnly dtpSupplierHistoryTo As New DateTimePicker()
        Private ReadOnly txtSupplierHistorySearch As New TextBox()
        Private ReadOnly btnRefreshSupplierHistory As New Button()
        Private ReadOnly dgvSupplierHistory As New DataGridView()
        Private ReadOnly lblSupplierHistoryCaption As New Label()

        Private _isBusy As Boolean

        Public Sub New(settlementService As SettlementService, customerService As CustomerService, supplierService As SupplierService)
            _settlementService = settlementService
            _customerService = customerService
            _supplierService = supplierService

            Text = "Healthwond Billing System - Settlements"
            StartPosition = FormStartPosition.CenterParent
            WindowState = FormWindowState.Maximized
            MinimumSize = New Size(1520, 920)
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
            lblStatus.ForeColor = ThemePalette.TextMuted
            lblStatus.TextAlign = ContentAlignment.MiddleLeft
            root.Controls.Add(lblStatus, 0, 2)

            Controls.Add(root)
        End Sub

        Private Function BuildHeaderPanel() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = ThemePalette.AppBackground}

            Dim title As New Label With {
                .Dock = DockStyle.Top,
                .Height = 40,
                .Text = "Settlements",
                .Font = New Font("Segoe UI Semibold", 24.0F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary
            }

            Dim subtitle As New Label With {
                .Dock = DockStyle.Fill,
                .Text = "Record customer collections and supplier payments with searchable settlement history and automatic outstanding-balance updates.",
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
            tabs.TabPages.Add(BuildCustomerTab())
            tabs.TabPages.Add(BuildSupplierTab())
            Return tabs
        End Function

        Private Function BuildCustomerTab() As TabPage
            Dim page As New TabPage("Customer Collections") With {.BackColor = ThemePalette.AppBackground}

            Dim root As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 2,
                .Padding = New Padding(8),
                .BackColor = ThemePalette.AppBackground
            }
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 86))
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            root.Controls.Add(BuildCustomerFilterPanel(), 0, 0)
            root.Controls.Add(BuildCustomerWorkspace(), 0, 1)

            page.Controls.Add(root)
            Return page
        End Function

        Private Function BuildSupplierTab() As TabPage
            Dim page As New TabPage("Supplier Payments") With {.BackColor = ThemePalette.AppBackground}

            Dim root As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 2,
                .Padding = New Padding(8),
                .BackColor = ThemePalette.AppBackground
            }
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 86))
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            root.Controls.Add(BuildSupplierFilterPanel(), 0, 0)
            root.Controls.Add(BuildSupplierWorkspace(), 0, 1)

            page.Controls.Add(root)
            Return page
        End Function

        Private Function BuildCustomerFilterPanel() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)

            Dim layout As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2
            }
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 140))

            layout.Controls.Add(CreateInputHost("Search customers", txtCustomerSearch), 0, 0)

            UiStyler.StylePrimaryButton(btnRefreshCustomers)
            btnRefreshCustomers.Text = "Refresh"
            btnRefreshCustomers.Width = 110
            btnRefreshCustomers.Margin = New Padding(0, 24, 0, 0)
            layout.Controls.Add(btnRefreshCustomers, 1, 0)

            panel.Controls.Add(layout)
            Return panel
        End Function

        Private Function BuildSupplierFilterPanel() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)

            Dim layout As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2
            }
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 140))

            layout.Controls.Add(CreateInputHost("Search suppliers", txtSupplierSearch), 0, 0)

            UiStyler.StylePrimaryButton(btnRefreshSuppliers)
            btnRefreshSuppliers.Text = "Refresh"
            btnRefreshSuppliers.Width = 110
            btnRefreshSuppliers.Margin = New Padding(0, 24, 0, 0)
            layout.Controls.Add(btnRefreshSuppliers, 1, 0)

            panel.Controls.Add(layout)
            Return panel
        End Function

        Private Function BuildCustomerWorkspace() As Control
            Dim split As New SplitContainer With {
                .Dock = DockStyle.Fill,
                .SplitterDistance = 430,
                .BackColor = ThemePalette.AppBackground
            }

            split.Panel1.Controls.Add(BuildCustomerGridCard())
            split.Panel2.Controls.Add(BuildCustomerDetailsCard())
            Return split
        End Function

        Private Function BuildSupplierWorkspace() As Control
            Dim split As New SplitContainer With {
                .Dock = DockStyle.Fill,
                .SplitterDistance = 430,
                .BackColor = ThemePalette.AppBackground
            }

            split.Panel1.Controls.Add(BuildSupplierGridCard())
            split.Panel2.Controls.Add(BuildSupplierDetailsCard())
            Return split
        End Function

        Private Function BuildCustomerGridCard() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)

            Dim title As New Label With {
                .Dock = DockStyle.Top,
                .Height = 30,
                .Text = "Customer dues",
                .Font = New Font("Segoe UI Semibold", 14.0F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary
            }

            dgvCustomers.Dock = DockStyle.Fill
            panel.Controls.Add(dgvCustomers)
            panel.Controls.Add(title)
            Return panel
        End Function

        Private Function BuildSupplierGridCard() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)

            Dim title As New Label With {
                .Dock = DockStyle.Top,
                .Height = 30,
                .Text = "Supplier payables",
                .Font = New Font("Segoe UI Semibold", 14.0F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary
            }

            dgvSuppliers.Dock = DockStyle.Fill
            panel.Controls.Add(dgvSuppliers)
            panel.Controls.Add(title)
            Return panel
        End Function

        Private Function BuildCustomerDetailsCard() As Control
            Dim root As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 4
            }
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 114))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 196))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 86))
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            root.Controls.Add(BuildCustomerSummaryCard(), 0, 0)
            root.Controls.Add(BuildCustomerEntryCard(), 0, 1)
            root.Controls.Add(BuildCustomerHistoryFilterCard(), 0, 2)
            root.Controls.Add(BuildCustomerHistoryCard(), 0, 3)

            Return root
        End Function

        Private Function BuildSupplierDetailsCard() As Control
            Dim root As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 4
            }
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 114))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 196))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 86))
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            root.Controls.Add(BuildSupplierSummaryCard(), 0, 0)
            root.Controls.Add(BuildSupplierEntryCard(), 0, 1)
            root.Controls.Add(BuildSupplierHistoryFilterCard(), 0, 2)
            root.Controls.Add(BuildSupplierHistoryCard(), 0, 3)

            Return root
        End Function

        Private Function BuildCustomerSummaryCard() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)

            Dim layout As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 3,
                .RowCount = 2
            }
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 44.0F))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 28.0F))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 28.0F))
            layout.RowStyles.Add(New RowStyle(SizeType.Absolute, 30))
            layout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            lblSelectedCustomer.Dock = DockStyle.Fill
            lblSelectedCustomer.Font = New Font("Segoe UI Semibold", 12.5F, FontStyle.Bold)
            lblSelectedCustomer.ForeColor = ThemePalette.TextPrimary
            lblSelectedCustomer.Text = "Select a customer to start a collection entry."

            layout.Controls.Add(lblSelectedCustomer, 0, 0)
            layout.SetColumnSpan(lblSelectedCustomer, 3)
            layout.Controls.Add(CreateSummaryValuePanel("Current Outstanding", lblCustomerOutstandingValue), 1, 1)
            layout.Controls.Add(CreateSummaryValuePanel("Balance After Collection", lblCustomerBalanceAfterValue), 2, 1)

            panel.Controls.Add(layout)
            Return panel
        End Function

        Private Function BuildSupplierSummaryCard() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)

            Dim layout As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 3,
                .RowCount = 2
            }
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 44.0F))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 28.0F))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 28.0F))
            layout.RowStyles.Add(New RowStyle(SizeType.Absolute, 30))
            layout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            lblSelectedSupplier.Dock = DockStyle.Fill
            lblSelectedSupplier.Font = New Font("Segoe UI Semibold", 12.5F, FontStyle.Bold)
            lblSelectedSupplier.ForeColor = ThemePalette.TextPrimary
            lblSelectedSupplier.Text = "Select a supplier to start a payment entry."

            layout.Controls.Add(lblSelectedSupplier, 0, 0)
            layout.SetColumnSpan(lblSelectedSupplier, 3)
            layout.Controls.Add(CreateSummaryValuePanel("Current Outstanding", lblSupplierOutstandingValue), 1, 1)
            layout.Controls.Add(CreateSummaryValuePanel("Balance After Payment", lblSupplierBalanceAfterValue), 2, 1)

            panel.Controls.Add(layout)
            Return panel
        End Function

        Private Function BuildCustomerEntryCard() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)

            Dim layout As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 4,
                .RowCount = 2
            }
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 26.0F))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 170))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 180))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 34.0F))
            layout.RowStyles.Add(New RowStyle(SizeType.Absolute, 74))
            layout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            layout.Controls.Add(CreateInputHost("Receipt Number", txtReceiptNumber), 0, 0)
            layout.Controls.Add(CreateInputHost("Collection Date", dtpCustomerPaymentDate), 1, 0)
            layout.Controls.Add(CreateInputHost("Payment Mode", cboCustomerPaymentMode), 2, 0)
            layout.Controls.Add(CreateInputHost("Reference Number", txtCustomerReference), 3, 0)

            Dim lowerLayout As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 3
            }
            lowerLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 220))
            lowerLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            lowerLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 330))

            lowerLayout.Controls.Add(CreateInputHost("Collection Amount", nudCustomerPaymentAmount), 0, 0)

            txtCustomerNotes.Multiline = True
            txtCustomerNotes.ScrollBars = ScrollBars.Vertical
            lowerLayout.Controls.Add(CreateInputHost("Notes", txtCustomerNotes), 1, 0)

            Dim buttonFlow As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.LeftToRight,
                .WrapContents = False,
                .BackColor = Color.Transparent,
                .Padding = New Padding(0, 22, 0, 0)
            }

            UiStyler.StyleSecondaryButton(btnCustomerFillDue)
            btnCustomerFillDue.Text = "Use Full Due"
            btnCustomerFillDue.Width = 110

            UiStyler.StyleSecondaryButton(btnCustomerNew)
            btnCustomerNew.Text = "New"
            btnCustomerNew.Width = 90

            UiStyler.StylePrimaryButton(btnCustomerSave)
            btnCustomerSave.Text = "Save Collection"
            btnCustomerSave.Width = 130

            buttonFlow.Controls.Add(btnCustomerFillDue)
            buttonFlow.Controls.Add(btnCustomerNew)
            buttonFlow.Controls.Add(btnCustomerSave)
            lowerLayout.Controls.Add(buttonFlow, 2, 0)

            layout.Controls.Add(lowerLayout, 0, 1)
            layout.SetColumnSpan(lowerLayout, 4)

            panel.Controls.Add(layout)
            Return panel
        End Function

        Private Function BuildSupplierEntryCard() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)

            Dim layout As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 4,
                .RowCount = 2
            }
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 26.0F))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 170))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 180))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 34.0F))
            layout.RowStyles.Add(New RowStyle(SizeType.Absolute, 74))
            layout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            layout.Controls.Add(CreateInputHost("Payment Number", txtSupplierPaymentNumber), 0, 0)
            layout.Controls.Add(CreateInputHost("Payment Date", dtpSupplierPaymentDate), 1, 0)
            layout.Controls.Add(CreateInputHost("Payment Mode", cboSupplierPaymentMode), 2, 0)
            layout.Controls.Add(CreateInputHost("Reference Number", txtSupplierReference), 3, 0)

            Dim lowerLayout As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 3
            }
            lowerLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 220))
            lowerLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            lowerLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 330))

            lowerLayout.Controls.Add(CreateInputHost("Payment Amount", nudSupplierPaymentAmount), 0, 0)

            txtSupplierNotes.Multiline = True
            txtSupplierNotes.ScrollBars = ScrollBars.Vertical
            lowerLayout.Controls.Add(CreateInputHost("Notes", txtSupplierNotes), 1, 0)

            Dim buttonFlow As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.LeftToRight,
                .WrapContents = False,
                .BackColor = Color.Transparent,
                .Padding = New Padding(0, 22, 0, 0)
            }

            UiStyler.StyleSecondaryButton(btnSupplierFillDue)
            btnSupplierFillDue.Text = "Use Full Due"
            btnSupplierFillDue.Width = 110

            UiStyler.StyleSecondaryButton(btnSupplierNew)
            btnSupplierNew.Text = "New"
            btnSupplierNew.Width = 90

            UiStyler.StylePrimaryButton(btnSupplierSave)
            btnSupplierSave.Text = "Save Payment"
            btnSupplierSave.Width = 130

            buttonFlow.Controls.Add(btnSupplierFillDue)
            buttonFlow.Controls.Add(btnSupplierNew)
            buttonFlow.Controls.Add(btnSupplierSave)
            lowerLayout.Controls.Add(buttonFlow, 2, 0)

            layout.Controls.Add(lowerLayout, 0, 1)
            layout.SetColumnSpan(lowerLayout, 4)

            panel.Controls.Add(layout)
            Return panel
        End Function

        Private Function BuildCustomerHistoryFilterCard() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)

            Dim layout As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 4
            }
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 170))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 170))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 140))

            layout.Controls.Add(CreateInputHost("History From", dtpCustomerHistoryFrom), 0, 0)
            layout.Controls.Add(CreateInputHost("History To", dtpCustomerHistoryTo), 1, 0)
            layout.Controls.Add(CreateInputHost("Search history", txtCustomerHistorySearch), 2, 0)

            UiStyler.StyleSecondaryButton(btnRefreshCustomerHistory)
            btnRefreshCustomerHistory.Text = "Refresh"
            btnRefreshCustomerHistory.Width = 110
            btnRefreshCustomerHistory.Margin = New Padding(0, 24, 0, 0)
            layout.Controls.Add(btnRefreshCustomerHistory, 3, 0)

            panel.Controls.Add(layout)
            Return panel
        End Function

        Private Function BuildSupplierHistoryFilterCard() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)

            Dim layout As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 4
            }
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 170))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 170))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 140))

            layout.Controls.Add(CreateInputHost("History From", dtpSupplierHistoryFrom), 0, 0)
            layout.Controls.Add(CreateInputHost("History To", dtpSupplierHistoryTo), 1, 0)
            layout.Controls.Add(CreateInputHost("Search history", txtSupplierHistorySearch), 2, 0)

            UiStyler.StyleSecondaryButton(btnRefreshSupplierHistory)
            btnRefreshSupplierHistory.Text = "Refresh"
            btnRefreshSupplierHistory.Width = 110
            btnRefreshSupplierHistory.Margin = New Padding(0, 24, 0, 0)
            layout.Controls.Add(btnRefreshSupplierHistory, 3, 0)

            panel.Controls.Add(layout)
            Return panel
        End Function

        Private Function BuildCustomerHistoryCard() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)

            lblCustomerHistoryCaption.Dock = DockStyle.Top
            lblCustomerHistoryCaption.Height = 28
            lblCustomerHistoryCaption.Font = New Font("Segoe UI Semibold", 12.0F, FontStyle.Bold)
            lblCustomerHistoryCaption.ForeColor = ThemePalette.TextPrimary
            lblCustomerHistoryCaption.Text = "Collection history"

            dgvCustomerHistory.Dock = DockStyle.Fill

            panel.Controls.Add(dgvCustomerHistory)
            panel.Controls.Add(lblCustomerHistoryCaption)
            Return panel
        End Function

        Private Function BuildSupplierHistoryCard() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)

            lblSupplierHistoryCaption.Dock = DockStyle.Top
            lblSupplierHistoryCaption.Height = 28
            lblSupplierHistoryCaption.Font = New Font("Segoe UI Semibold", 12.0F, FontStyle.Bold)
            lblSupplierHistoryCaption.ForeColor = ThemePalette.TextPrimary
            lblSupplierHistoryCaption.Text = "Payment history"

            dgvSupplierHistory.Dock = DockStyle.Fill

            panel.Controls.Add(dgvSupplierHistory)
            panel.Controls.Add(lblSupplierHistoryCaption)
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

        Private Function CreateSummaryValuePanel(caption As String, valueLabel As Label) As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill}

            Dim captionLabel As New Label With {
                .Dock = DockStyle.Top,
                .Height = 22,
                .Text = caption,
                .Font = New Font("Segoe UI", 9.5F, FontStyle.Regular),
                .ForeColor = ThemePalette.TextMuted
            }

            valueLabel.Dock = DockStyle.Fill
            valueLabel.Font = New Font("Segoe UI Semibold", 17.0F, FontStyle.Bold)
            valueLabel.ForeColor = ThemePalette.TextPrimary
            valueLabel.TextAlign = ContentAlignment.MiddleLeft
            valueLabel.Text = "Rs. 0.00"

            panel.Controls.Add(valueLabel)
            panel.Controls.Add(captionLabel)
            Return panel
        End Function

        Private Sub ConfigureControls()
            For Each editorTextBox As TextBox In New TextBox() {txtCustomerSearch, txtReceiptNumber, txtCustomerReference, txtCustomerNotes, txtCustomerHistorySearch, txtSupplierSearch, txtSupplierPaymentNumber, txtSupplierReference, txtSupplierNotes, txtSupplierHistorySearch}
                editorTextBox.BorderStyle = BorderStyle.FixedSingle
                UiStyler.StyleInput(editorTextBox)
            Next

            ConfigureComboBox(cboCustomerPaymentMode)
            ConfigureComboBox(cboSupplierPaymentMode)
            PopulatePaymentModes(cboCustomerPaymentMode)
            PopulatePaymentModes(cboSupplierPaymentMode)

            ConfigureDatePicker(dtpCustomerPaymentDate)
            ConfigureDatePicker(dtpCustomerHistoryFrom)
            ConfigureDatePicker(dtpCustomerHistoryTo)
            ConfigureDatePicker(dtpSupplierPaymentDate)
            ConfigureDatePicker(dtpSupplierHistoryFrom)
            ConfigureDatePicker(dtpSupplierHistoryTo)

            ConfigureCurrencyInput(nudCustomerPaymentAmount)
            ConfigureCurrencyInput(nudSupplierPaymentAmount)

            txtCustomerNotes.Height = 72
            txtSupplierNotes.Height = 72
        End Sub

        Private Sub ConfigureComboBox(comboBox As ComboBox)
            comboBox.DropDownStyle = ComboBoxStyle.DropDownList
            comboBox.Font = New Font("Segoe UI", 10.0F, FontStyle.Regular)
            comboBox.BackColor = Color.White
        End Sub

        Private Sub PopulatePaymentModes(comboBox As ComboBox)
            comboBox.Items.Clear()
            comboBox.Items.AddRange(New Object() {"Cash", "UPI", "Bank Transfer", "Cheque", "Card", "Adjustment"})
            comboBox.SelectedIndex = 0
        End Sub

        Private Sub ConfigureDatePicker(datePicker As DateTimePicker)
            datePicker.Format = DateTimePickerFormat.Custom
            datePicker.CustomFormat = "dd MMM yyyy"
        End Sub

        Private Sub ConfigureCurrencyInput(inputControl As NumericUpDown)
            inputControl.DecimalPlaces = 2
            inputControl.Maximum = 100000000D
            inputControl.Minimum = 0D
            inputControl.Increment = 50D
            inputControl.ThousandsSeparator = True
            inputControl.Font = New Font("Segoe UI", 10.5F, FontStyle.Regular)
        End Sub

        Private Sub ConfigureGrids()
            UiStyler.StyleDataGrid(dgvCustomers)
            dgvCustomers.DataSource = _customerRows
            dgvCustomers.Columns.Clear()
            dgvCustomers.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "CustomerName", .HeaderText = "Customer", .FillWeight = 160.0F})
            dgvCustomers.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "Phone", .HeaderText = "Phone", .FillWeight = 82.0F})
            dgvCustomers.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "OutstandingBalance", .HeaderText = "Outstanding", .FillWeight = 78.0F, .DefaultCellStyle = New DataGridViewCellStyle With {.Format = "N2"}})

            UiStyler.StyleDataGrid(dgvSuppliers)
            dgvSuppliers.DataSource = _supplierRows
            dgvSuppliers.Columns.Clear()
            dgvSuppliers.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "SupplierName", .HeaderText = "Supplier", .FillWeight = 160.0F})
            dgvSuppliers.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "Phone", .HeaderText = "Phone", .FillWeight = 82.0F})
            dgvSuppliers.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "OutstandingBalance", .HeaderText = "Outstanding", .FillWeight = 78.0F, .DefaultCellStyle = New DataGridViewCellStyle With {.Format = "N2"}})

            UiStyler.StyleDataGrid(dgvCustomerHistory)
            dgvCustomerHistory.DataSource = _customerPaymentRows
            dgvCustomerHistory.Columns.Clear()
            dgvCustomerHistory.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "ReceiptNumber", .HeaderText = "Receipt No.", .FillWeight = 110.0F})
            dgvCustomerHistory.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "PaymentDate", .HeaderText = "Date", .FillWeight = 85.0F, .DefaultCellStyle = New DataGridViewCellStyle With {.Format = "dd MMM yyyy"}})
            dgvCustomerHistory.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "PaymentMode", .HeaderText = "Mode", .FillWeight = 82.0F})
            dgvCustomerHistory.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "ReferenceNumber", .HeaderText = "Reference", .FillWeight = 95.0F})
            dgvCustomerHistory.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "Amount", .HeaderText = "Amount", .FillWeight = 82.0F, .DefaultCellStyle = New DataGridViewCellStyle With {.Format = "N2"}})
            dgvCustomerHistory.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "BalanceAfterPayment", .HeaderText = "Balance After", .FillWeight = 94.0F, .DefaultCellStyle = New DataGridViewCellStyle With {.Format = "N2"}})

            UiStyler.StyleDataGrid(dgvSupplierHistory)
            dgvSupplierHistory.DataSource = _supplierPaymentRows
            dgvSupplierHistory.Columns.Clear()
            dgvSupplierHistory.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "PaymentNumber", .HeaderText = "Payment No.", .FillWeight = 110.0F})
            dgvSupplierHistory.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "PaymentDate", .HeaderText = "Date", .FillWeight = 85.0F, .DefaultCellStyle = New DataGridViewCellStyle With {.Format = "dd MMM yyyy"}})
            dgvSupplierHistory.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "PaymentMode", .HeaderText = "Mode", .FillWeight = 82.0F})
            dgvSupplierHistory.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "ReferenceNumber", .HeaderText = "Reference", .FillWeight = 95.0F})
            dgvSupplierHistory.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "Amount", .HeaderText = "Amount", .FillWeight = 82.0F, .DefaultCellStyle = New DataGridViewCellStyle With {.Format = "N2"}})
            dgvSupplierHistory.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "BalanceAfterPayment", .HeaderText = "Balance After", .FillWeight = 94.0F, .DefaultCellStyle = New DataGridViewCellStyle With {.Format = "N2"}})
        End Sub

        Private Sub WireEvents()
            AddHandler Load, AddressOf FrmSettlements_Load
            AddHandler btnRefreshCustomers.Click, AddressOf btnRefreshCustomers_Click
            AddHandler btnRefreshSuppliers.Click, AddressOf btnRefreshSuppliers_Click
            AddHandler dgvCustomers.SelectionChanged, AddressOf dgvCustomers_SelectionChanged
            AddHandler dgvSuppliers.SelectionChanged, AddressOf dgvSuppliers_SelectionChanged
            AddHandler btnCustomerFillDue.Click, AddressOf btnCustomerFillDue_Click
            AddHandler btnSupplierFillDue.Click, AddressOf btnSupplierFillDue_Click
            AddHandler btnCustomerNew.Click, AddressOf btnCustomerNew_Click
            AddHandler btnSupplierNew.Click, AddressOf btnSupplierNew_Click
            AddHandler btnCustomerSave.Click, AddressOf btnCustomerSave_Click
            AddHandler btnSupplierSave.Click, AddressOf btnSupplierSave_Click
            AddHandler btnRefreshCustomerHistory.Click, AddressOf btnRefreshCustomerHistory_Click
            AddHandler btnRefreshSupplierHistory.Click, AddressOf btnRefreshSupplierHistory_Click
            AddHandler nudCustomerPaymentAmount.ValueChanged, AddressOf nudCustomerPaymentAmount_ValueChanged
            AddHandler nudSupplierPaymentAmount.ValueChanged, AddressOf nudSupplierPaymentAmount_ValueChanged
            AddHandler dtpCustomerPaymentDate.ValueChanged, AddressOf dtpCustomerPaymentDate_ValueChanged
            AddHandler dtpSupplierPaymentDate.ValueChanged, AddressOf dtpSupplierPaymentDate_ValueChanged
            AddHandler txtCustomerSearch.KeyDown, AddressOf txtCustomerSearch_KeyDown
            AddHandler txtSupplierSearch.KeyDown, AddressOf txtSupplierSearch_KeyDown
            AddHandler txtCustomerHistorySearch.KeyDown, AddressOf txtCustomerHistorySearch_KeyDown
            AddHandler txtSupplierHistorySearch.KeyDown, AddressOf txtSupplierHistorySearch_KeyDown
        End Sub

        Private Async Sub FrmSettlements_Load(sender As Object, e As EventArgs)
            Dim monthStart As New DateTime(DateTime.Today.Year, DateTime.Today.Month, 1)
            dtpCustomerHistoryFrom.Value = monthStart
            dtpCustomerHistoryTo.Value = DateTime.Today
            dtpSupplierHistoryFrom.Value = monthStart
            dtpSupplierHistoryTo.Value = DateTime.Today
            dtpCustomerPaymentDate.Value = DateTime.Today
            dtpSupplierPaymentDate.Value = DateTime.Today

            Await LoadCustomersAsync()
            Await LoadSuppliersAsync()
        End Sub

        Private Async Sub btnRefreshCustomers_Click(sender As Object, e As EventArgs)
            Await LoadCustomersAsync(GetSelectedCustomerId())
        End Sub

        Private Async Sub btnRefreshSuppliers_Click(sender As Object, e As EventArgs)
            Await LoadSuppliersAsync(GetSelectedSupplierId())
        End Sub

        Private Async Function LoadCustomersAsync(Optional preferredCustomerId As Integer = 0) As Task
            SetBusy(True, "Loading customers with outstanding balances...")

            Try
                Dim customers As List(Of CustomerRecord) =
                    (Await _customerService.SearchAsync(txtCustomerSearch.Text.Trim())).
                    OrderByDescending(Function(row) row.OutstandingBalance).
                    ThenBy(Function(row) row.CustomerName).
                    ToList()

                _customerRows.Clear()
                For Each customer As CustomerRecord In customers
                    _customerRows.Add(customer)
                Next

                If _customerRows.Count = 0 Then
                    ResetCustomerSelection()
                    ShowStatus("No customers matched the current search.", False)
                    Return
                End If

                SelectCustomerRow(preferredCustomerId)
            Catch ex As Exception
                AppLogger.Error("Customer settlement list load failed.", ex)
                ShowStatus("Customers could not be loaded for settlements.", True)
                Return
            Finally
                SetBusy(False)
            End Try

            Await HandleCustomerSelectionChangedAsync()
        End Function

        Private Async Function LoadSuppliersAsync(Optional preferredSupplierId As Integer = 0) As Task
            SetBusy(True, "Loading suppliers with outstanding balances...")

            Try
                Dim suppliers As List(Of SupplierRecord) =
                    (Await _supplierService.SearchAsync(txtSupplierSearch.Text.Trim())).
                    OrderByDescending(Function(row) row.OutstandingBalance).
                    ThenBy(Function(row) row.SupplierName).
                    ToList()

                _supplierRows.Clear()
                For Each supplier As SupplierRecord In suppliers
                    _supplierRows.Add(supplier)
                Next

                If _supplierRows.Count = 0 Then
                    ResetSupplierSelection()
                    ShowStatus("No suppliers matched the current search.", False)
                    Return
                End If

                SelectSupplierRow(preferredSupplierId)
            Catch ex As Exception
                AppLogger.Error("Supplier settlement list load failed.", ex)
                ShowStatus("Suppliers could not be loaded for settlements.", True)
                Return
            Finally
                SetBusy(False)
            End Try

            Await HandleSupplierSelectionChangedAsync()
        End Function

        Private Async Sub dgvCustomers_SelectionChanged(sender As Object, e As EventArgs)
            If _isBusy Then
                Return
            End If

            Await HandleCustomerSelectionChangedAsync()
        End Sub

        Private Async Sub dgvSuppliers_SelectionChanged(sender As Object, e As EventArgs)
            If _isBusy Then
                Return
            End If

            Await HandleSupplierSelectionChangedAsync()
        End Sub

        Private Async Function HandleCustomerSelectionChangedAsync() As Task
            Dim customer As CustomerRecord = GetSelectedCustomer()
            If customer Is Nothing Then
                ResetCustomerSelection()
                Return
            End If

            lblSelectedCustomer.Text = $"{customer.CustomerName} | Phone {If(customer.Phone, String.Empty)}"
            Await PrepareCustomerEntryAsync()
            Await LoadCustomerHistoryAsync(customer.Id)
        End Function

        Private Async Function HandleSupplierSelectionChangedAsync() As Task
            Dim supplier As SupplierRecord = GetSelectedSupplier()
            If supplier Is Nothing Then
                ResetSupplierSelection()
                Return
            End If

            lblSelectedSupplier.Text = $"{supplier.SupplierName} | Phone {If(supplier.Phone, String.Empty)}"
            Await PrepareSupplierEntryAsync()
            Await LoadSupplierHistoryAsync(supplier.Id)
        End Function

        Private Async Function PrepareCustomerEntryAsync() As Task
            txtReceiptNumber.Text = Await _settlementService.GenerateNextCustomerReceiptNumberAsync(dtpCustomerPaymentDate.Value.Date)
            If cboCustomerPaymentMode.Items.Count > 0 AndAlso cboCustomerPaymentMode.SelectedIndex < 0 Then
                cboCustomerPaymentMode.SelectedIndex = 0
            End If

            txtCustomerReference.Clear()
            txtCustomerNotes.Clear()
            nudCustomerPaymentAmount.Value = 0D
            UpdateCustomerBalancePreview()
        End Function

        Private Async Function PrepareSupplierEntryAsync() As Task
            txtSupplierPaymentNumber.Text = Await _settlementService.GenerateNextSupplierPaymentNumberAsync(dtpSupplierPaymentDate.Value.Date)
            If cboSupplierPaymentMode.Items.Count > 0 AndAlso cboSupplierPaymentMode.SelectedIndex < 0 Then
                cboSupplierPaymentMode.SelectedIndex = 0
            End If

            txtSupplierReference.Clear()
            txtSupplierNotes.Clear()
            nudSupplierPaymentAmount.Value = 0D
            UpdateSupplierBalancePreview()
        End Function

        Private Async Function LoadCustomerHistoryAsync(customerId As Integer) As Task
            If dtpCustomerHistoryFrom.Value.Date > dtpCustomerHistoryTo.Value.Date Then
                ShowStatus("Customer history From date cannot be later than To date.", True)
                Return
            End If

            SetBusy(True, "Loading customer collection history...")

            Try
                Dim rows As List(Of CustomerPaymentHistoryRow) =
                    Await _settlementService.SearchCustomerPaymentsAsync(customerId, dtpCustomerHistoryFrom.Value.Date, dtpCustomerHistoryTo.Value.Date, txtCustomerHistorySearch.Text.Trim())

                _customerPaymentRows.Clear()
                For Each row As CustomerPaymentHistoryRow In rows
                    _customerPaymentRows.Add(row)
                Next

                lblCustomerHistoryCaption.Text = $"Collection history ({rows.Count:N0})"
            Catch ex As Exception
                AppLogger.Error("Customer collection history load failed.", ex)
                ShowStatus("Customer collection history could not be loaded.", True)
            Finally
                SetBusy(False)
            End Try
        End Function

        Private Async Function LoadSupplierHistoryAsync(supplierId As Integer) As Task
            If dtpSupplierHistoryFrom.Value.Date > dtpSupplierHistoryTo.Value.Date Then
                ShowStatus("Supplier history From date cannot be later than To date.", True)
                Return
            End If

            SetBusy(True, "Loading supplier payment history...")

            Try
                Dim rows As List(Of SupplierPaymentHistoryRow) =
                    Await _settlementService.SearchSupplierPaymentsAsync(supplierId, dtpSupplierHistoryFrom.Value.Date, dtpSupplierHistoryTo.Value.Date, txtSupplierHistorySearch.Text.Trim())

                _supplierPaymentRows.Clear()
                For Each row As SupplierPaymentHistoryRow In rows
                    _supplierPaymentRows.Add(row)
                Next

                lblSupplierHistoryCaption.Text = $"Payment history ({rows.Count:N0})"
            Catch ex As Exception
                AppLogger.Error("Supplier payment history load failed.", ex)
                ShowStatus("Supplier payment history could not be loaded.", True)
            Finally
                SetBusy(False)
            End Try
        End Function

        Private Async Sub btnCustomerNew_Click(sender As Object, e As EventArgs)
            Await PrepareCustomerEntryAsync()
            ShowStatus("Ready for a new customer collection entry.", False)
        End Sub

        Private Async Sub btnSupplierNew_Click(sender As Object, e As EventArgs)
            Await PrepareSupplierEntryAsync()
            ShowStatus("Ready for a new supplier payment entry.", False)
        End Sub

        Private Sub btnCustomerFillDue_Click(sender As Object, e As EventArgs)
            Dim customer As CustomerRecord = GetSelectedCustomer()
            If customer Is Nothing Then
                ShowStatus("Select a customer before entering collection amount.", True)
                Return
            End If

            nudCustomerPaymentAmount.Value = Math.Min(nudCustomerPaymentAmount.Maximum, customer.OutstandingBalance)
        End Sub

        Private Sub btnSupplierFillDue_Click(sender As Object, e As EventArgs)
            Dim supplier As SupplierRecord = GetSelectedSupplier()
            If supplier Is Nothing Then
                ShowStatus("Select a supplier before entering payment amount.", True)
                Return
            End If

            nudSupplierPaymentAmount.Value = Math.Min(nudSupplierPaymentAmount.Maximum, supplier.OutstandingBalance)
        End Sub

        Private Async Sub btnCustomerSave_Click(sender As Object, e As EventArgs)
            Dim customer As CustomerRecord = GetSelectedCustomer()
            If customer Is Nothing Then
                ShowStatus("Select a customer before saving the collection.", True)
                Return
            End If

            Dim draft As New CustomerPaymentDraft With {
                .ReceiptNumber = txtReceiptNumber.Text,
                .CustomerId = customer.Id,
                .CustomerName = customer.CustomerName,
                .PaymentDate = dtpCustomerPaymentDate.Value.Date,
                .PaymentMode = Convert.ToString(cboCustomerPaymentMode.SelectedItem),
                .ReferenceNumber = txtCustomerReference.Text,
                .Amount = nudCustomerPaymentAmount.Value,
                .OutstandingBeforePayment = customer.OutstandingBalance,
                .Notes = txtCustomerNotes.Text
            }

            SetBusy(True, "Saving customer collection...")
            Dim result As EntityOperationResult = Await _settlementService.SaveCustomerPaymentAsync(draft, If(SessionManager.CurrentUser Is Nothing, 0, SessionManager.CurrentUser.Id))
            SetBusy(False)
            ShowStatus(result.Message, Not result.IsSuccess)

            If result.IsSuccess Then
                Await LoadCustomersAsync(customer.Id)
            End If
        End Sub

        Private Async Sub btnSupplierSave_Click(sender As Object, e As EventArgs)
            Dim supplier As SupplierRecord = GetSelectedSupplier()
            If supplier Is Nothing Then
                ShowStatus("Select a supplier before saving the payment.", True)
                Return
            End If

            Dim draft As New SupplierPaymentDraft With {
                .PaymentNumber = txtSupplierPaymentNumber.Text,
                .SupplierId = supplier.Id,
                .SupplierName = supplier.SupplierName,
                .PaymentDate = dtpSupplierPaymentDate.Value.Date,
                .PaymentMode = Convert.ToString(cboSupplierPaymentMode.SelectedItem),
                .ReferenceNumber = txtSupplierReference.Text,
                .Amount = nudSupplierPaymentAmount.Value,
                .OutstandingBeforePayment = supplier.OutstandingBalance,
                .Notes = txtSupplierNotes.Text
            }

            SetBusy(True, "Saving supplier payment...")
            Dim result As EntityOperationResult = Await _settlementService.SaveSupplierPaymentAsync(draft, If(SessionManager.CurrentUser Is Nothing, 0, SessionManager.CurrentUser.Id))
            SetBusy(False)
            ShowStatus(result.Message, Not result.IsSuccess)

            If result.IsSuccess Then
                Await LoadSuppliersAsync(supplier.Id)
            End If
        End Sub

        Private Async Sub btnRefreshCustomerHistory_Click(sender As Object, e As EventArgs)
            Await LoadCustomerHistoryAsync(GetSelectedCustomerId())
        End Sub

        Private Async Sub btnRefreshSupplierHistory_Click(sender As Object, e As EventArgs)
            Await LoadSupplierHistoryAsync(GetSelectedSupplierId())
        End Sub

        Private Sub nudCustomerPaymentAmount_ValueChanged(sender As Object, e As EventArgs)
            UpdateCustomerBalancePreview()
        End Sub

        Private Sub nudSupplierPaymentAmount_ValueChanged(sender As Object, e As EventArgs)
            UpdateSupplierBalancePreview()
        End Sub

        Private Async Sub dtpCustomerPaymentDate_ValueChanged(sender As Object, e As EventArgs)
            If Not _isBusy Then
                txtReceiptNumber.Text = Await _settlementService.GenerateNextCustomerReceiptNumberAsync(dtpCustomerPaymentDate.Value.Date)
            End If
        End Sub

        Private Async Sub dtpSupplierPaymentDate_ValueChanged(sender As Object, e As EventArgs)
            If Not _isBusy Then
                txtSupplierPaymentNumber.Text = Await _settlementService.GenerateNextSupplierPaymentNumberAsync(dtpSupplierPaymentDate.Value.Date)
            End If
        End Sub

        Private Sub UpdateCustomerBalancePreview()
            Dim customer As CustomerRecord = GetSelectedCustomer()
            Dim currentOutstanding As Decimal = If(customer Is Nothing, 0D, customer.OutstandingBalance)
            Dim collectionAmount As Decimal = nudCustomerPaymentAmount.Value
            Dim balanceAfter As Decimal = Decimal.Round(Math.Max(0D, currentOutstanding - collectionAmount), 2, MidpointRounding.AwayFromZero)
            lblCustomerOutstandingValue.Text = $"Rs. {currentOutstanding:N2}"
            lblCustomerBalanceAfterValue.Text = $"Rs. {balanceAfter:N2}"
        End Sub

        Private Sub UpdateSupplierBalancePreview()
            Dim supplier As SupplierRecord = GetSelectedSupplier()
            Dim currentOutstanding As Decimal = If(supplier Is Nothing, 0D, supplier.OutstandingBalance)
            Dim paymentAmount As Decimal = nudSupplierPaymentAmount.Value
            Dim balanceAfter As Decimal = Decimal.Round(Math.Max(0D, currentOutstanding - paymentAmount), 2, MidpointRounding.AwayFromZero)
            lblSupplierOutstandingValue.Text = $"Rs. {currentOutstanding:N2}"
            lblSupplierBalanceAfterValue.Text = $"Rs. {balanceAfter:N2}"
        End Sub

        Private Function GetSelectedCustomer() As CustomerRecord
            If dgvCustomers.CurrentRow Is Nothing Then
                Return Nothing
            End If

            Return TryCast(dgvCustomers.CurrentRow.DataBoundItem, CustomerRecord)
        End Function

        Private Function GetSelectedSupplier() As SupplierRecord
            If dgvSuppliers.CurrentRow Is Nothing Then
                Return Nothing
            End If

            Return TryCast(dgvSuppliers.CurrentRow.DataBoundItem, SupplierRecord)
        End Function

        Private Function GetSelectedCustomerId() As Integer
            Dim customer As CustomerRecord = GetSelectedCustomer()
            If customer Is Nothing Then
                Return 0
            End If

            Return customer.Id
        End Function

        Private Function GetSelectedSupplierId() As Integer
            Dim supplier As SupplierRecord = GetSelectedSupplier()
            If supplier Is Nothing Then
                Return 0
            End If

            Return supplier.Id
        End Function

        Private Sub SelectCustomerRow(preferredCustomerId As Integer)
            If dgvCustomers.Rows.Count = 0 Then
                Return
            End If

            dgvCustomers.ClearSelection()
            If preferredCustomerId > 0 Then
                For Each row As DataGridViewRow In dgvCustomers.Rows
                    Dim customer As CustomerRecord = TryCast(row.DataBoundItem, CustomerRecord)
                    If customer IsNot Nothing AndAlso customer.Id = preferredCustomerId Then
                        row.Selected = True
                        dgvCustomers.CurrentCell = row.Cells(0)
                        Return
                    End If
                Next
            End If

            dgvCustomers.Rows(0).Selected = True
            dgvCustomers.CurrentCell = dgvCustomers.Rows(0).Cells(0)
        End Sub

        Private Sub SelectSupplierRow(preferredSupplierId As Integer)
            If dgvSuppliers.Rows.Count = 0 Then
                Return
            End If

            dgvSuppliers.ClearSelection()
            If preferredSupplierId > 0 Then
                For Each row As DataGridViewRow In dgvSuppliers.Rows
                    Dim supplier As SupplierRecord = TryCast(row.DataBoundItem, SupplierRecord)
                    If supplier IsNot Nothing AndAlso supplier.Id = preferredSupplierId Then
                        row.Selected = True
                        dgvSuppliers.CurrentCell = row.Cells(0)
                        Return
                    End If
                Next
            End If

            dgvSuppliers.Rows(0).Selected = True
            dgvSuppliers.CurrentCell = dgvSuppliers.Rows(0).Cells(0)
        End Sub

        Private Sub ResetCustomerSelection()
            lblSelectedCustomer.Text = "Select a customer to start a collection entry."
            lblCustomerHistoryCaption.Text = "Collection history"
            txtReceiptNumber.Clear()
            txtCustomerReference.Clear()
            txtCustomerNotes.Clear()
            nudCustomerPaymentAmount.Value = 0D
            _customerPaymentRows.Clear()
            UpdateCustomerBalancePreview()
        End Sub

        Private Sub ResetSupplierSelection()
            lblSelectedSupplier.Text = "Select a supplier to start a payment entry."
            lblSupplierHistoryCaption.Text = "Payment history"
            txtSupplierPaymentNumber.Clear()
            txtSupplierReference.Clear()
            txtSupplierNotes.Clear()
            nudSupplierPaymentAmount.Value = 0D
            _supplierPaymentRows.Clear()
            UpdateSupplierBalancePreview()
        End Sub

        Private Async Sub txtCustomerSearch_KeyDown(sender As Object, e As KeyEventArgs)
            If e.KeyCode = Keys.Enter Then
                e.SuppressKeyPress = True
                Await LoadCustomersAsync(GetSelectedCustomerId())
            End If
        End Sub

        Private Async Sub txtSupplierSearch_KeyDown(sender As Object, e As KeyEventArgs)
            If e.KeyCode = Keys.Enter Then
                e.SuppressKeyPress = True
                Await LoadSuppliersAsync(GetSelectedSupplierId())
            End If
        End Sub

        Private Async Sub txtCustomerHistorySearch_KeyDown(sender As Object, e As KeyEventArgs)
            If e.KeyCode = Keys.Enter Then
                e.SuppressKeyPress = True
                Await LoadCustomerHistoryAsync(GetSelectedCustomerId())
            End If
        End Sub

        Private Async Sub txtSupplierHistorySearch_KeyDown(sender As Object, e As KeyEventArgs)
            If e.KeyCode = Keys.Enter Then
                e.SuppressKeyPress = True
                Await LoadSupplierHistoryAsync(GetSelectedSupplierId())
            End If
        End Sub

        Private Sub SetBusy(isBusy As Boolean, Optional message As String = "")
            _isBusy = isBusy
            tabs.Enabled = Not isBusy
            UseWaitCursor = isBusy

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
                        Dim loadTask As Task = LoadCustomersAsync(GetSelectedCustomerId())
                    Else
                        Dim loadTask As Task = LoadSuppliersAsync(GetSelectedSupplierId())
                    End If
                    Return True
                Case Keys.Control Or Keys.S
                    If tabs.SelectedIndex = 0 Then
                        btnCustomerSave.PerformClick()
                    Else
                        btnSupplierSave.PerformClick()
                    End If
                    Return True
                Case Keys.Control Or Keys.N
                    If tabs.SelectedIndex = 0 Then
                        btnCustomerNew.PerformClick()
                    Else
                        btnSupplierNew.PerformClick()
                    End If
                    Return True
                Case Keys.Control Or Keys.D
                    If tabs.SelectedIndex = 0 Then
                        btnCustomerFillDue.PerformClick()
                    Else
                        btnSupplierFillDue.PerformClick()
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
