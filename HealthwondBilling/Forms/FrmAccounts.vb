Imports HealthwondBilling.Models
Imports HealthwondBilling.Services
Imports HealthwondBilling.Utilities
Imports System.ComponentModel
Imports System.Linq

Namespace Forms

    Public Class FrmAccounts
        Inherits Form

        Private ReadOnly _accountingService As AccountingService
        Private ReadOnly dtpFromDate As New DateTimePicker()
        Private ReadOnly dtpToDate As New DateTimePicker()
        Private ReadOnly btnRefresh As New Button()
        Private ReadOnly btnClose As New Button()
        Private ReadOnly tabs As New TabControl()
        Private ReadOnly lblStatus As New Label()

        Private ReadOnly overviewLabels As New Dictionary(Of String, Label)()

        Private ReadOnly txtLedgerSearch As New TextBox()
        Private ReadOnly btnSearchLedgers As New Button()
        Private ReadOnly btnNewLedger As New Button()
        Private ReadOnly btnSaveLedger As New Button()
        Private ReadOnly dgvLedgers As New DataGridView()
        Private ReadOnly txtLedgerName As New TextBox()
        Private ReadOnly cboLedgerGroup As New ComboBox()
        Private ReadOnly nudOpeningBalance As New NumericUpDown()
        Private ReadOnly cboOpeningType As New ComboBox()
        Private ReadOnly txtLedgerNotes As New TextBox()
        Private ReadOnly lblLedgerMeta As New Label()

        Private ReadOnly txtVoucherNumber As New TextBox()
        Private ReadOnly cboVoucherType As New ComboBox()
        Private ReadOnly dtpVoucherDate As New DateTimePicker()
        Private ReadOnly txtVoucherReference As New TextBox()
        Private ReadOnly txtVoucherNarration As New TextBox()
        Private ReadOnly cboVoucherLineLedger As New ComboBox()
        Private ReadOnly cboVoucherLineEntryType As New ComboBox()
        Private ReadOnly nudVoucherLineAmount As New NumericUpDown()
        Private ReadOnly txtVoucherLineRemarks As New TextBox()
        Private ReadOnly btnAddVoucherLine As New Button()
        Private ReadOnly btnRemoveVoucherLine As New Button()
        Private ReadOnly btnNewVoucher As New Button()
        Private ReadOnly btnSaveVoucher As New Button()
        Private ReadOnly lblVoucherTotals As New Label()
        Private ReadOnly cboVoucherFilterType As New ComboBox()
        Private ReadOnly txtVoucherSearch As New TextBox()
        Private ReadOnly dgvVoucherLines As New DataGridView()
        Private ReadOnly dgvVoucherHistory As New DataGridView()

        Private ReadOnly cboStatementLedger As New ComboBox()
        Private ReadOnly btnRefreshStatement As New Button()
        Private ReadOnly dgvLedgerStatement As New DataGridView()
        Private ReadOnly lblStatementSummary As New Label()

        Private ReadOnly _voucherLines As New BindingList(Of VoucherLineItem)()
        Private _accountGroups As New List(Of AccountGroupRecord)()
        Private _allLedgers As New List(Of LedgerRecord)()
        Private _selectedLedgerId As Integer
        Private _isBusy As Boolean

        Public Sub New(accountingService As AccountingService)
            _accountingService = accountingService

            Text = "Healthwond Billing System - Accounts"
            StartPosition = FormStartPosition.CenterParent
            WindowState = FormWindowState.Maximized
            MinimumSize = New Size(1420, 900)
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
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 104))
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 36))

            root.Controls.Add(BuildHeaderPanel(), 0, 0)
            root.Controls.Add(BuildFilterPanel(), 0, 1)
            root.Controls.Add(BuildTabs(), 0, 2)

            lblStatus.Dock = DockStyle.Fill
            lblStatus.TextAlign = ContentAlignment.MiddleLeft
            lblStatus.Font = New Font("Segoe UI Semibold", 9.5F, FontStyle.Bold)
            root.Controls.Add(lblStatus, 0, 3)

            Controls.Add(root)
        End Sub

        Private Function BuildHeaderPanel() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = ThemePalette.AppBackground}

            Dim title As New Label With {
                .Dock = DockStyle.Top,
                .Height = 40,
                .Text = "Accounts",
                .Font = New Font("Segoe UI Semibold", 24.0F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary
            }

            Dim subtitle As New Label With {
                .Dock = DockStyle.Fill,
                .Text = "Review ledger balances, create manual vouchers, inspect day-book activity, and reconcile automated accounting posted from billing, purchases, returns, and settlements.",
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
                .ColumnCount = 5
            }
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 180))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 180))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 110))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 90))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))

            layout.Controls.Add(CreateInputHost("From Date", dtpFromDate), 0, 0)
            layout.Controls.Add(CreateInputHost("To Date", dtpToDate), 1, 0)

            UiStyler.StylePrimaryButton(btnRefresh)
            btnRefresh.Text = "Refresh"
            btnRefresh.Width = 100
            btnRefresh.Margin = New Padding(0, 24, 12, 0)

            UiStyler.StyleSecondaryButton(btnClose)
            btnClose.Text = "Close"
            btnClose.Width = 90
            btnClose.Margin = New Padding(0, 24, 0, 0)

            layout.Controls.Add(btnRefresh, 2, 0)
            layout.Controls.Add(btnClose, 3, 0)
            panel.Controls.Add(layout)
            Return panel
        End Function

        Private Function BuildTabs() As Control
            tabs.Dock = DockStyle.Fill
            tabs.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
            tabs.TabPages.Add(CreateOverviewTab())
            tabs.TabPages.Add(CreateLedgersTab())
            tabs.TabPages.Add(CreateVouchersTab())
            tabs.TabPages.Add(CreateStatementTab())
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
            table.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 58.0F))
            table.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 42.0F))

            For Each metric As String In New String() {
                "Voucher Count",
                "Manual Voucher Count",
                "Auto Voucher Count",
                "Total Debit",
                "Total Credit",
                "Cash Balance",
                "Bank Balance",
                "Receivable Balance",
                "Payable Balance"
            }
                table.RowStyles.Add(New RowStyle(SizeType.Absolute, 40))

                Dim caption As New Label With {
                    .Dock = DockStyle.Fill,
                    .Text = metric,
                    .ForeColor = ThemePalette.TextPrimary,
                    .Font = New Font("Segoe UI", 10.5F, FontStyle.Regular),
                    .TextAlign = ContentAlignment.MiddleLeft
                }

                Dim valueLabel As New Label With {
                    .Dock = DockStyle.Fill,
                    .Text = "0.00",
                    .ForeColor = ThemePalette.TextPrimary,
                    .Font = New Font("Segoe UI Semibold", 11.0F, FontStyle.Bold),
                    .TextAlign = ContentAlignment.MiddleRight
                }

                overviewLabels(metric) = valueLabel
                table.Controls.Add(caption)
                table.Controls.Add(valueLabel)
            Next

            host.Controls.Add(table)
            page.Controls.Add(host)
            Return page
        End Function

        Private Function CreateLedgersTab() As TabPage
            Dim page As New TabPage("Ledgers") With {.BackColor = ThemePalette.AppBackground}

            Dim split As New SplitContainer With {
                .Dock = DockStyle.Fill,
                .SplitterDistance = 720
            }

            Dim leftHost As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(leftHost)

            Dim leftLayout As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .RowCount = 2,
                .ColumnCount = 1
            }
            leftLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 76))
            leftLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            Dim searchPanel As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 3
            }
            searchPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            searchPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 120))
            searchPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 120))

            searchPanel.Controls.Add(CreateInputHost("Search Ledgers", txtLedgerSearch), 0, 0)

            UiStyler.StyleSecondaryButton(btnSearchLedgers)
            btnSearchLedgers.Text = "Search"
            btnSearchLedgers.Dock = DockStyle.Fill
            btnSearchLedgers.Margin = New Padding(0, 24, 10, 0)

            UiStyler.StyleSecondaryButton(btnNewLedger)
            btnNewLedger.Text = "New"
            btnNewLedger.Dock = DockStyle.Fill
            btnNewLedger.Margin = New Padding(0, 24, 0, 0)

            searchPanel.Controls.Add(btnSearchLedgers, 1, 0)
            searchPanel.Controls.Add(btnNewLedger, 2, 0)

            dgvLedgers.Dock = DockStyle.Fill
            leftLayout.Controls.Add(searchPanel, 0, 0)
            leftLayout.Controls.Add(dgvLedgers, 0, 1)
            leftHost.Controls.Add(leftLayout)

            Dim rightHost As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(rightHost)

            Dim editor As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2,
                .Padding = New Padding(0, 14, 0, 0)
            }
            editor.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0F))
            editor.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0F))
            editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 72))
            editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 72))
            editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 110))
            editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 84))
            editor.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            editor.Controls.Add(CreateInputHost("Ledger Name", txtLedgerName), 0, 0)
            editor.Controls.Add(CreateInputHost("Account Group", cboLedgerGroup), 1, 0)
            editor.Controls.Add(CreateInputHost("Opening Balance", nudOpeningBalance), 0, 1)
            editor.Controls.Add(CreateInputHost("Opening Type", cboOpeningType), 1, 1)

            txtLedgerNotes.Multiline = True
            txtLedgerNotes.ScrollBars = ScrollBars.Vertical
            Dim notesHost As Control = CreateInputHost("Notes", txtLedgerNotes)
            editor.Controls.Add(notesHost, 0, 2)
            editor.SetColumnSpan(notesHost, 2)

            lblLedgerMeta.Dock = DockStyle.Fill
            lblLedgerMeta.Font = New Font("Segoe UI", 9.5F, FontStyle.Italic)
            lblLedgerMeta.ForeColor = ThemePalette.TextMuted
            lblLedgerMeta.TextAlign = ContentAlignment.MiddleLeft
            editor.Controls.Add(lblLedgerMeta, 0, 3)
            editor.SetColumnSpan(lblLedgerMeta, 2)

            UiStyler.StylePrimaryButton(btnSaveLedger)
            btnSaveLedger.Text = "Save Ledger"
            btnSaveLedger.Width = 140
            btnSaveLedger.Margin = New Padding(0, 12, 0, 0)

            Dim actionHost As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.LeftToRight,
                .WrapContents = False
            }
            actionHost.Controls.Add(btnSaveLedger)
            editor.Controls.Add(actionHost, 0, 4)
            editor.SetColumnSpan(actionHost, 2)

            rightHost.Controls.Add(UiStyler.CreateScrollableHost(editor))

            split.Panel1.Controls.Add(leftHost)
            split.Panel2.Controls.Add(rightHost)

            page.Controls.Add(split)
            Return page
        End Function

        Private Function CreateVouchersTab() As TabPage
            Dim page As New TabPage("Vouchers") With {.BackColor = ThemePalette.AppBackground}
            Dim split As New SplitContainer With {
                .Dock = DockStyle.Fill,
                .Orientation = Orientation.Horizontal,
                .SplitterDistance = 430
            }

            Dim topHost As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(topHost)

            Dim shell As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1
            }
            shell.RowStyles.Add(New RowStyle(SizeType.Absolute, 174))
            shell.RowStyles.Add(New RowStyle(SizeType.Absolute, 84))
            shell.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            shell.RowStyles.Add(New RowStyle(SizeType.Absolute, 64))

            Dim headerGrid As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 3,
                .Padding = New Padding(0, 14, 0, 0)
            }
            headerGrid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 33.3333F))
            headerGrid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 33.3333F))
            headerGrid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 33.3333F))
            headerGrid.RowStyles.Add(New RowStyle(SizeType.Absolute, 72))
            headerGrid.RowStyles.Add(New RowStyle(SizeType.Absolute, 72))

            headerGrid.Controls.Add(CreateInputHost("Voucher Number", txtVoucherNumber), 0, 0)
            headerGrid.Controls.Add(CreateInputHost("Voucher Type", cboVoucherType), 1, 0)
            headerGrid.Controls.Add(CreateInputHost("Voucher Date", dtpVoucherDate), 2, 0)
            headerGrid.Controls.Add(CreateInputHost("Reference Number", txtVoucherReference), 0, 1)

            txtVoucherNarration.Multiline = True
            Dim narrationHost As Control = CreateInputHost("Narration", txtVoucherNarration)
            headerGrid.Controls.Add(narrationHost, 1, 1)
            headerGrid.SetColumnSpan(narrationHost, 2)

            Dim lineEditor As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 6
            }
            lineEditor.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 42.0F))
            lineEditor.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 100))
            lineEditor.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 140))
            lineEditor.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 28.0F))
            lineEditor.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 110))
            lineEditor.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 130))

            lineEditor.Controls.Add(CreateInputHost("Ledger", cboVoucherLineLedger), 0, 0)
            lineEditor.Controls.Add(CreateInputHost("Dr / Cr", cboVoucherLineEntryType), 1, 0)
            lineEditor.Controls.Add(CreateInputHost("Amount", nudVoucherLineAmount), 2, 0)
            lineEditor.Controls.Add(CreateInputHost("Remarks", txtVoucherLineRemarks), 3, 0)

            UiStyler.StylePrimaryButton(btnAddVoucherLine)
            btnAddVoucherLine.Text = "Add Line"
            btnAddVoucherLine.Dock = DockStyle.Fill
            btnAddVoucherLine.Margin = New Padding(0, 24, 10, 0)

            UiStyler.StyleDangerButton(btnRemoveVoucherLine)
            btnRemoveVoucherLine.Text = "Remove"
            btnRemoveVoucherLine.Dock = DockStyle.Fill
            btnRemoveVoucherLine.Margin = New Padding(0, 24, 0, 0)

            lineEditor.Controls.Add(btnAddVoucherLine, 4, 0)
            lineEditor.Controls.Add(btnRemoveVoucherLine, 5, 0)

            dgvVoucherLines.Dock = DockStyle.Fill

            Dim actionPanel As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 3
            }
            actionPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            actionPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 130))
            actionPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 150))

            lblVoucherTotals.Dock = DockStyle.Fill
            lblVoucherTotals.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
            lblVoucherTotals.ForeColor = ThemePalette.TextPrimary
            lblVoucherTotals.TextAlign = ContentAlignment.MiddleLeft

            UiStyler.StyleSecondaryButton(btnNewVoucher)
            btnNewVoucher.Text = "New Voucher"
            btnNewVoucher.Dock = DockStyle.Fill
            btnNewVoucher.Margin = New Padding(0, 0, 10, 0)

            UiStyler.StylePrimaryButton(btnSaveVoucher)
            btnSaveVoucher.Text = "Save Voucher"
            btnSaveVoucher.Dock = DockStyle.Fill

            actionPanel.Controls.Add(lblVoucherTotals, 0, 0)
            actionPanel.Controls.Add(btnNewVoucher, 1, 0)
            actionPanel.Controls.Add(btnSaveVoucher, 2, 0)

            shell.Controls.Add(headerGrid, 0, 0)
            shell.Controls.Add(lineEditor, 0, 1)
            shell.Controls.Add(dgvVoucherLines, 0, 2)
            shell.Controls.Add(actionPanel, 0, 3)
            topHost.Controls.Add(shell)

            Dim bottomHost As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(bottomHost)

            Dim historyShell As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1
            }
            historyShell.RowStyles.Add(New RowStyle(SizeType.Absolute, 78))
            historyShell.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            Dim historyFilter As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 3
            }
            historyFilter.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 220))
            historyFilter.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            historyFilter.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 120))

            historyFilter.Controls.Add(CreateInputHost("Voucher Type Filter", cboVoucherFilterType), 0, 0)
            historyFilter.Controls.Add(CreateInputHost("Search", txtVoucherSearch), 1, 0)

            Dim btnRefreshHistory As New Button()
            UiStyler.StyleSecondaryButton(btnRefreshHistory)
            btnRefreshHistory.Text = "Reload"
            btnRefreshHistory.Dock = DockStyle.Fill
            btnRefreshHistory.Margin = New Padding(0, 24, 0, 0)
            AddHandler btnRefreshHistory.Click, AddressOf btnRefresh_Click
            historyFilter.Controls.Add(btnRefreshHistory, 2, 0)

            dgvVoucherHistory.Dock = DockStyle.Fill
            historyShell.Controls.Add(historyFilter, 0, 0)
            historyShell.Controls.Add(dgvVoucherHistory, 0, 1)
            bottomHost.Controls.Add(historyShell)

            split.Panel1.Controls.Add(topHost)
            split.Panel2.Controls.Add(bottomHost)
            page.Controls.Add(split)
            Return page
        End Function

        Private Function CreateStatementTab() As TabPage
            Dim page As New TabPage("Ledger Statement") With {.BackColor = ThemePalette.AppBackground}
            Dim host As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(host)

            Dim shell As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1
            }
            shell.RowStyles.Add(New RowStyle(SizeType.Absolute, 76))
            shell.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            shell.RowStyles.Add(New RowStyle(SizeType.Absolute, 42))

            Dim filter As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2
            }
            filter.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            filter.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 140))

            filter.Controls.Add(CreateInputHost("Ledger", cboStatementLedger), 0, 0)

            UiStyler.StyleSecondaryButton(btnRefreshStatement)
            btnRefreshStatement.Text = "Load Statement"
            btnRefreshStatement.Dock = DockStyle.Fill
            btnRefreshStatement.Margin = New Padding(0, 24, 0, 0)

            filter.Controls.Add(btnRefreshStatement, 1, 0)

            dgvLedgerStatement.Dock = DockStyle.Fill
            lblStatementSummary.Dock = DockStyle.Fill
            lblStatementSummary.Font = New Font("Segoe UI Semibold", 9.75F, FontStyle.Bold)
            lblStatementSummary.ForeColor = ThemePalette.TextPrimary
            lblStatementSummary.TextAlign = ContentAlignment.MiddleLeft

            shell.Controls.Add(filter, 0, 0)
            shell.Controls.Add(dgvLedgerStatement, 0, 1)
            shell.Controls.Add(lblStatementSummary, 0, 2)
            host.Controls.Add(shell)
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

            dtpVoucherDate.Format = DateTimePickerFormat.Custom
            dtpVoucherDate.CustomFormat = "dd-MMM-yyyy"
            dtpVoucherDate.Value = DateTime.Today

            For Each textbox As TextBox In New TextBox() {txtLedgerSearch, txtLedgerName, txtLedgerNotes, txtVoucherNumber, txtVoucherReference, txtVoucherNarration, txtVoucherLineRemarks, txtVoucherSearch}
                textbox.BorderStyle = BorderStyle.FixedSingle
                UiStyler.StyleInput(textbox)
            Next

            txtLedgerNotes.Multiline = True
            txtLedgerNotes.ScrollBars = ScrollBars.Vertical
            txtVoucherNarration.Multiline = True
            txtVoucherNarration.ScrollBars = ScrollBars.Vertical

            ConfigureCombo(cboLedgerGroup)
            ConfigureCombo(cboOpeningType)
            ConfigureCombo(cboVoucherType)
            ConfigureCombo(cboVoucherLineLedger)
            ConfigureCombo(cboVoucherLineEntryType)
            ConfigureCombo(cboVoucherFilterType)
            ConfigureCombo(cboStatementLedger)

            cboOpeningType.Items.AddRange(New Object() {"Dr", "Cr"})
            cboOpeningType.SelectedIndex = 0

            cboVoucherLineEntryType.Items.AddRange(New Object() {"Dr", "Cr"})
            cboVoucherLineEntryType.SelectedIndex = 0

            nudOpeningBalance.Maximum = 100000000D
            nudOpeningBalance.DecimalPlaces = 2
            nudOpeningBalance.ThousandsSeparator = True
            nudVoucherLineAmount.Maximum = 100000000D
            nudVoucherLineAmount.DecimalPlaces = 2
            nudVoucherLineAmount.ThousandsSeparator = True

            ConfigureGrid(dgvLedgers)
            ConfigureGrid(dgvVoucherHistory)
            ConfigureGrid(dgvLedgerStatement)
            ConfigureGrid(dgvVoucherLines)
            dgvVoucherLines.AutoGenerateColumns = True
            dgvVoucherLines.DataSource = _voucherLines

            cboVoucherFilterType.Items.Add("All")
            For Each value As String In _accountingService.GetVoucherTypes()
                cboVoucherType.Items.Add(value)
                cboVoucherFilterType.Items.Add(value)
            Next
            cboVoucherFilterType.SelectedIndex = 0
            cboVoucherType.SelectedIndex = 0
        End Sub

        Private Sub ConfigureCombo(combo As ComboBox)
            combo.DropDownStyle = ComboBoxStyle.DropDownList
            combo.Font = New Font("Segoe UI", 10.0F, FontStyle.Regular)
        End Sub

        Private Sub ConfigureGrid(grid As DataGridView)
            UiStyler.StyleDataGrid(grid)
            grid.AutoGenerateColumns = True
            grid.ReadOnly = True
        End Sub

        Private Sub WireEvents()
            AddHandler Load, AddressOf FrmAccounts_Load
            AddHandler btnRefresh.Click, AddressOf btnRefresh_Click
            AddHandler btnClose.Click, AddressOf btnClose_Click
            AddHandler btnSearchLedgers.Click, AddressOf btnSearchLedgers_Click
            AddHandler btnNewLedger.Click, AddressOf btnNewLedger_Click
            AddHandler btnSaveLedger.Click, AddressOf btnSaveLedger_Click
            AddHandler dgvLedgers.SelectionChanged, AddressOf dgvLedgers_SelectionChanged
            AddHandler btnAddVoucherLine.Click, AddressOf btnAddVoucherLine_Click
            AddHandler btnRemoveVoucherLine.Click, AddressOf btnRemoveVoucherLine_Click
            AddHandler btnNewVoucher.Click, AddressOf btnNewVoucher_Click
            AddHandler btnSaveVoucher.Click, AddressOf btnSaveVoucher_Click
            AddHandler cboVoucherType.SelectedIndexChanged, AddressOf VoucherIdentityChanged
            AddHandler dtpVoucherDate.ValueChanged, AddressOf VoucherIdentityChanged
            AddHandler btnRefreshStatement.Click, AddressOf btnRefreshStatement_Click
        End Sub

        Private Async Sub FrmAccounts_Load(sender As Object, e As EventArgs)
            Await LoadAllAsync(True)
        End Sub

        Private Async Sub btnRefresh_Click(sender As Object, e As EventArgs)
            Await LoadAllAsync(False)
        End Sub

        Private Async Function LoadAllAsync(includeMasters As Boolean) As Task
            If dtpFromDate.Value.Date > dtpToDate.Value.Date Then
                ShowStatus("From date cannot be later than To date.", True)
                Return
            End If

            SetBusy(True, "Loading accounts...")

            Try
                If includeMasters Then
                    Dim groupsTask = _accountingService.LoadAccountGroupsAsync()
                    Dim ledgersTask = _accountingService.LoadLedgersAsync(String.Empty)
                    Await Task.WhenAll(groupsTask, ledgersTask)
                    _accountGroups = groupsTask.Result
                    _allLedgers = ledgersTask.Result
                    BindGroupCombos()
                    BindLedgersGrid(_allLedgers)
                    ResetLedgerEditor()
                    ResetVoucherEditor()
                Else
                    _allLedgers = Await _accountingService.LoadLedgersAsync(txtLedgerSearch.Text)
                    BindLedgersGrid(_allLedgers)
                    BindLedgerCombos()
                End If

                Dim overviewTask = _accountingService.LoadOverviewAsync(dtpFromDate.Value.Date, dtpToDate.Value.Date)
                Dim vouchersTask = _accountingService.LoadVouchersAsync(dtpFromDate.Value.Date, dtpToDate.Value.Date, GetSelectedVoucherTypeFilter(), txtVoucherSearch.Text)
                Await Task.WhenAll(overviewTask, vouchersTask)

                BindOverview(overviewTask.Result)
                BindVoucherHistory(vouchersTask.Result)
                Await LoadStatementAsync()
                ShowStatus("Accounts loaded successfully.", False)
            Catch ex As Exception
                AppLogger.Error("Accounts workspace load failed.", ex)
                ShowStatus("Accounts workspace could not be loaded.", True)
            Finally
                SetBusy(False)
            End Try
        End Function

        Private Sub BindOverview(overview As AccountingOverview)
            SetOverview("Voucher Count", overview.VoucherCount.ToString("N0"))
            SetOverview("Manual Voucher Count", overview.ManualVoucherCount.ToString("N0"))
            SetOverview("Auto Voucher Count", overview.AutoVoucherCount.ToString("N0"))
            SetOverview("Total Debit", overview.TotalDebit.ToString("N2"))
            SetOverview("Total Credit", overview.TotalCredit.ToString("N2"))
            SetOverview("Cash Balance", $"{overview.CashBalance:N2} {overview.CashBalanceType}")
            SetOverview("Bank Balance", $"{overview.BankBalance:N2} {overview.BankBalanceType}")
            SetOverview("Receivable Balance", $"{overview.ReceivableBalance:N2} {overview.ReceivableBalanceType}")
            SetOverview("Payable Balance", $"{overview.PayableBalance:N2} {overview.PayableBalanceType}")
        End Sub

        Private Sub SetOverview(metricName As String, value As String)
            If overviewLabels.ContainsKey(metricName) Then
                overviewLabels(metricName).Text = value
            End If
        End Sub

        Private Sub BindGroupCombos()
            cboLedgerGroup.DataSource = Nothing
            cboLedgerGroup.DataSource = _accountGroups.ToList()
            cboLedgerGroup.DisplayMember = NameOf(AccountGroupRecord.GroupName)
            cboLedgerGroup.ValueMember = NameOf(AccountGroupRecord.Id)
            BindLedgerCombos()
        End Sub

        Private Sub BindLedgerCombos()
            Dim selectableLedgers As List(Of LedgerRecord) = _allLedgers.OrderBy(Function(item) item.LedgerName).ToList()

            cboVoucherLineLedger.DataSource = Nothing
            cboVoucherLineLedger.DataSource = selectableLedgers.ToList()
            cboVoucherLineLedger.DisplayMember = NameOf(LedgerRecord.LedgerName)
            cboVoucherLineLedger.ValueMember = NameOf(LedgerRecord.Id)

            cboStatementLedger.DataSource = Nothing
            cboStatementLedger.DataSource = selectableLedgers.ToList()
            cboStatementLedger.DisplayMember = NameOf(LedgerRecord.LedgerName)
            cboStatementLedger.ValueMember = NameOf(LedgerRecord.Id)

            If _selectedLedgerId > 0 Then
                cboStatementLedger.SelectedValue = _selectedLedgerId
            End If
        End Sub

        Private Sub BindLedgersGrid(rows As List(Of LedgerRecord))
            dgvLedgers.DataSource = Nothing
            dgvLedgers.DataSource = rows
            FormatGrid(dgvLedgers)
        End Sub

        Private Sub BindVoucherHistory(rows As List(Of VoucherHistoryRow))
            dgvVoucherHistory.DataSource = Nothing
            dgvVoucherHistory.DataSource = rows
            FormatGrid(dgvVoucherHistory)
        End Sub

        Private Sub FormatGrid(grid As DataGridView)
            For Each column As DataGridViewColumn In grid.Columns
                Dim propertyName As String = column.DataPropertyName
                column.HeaderText = InsertSpaces(propertyName)

                If propertyName.EndsWith("Date", StringComparison.OrdinalIgnoreCase) Then
                    column.DefaultCellStyle.Format = "dd-MMM-yyyy"
                ElseIf propertyName.IndexOf("Amount", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                       propertyName.IndexOf("Balance", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    column.DefaultCellStyle.Format = "N2"
                End If
            Next
        End Sub

        Private Async Sub btnSearchLedgers_Click(sender As Object, e As EventArgs)
            SetBusy(True, "Searching ledgers...")
            Try
                _allLedgers = Await _accountingService.LoadLedgersAsync(txtLedgerSearch.Text)
                BindLedgersGrid(_allLedgers)
                BindLedgerCombos()
                ShowStatus("Ledgers refreshed.", False)
            Catch ex As Exception
                AppLogger.Error("Ledger search failed.", ex)
                ShowStatus("Ledgers could not be loaded.", True)
            Finally
                SetBusy(False)
            End Try
        End Sub

        Private Sub btnNewLedger_Click(sender As Object, e As EventArgs)
            ResetLedgerEditor()
        End Sub

        Private Async Sub btnSaveLedger_Click(sender As Object, e As EventArgs)
            Dim record As New LedgerRecord With {
                .Id = _selectedLedgerId,
                .LedgerName = txtLedgerName.Text,
                .AccountGroupId = GetSelectedComboValue(cboLedgerGroup),
                .OpeningBalance = nudOpeningBalance.Value,
                .OpeningBalanceType = Convert.ToString(cboOpeningType.SelectedItem),
                .Notes = txtLedgerNotes.Text
            }

            SetBusy(True, "Saving ledger...")
            Dim result As EntityOperationResult = Await _accountingService.SaveLedgerAsync(record)
            SetBusy(False)
            ShowStatus(result.Message, Not result.IsSuccess)

            If result.IsSuccess Then
                _selectedLedgerId = result.EntityId
                Await btnSearchLedgers_ClickAsync()
            End If
        End Sub

        Private Async Function btnSearchLedgers_ClickAsync() As Task
            _allLedgers = Await _accountingService.LoadLedgersAsync(txtLedgerSearch.Text)
            BindLedgersGrid(_allLedgers)
            BindLedgerCombos()
            Dim selectedRecord As LedgerRecord = _allLedgers.FirstOrDefault(Function(item) item.Id = _selectedLedgerId)
            If selectedRecord IsNot Nothing Then
                BindLedgerEditor(selectedRecord)
            Else
                ResetLedgerEditor()
            End If
        End Function

        Private Sub dgvLedgers_SelectionChanged(sender As Object, e As EventArgs)
            If dgvLedgers.CurrentRow Is Nothing Then
                Return
            End If

            Dim record As LedgerRecord = TryCast(dgvLedgers.CurrentRow.DataBoundItem, LedgerRecord)
            If record Is Nothing Then
                Return
            End If

            BindLedgerEditor(record)
        End Sub

        Private Sub BindLedgerEditor(record As LedgerRecord)
            _selectedLedgerId = record.Id
            txtLedgerName.Text = record.LedgerName
            cboLedgerGroup.SelectedValue = record.AccountGroupId
            nudOpeningBalance.Value = Math.Min(nudOpeningBalance.Maximum, record.OpeningBalance)
            cboOpeningType.SelectedItem = record.OpeningBalanceType
            txtLedgerNotes.Text = record.Notes
            lblLedgerMeta.Text = $"Current Balance: {record.CurrentBalance:N2} {record.CurrentBalanceType} | System: {If(record.IsSystem, "Yes", "No")} | Party Ledger: {If(record.IsPartyLedger, "Yes", "No")}"
            cboStatementLedger.SelectedValue = record.Id
        End Sub

        Private Sub ResetLedgerEditor()
            _selectedLedgerId = 0
            txtLedgerName.Clear()
            If cboLedgerGroup.Items.Count > 0 Then
                cboLedgerGroup.SelectedIndex = 0
            End If
            nudOpeningBalance.Value = 0D
            cboOpeningType.SelectedIndex = 0
            txtLedgerNotes.Clear()
            lblLedgerMeta.Text = "Create a new manual ledger or select an existing non-system ledger to edit."
        End Sub

        Private Sub btnAddVoucherLine_Click(sender As Object, e As EventArgs)
            Dim ledger As LedgerRecord = TryCast(cboVoucherLineLedger.SelectedItem, LedgerRecord)
            If ledger Is Nothing Then
                ShowStatus("Select a ledger for the voucher line.", True)
                Return
            End If

            Dim amount As Decimal = Decimal.Round(nudVoucherLineAmount.Value, 2, MidpointRounding.AwayFromZero)
            If amount <= 0D Then
                ShowStatus("Voucher line amount must be greater than zero.", True)
                Return
            End If

            _voucherLines.Add(New VoucherLineItem With {
                .LineNumber = _voucherLines.Count + 1,
                .LedgerId = ledger.Id,
                .LedgerName = ledger.LedgerName,
                .EntryType = Convert.ToString(cboVoucherLineEntryType.SelectedItem),
                .Amount = amount,
                .Remarks = txtVoucherLineRemarks.Text
            })
            RenumberVoucherLines()
            UpdateVoucherTotals()
            nudVoucherLineAmount.Value = 0D
            txtVoucherLineRemarks.Clear()
        End Sub

        Private Sub btnRemoveVoucherLine_Click(sender As Object, e As EventArgs)
            If dgvVoucherLines.CurrentRow Is Nothing Then
                Return
            End If

            Dim line As VoucherLineItem = TryCast(dgvVoucherLines.CurrentRow.DataBoundItem, VoucherLineItem)
            If line Is Nothing Then
                Return
            End If

            _voucherLines.Remove(line)
            RenumberVoucherLines()
            UpdateVoucherTotals()
        End Sub

        Private Async Sub btnNewVoucher_Click(sender As Object, e As EventArgs)
            ResetVoucherEditor()
            Await GenerateVoucherNumberAsync()
        End Sub

        Private Async Sub VoucherIdentityChanged(sender As Object, e As EventArgs)
            If _isBusy Then
                Return
            End If

            If txtVoucherNumber.Text.Trim() = String.Empty OrElse txtVoucherNumber.Tag Is Nothing OrElse Convert.ToBoolean(txtVoucherNumber.Tag) Then
                Await GenerateVoucherNumberAsync()
            End If
        End Sub

        Private Async Function GenerateVoucherNumberAsync() As Task
            If cboVoucherType.SelectedItem Is Nothing Then
                Return
            End If

            txtVoucherNumber.Tag = True
            txtVoucherNumber.Text = Await _accountingService.GenerateNextVoucherNumberAsync(Convert.ToString(cboVoucherType.SelectedItem), dtpVoucherDate.Value.Date)
            txtVoucherNumber.Tag = False
        End Function

        Private Async Sub btnSaveVoucher_Click(sender As Object, e As EventArgs)
            Dim draft As New AccountingVoucherDraft With {
                .VoucherNumber = txtVoucherNumber.Text,
                .VoucherType = Convert.ToString(cboVoucherType.SelectedItem),
                .VoucherDate = dtpVoucherDate.Value.Date,
                .ReferenceNumber = txtVoucherReference.Text,
                .Narration = txtVoucherNarration.Text,
                .Lines = _voucherLines.ToList()
            }

            SetBusy(True, "Saving voucher...")
            Dim result As EntityOperationResult = Await _accountingService.SaveManualVoucherAsync(draft, SessionManager.CurrentUser.Id)
            SetBusy(False)
            ShowStatus(result.Message, Not result.IsSuccess)

            If result.IsSuccess Then
                ResetVoucherEditor()
                Await LoadAllAsync(False)
                Await GenerateVoucherNumberAsync()
            End If
        End Sub

        Private Sub ResetVoucherEditor()
            txtVoucherReference.Clear()
            txtVoucherNarration.Clear()
            _voucherLines.Clear()
            RenumberVoucherLines()
            UpdateVoucherTotals()
            nudVoucherLineAmount.Value = 0D
            txtVoucherLineRemarks.Clear()
            dtpVoucherDate.Value = DateTime.Today
            If cboVoucherType.Items.Count > 0 Then
                cboVoucherType.SelectedIndex = 0
            End If
        End Sub

        Private Sub RenumberVoucherLines()
            Dim lineNumber As Integer = 1
            For Each line As VoucherLineItem In _voucherLines
                line.LineNumber = lineNumber
                line.PropertyChangedNotification()
                lineNumber += 1
            Next

            dgvVoucherLines.Refresh()
        End Sub

        Private Sub UpdateVoucherTotals()
            Dim totalDebit As Decimal
            Dim totalCredit As Decimal
            _accountingService.RecalculateVoucherTotals(_voucherLines, totalDebit, totalCredit)
            lblVoucherTotals.Text = $"Debit: {totalDebit:N2}    Credit: {totalCredit:N2}"
        End Sub

        Private Function GetSelectedVoucherTypeFilter() As String
            Dim selectedValue As String = Convert.ToString(cboVoucherFilterType.SelectedItem)
            If String.Equals(selectedValue, "All", StringComparison.OrdinalIgnoreCase) Then
                Return String.Empty
            End If

            Return selectedValue
        End Function

        Private Async Sub btnRefreshStatement_Click(sender As Object, e As EventArgs)
            SetBusy(True, "Loading ledger statement...")
            Try
                Await LoadStatementAsync()
                ShowStatus("Ledger statement loaded successfully.", False)
            Catch ex As Exception
                AppLogger.Error("Ledger statement loading failed.", ex)
                ShowStatus("Ledger statement could not be loaded.", True)
            Finally
                SetBusy(False)
            End Try
        End Sub

        Private Async Function LoadStatementAsync() As Task
            Dim selectedLedgerId As Integer = GetSelectedComboValue(cboStatementLedger)
            If selectedLedgerId <= 0 Then
                dgvLedgerStatement.DataSource = Nothing
                lblStatementSummary.Text = "Select a ledger to inspect its statement."
                Return
            End If

            Dim rows As List(Of LedgerStatementRow) =
                Await _accountingService.LoadLedgerStatementAsync(selectedLedgerId, dtpFromDate.Value.Date, dtpToDate.Value.Date)
            dgvLedgerStatement.DataSource = Nothing
            dgvLedgerStatement.DataSource = rows
            FormatGrid(dgvLedgerStatement)

            If rows.Count > 0 Then
                Dim lastRow As LedgerStatementRow = rows(rows.Count - 1)
                lblStatementSummary.Text = $"Closing Balance: {lastRow.RunningBalance:N2} {lastRow.RunningBalanceType}"
            Else
                lblStatementSummary.Text = "No entries found for the selected period."
            End If
        End Function

        Private Function GetSelectedComboValue(combo As ComboBox) As Integer
            If combo.SelectedValue Is Nothing Then
                Return 0
            End If

            Return Convert.ToInt32(combo.SelectedValue, Globalization.CultureInfo.InvariantCulture)
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

        Private Sub SetBusy(isBusy As Boolean, Optional message As String = "")
            _isBusy = isBusy
            dtpFromDate.Enabled = Not isBusy
            dtpToDate.Enabled = Not isBusy
            btnRefresh.Enabled = Not isBusy
            btnClose.Enabled = Not isBusy
            btnSearchLedgers.Enabled = Not isBusy
            btnNewLedger.Enabled = Not isBusy
            btnSaveLedger.Enabled = Not isBusy
            btnAddVoucherLine.Enabled = Not isBusy
            btnRemoveVoucherLine.Enabled = Not isBusy
            btnNewVoucher.Enabled = Not isBusy
            btnSaveVoucher.Enabled = Not isBusy
            btnRefreshStatement.Enabled = Not isBusy

            If isBusy Then
                lblStatus.ForeColor = ThemePalette.TextMuted
                lblStatus.Text = message
            End If
        End Sub

        Private Sub ShowStatus(message As String, isError As Boolean)
            lblStatus.ForeColor = If(isError, ThemePalette.DangerRed, ThemePalette.AccentGreen)
            lblStatus.Text = message
        End Sub

        Private Sub btnClose_Click(sender As Object, e As EventArgs)
            Close()
        End Sub

        Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
            Select Case keyData
                Case Keys.F5
                    If Not _isBusy Then
                        Dim loadTask As Task = LoadAllAsync(False)
                    End If
                    Return True
                Case Keys.Control Or Keys.S
                    If tabs.SelectedTab IsNot Nothing AndAlso tabs.SelectedTab.Text = "Ledgers" Then
                        btnSaveLedger.PerformClick()
                    ElseIf tabs.SelectedTab IsNot Nothing AndAlso tabs.SelectedTab.Text = "Vouchers" Then
                        btnSaveVoucher.PerformClick()
                    End If
                    Return True
                Case Keys.Escape
                    Close()
                    Return True
            End Select

            Return MyBase.ProcessCmdKey(msg, keyData)
        End Function

    End Class

    Friend Module VoucherLineItemExtensions
        <Runtime.CompilerServices.Extension>
        Public Sub PropertyChangedNotification(line As VoucherLineItem)
        End Sub
    End Module

End Namespace
