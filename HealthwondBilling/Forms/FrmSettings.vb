Imports HealthwondBilling.Models
Imports HealthwondBilling.Services
Imports HealthwondBilling.Utilities
Imports System.Diagnostics
Imports System.IO

Namespace Forms

    Public Class FrmSettings
        Inherits Form

        Private ReadOnly _settingsService As SettingsService
        Private ReadOnly _maintenanceService As MaintenanceService

        Private ReadOnly txtCompanyName As New TextBox()
        Private ReadOnly txtCompanyAddress As New TextBox()
        Private ReadOnly txtCompanyPhone As New TextBox()
        Private ReadOnly txtCompanyGstin As New TextBox()
        Private ReadOnly txtCompanyDrugLicense As New TextBox()
        Private ReadOnly txtInvoicePrefix As New TextBox()
        Private ReadOnly txtPurchasePrefix As New TextBox()
        Private ReadOnly txtReceiptPrefix As New TextBox()
        Private ReadOnly txtSupplierPaymentPrefix As New TextBox()
        Private ReadOnly nudLowStockThreshold As New NumericUpDown()
        Private ReadOnly txtCurrencySymbol As New TextBox()
        Private ReadOnly txtInvoiceTemplatePath As New TextBox()

        Private ReadOnly btnBrowseTemplate As New Button()
        Private ReadOnly btnResetTemplate As New Button()
        Private ReadOnly btnReload As New Button()
        Private ReadOnly btnSave As New Button()
        Private ReadOnly btnOpenTemplates As New Button()
        Private ReadOnly btnOpenDataRoot As New Button()
        Private ReadOnly btnClose As New Button()
        Private ReadOnly btnCreateBackup As New Button()
        Private ReadOnly btnRestoreBackup As New Button()
        Private ReadOnly btnOptimizeDatabase As New Button()
        Private ReadOnly btnOpenBackups As New Button()
        Private ReadOnly btnOpenInvoices As New Button()
        Private ReadOnly btnOpenReports As New Button()
        Private ReadOnly btnOpenLogs As New Button()
        Private ReadOnly lblStatus As New Label()
        Private ReadOnly lblResolvedTemplatePath As New Label()
        Private ReadOnly lblDataRootPath As Label = CreatePathValueLabel()
        Private ReadOnly lblDatabasePath As Label = CreatePathValueLabel()
        Private ReadOnly lblBackupsPath As Label = CreatePathValueLabel()
        Private ReadOnly lblInvoicesPath As Label = CreatePathValueLabel()
        Private ReadOnly lblReportsPath As Label = CreatePathValueLabel()
        Private ReadOnly lblLogsPath As Label = CreatePathValueLabel()
        Private ReadOnly lblTemplateFolderPath As Label = CreatePathValueLabel()

        Private _isBusy As Boolean

        Public Sub New(settingsService As SettingsService, maintenanceService As MaintenanceService)
            _settingsService = settingsService
            _maintenanceService = maintenanceService

            Text = "Healthwond Billing System - Settings"
            StartPosition = FormStartPosition.CenterParent
            Size = New Size(1420, 880)
            MinimumSize = New Size(1260, 780)
            BackColor = ThemePalette.AppBackground
            Font = New Font("Segoe UI", 10.0F, FontStyle.Regular)
            KeyPreview = True

            BuildLayout()
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
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 104))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 82))
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 36))

            root.Controls.Add(BuildHeaderPanel(), 0, 0)
            root.Controls.Add(BuildToolbarPanel(), 0, 1)
            root.Controls.Add(BuildMainPanel(), 0, 2)

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
                .Text = "Settings",
                .Font = New Font("Segoe UI Semibold", 24.0F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary
            }

            Dim subtitle As New Label With {
                .Dock = DockStyle.Fill,
                .Text = "Maintain company invoice identity, numbering prefixes, invoice template settings, and operational database maintenance from one admin workspace.",
                .Font = New Font("Segoe UI", 10.5F, FontStyle.Regular),
                .ForeColor = ThemePalette.TextMuted
            }

            panel.Controls.Add(subtitle)
            panel.Controls.Add(title)
            Return panel
        End Function

        Private Function BuildToolbarPanel() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)

            Dim flow As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.LeftToRight,
                .WrapContents = False,
                .BackColor = Color.Transparent,
                .Padding = New Padding(0, 8, 0, 0)
            }

            ConfigureToolbarButton(btnReload, "Reload", AddressOf btnReload_Click, False, 100)
            ConfigureToolbarButton(btnSave, "Save", AddressOf btnSave_Click, True, 100)
            ConfigureToolbarButton(btnOpenTemplates, "Open Templates", AddressOf btnOpenTemplates_Click, False, 130)
            ConfigureToolbarButton(btnOpenDataRoot, "Open Data Folder", AddressOf btnOpenDataRoot_Click, False, 140)
            ConfigureToolbarButton(btnClose, "Close", AddressOf btnClose_Click, False, 90)

            flow.Controls.AddRange(New Control() {btnReload, btnSave, btnOpenTemplates, btnOpenDataRoot, btnClose})
            panel.Controls.Add(flow)
            Return panel
        End Function

        Private Function BuildMainPanel() As Control
            Dim split As New SplitContainer With {
                .Dock = DockStyle.Fill,
                .SplitterDistance = 620,
                .BackColor = ThemePalette.AppBackground
            }

            split.Panel1.Controls.Add(BuildCompanyCard())
            split.Panel2.Controls.Add(BuildRightWorkspace())
            Return split
        End Function

        Private Function BuildCompanyCard() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)

            Dim title As New Label With {
                .Dock = DockStyle.Top,
                .Height = 30,
                .Text = "Company profile",
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
            editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 72))
            editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 72))
            editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 160))
            editor.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            editor.Controls.Add(CreateInputHost("Company Name", txtCompanyName), 0, 0)
            editor.Controls.Add(CreateInputHost("Company Phone", txtCompanyPhone), 1, 0)
            editor.Controls.Add(CreateInputHost("Company GSTIN", txtCompanyGstin), 0, 1)
            editor.Controls.Add(CreateInputHost("Drug License Number", txtCompanyDrugLicense), 1, 1)

            txtCompanyAddress.Multiline = True
            txtCompanyAddress.ScrollBars = ScrollBars.Vertical
            Dim addressHost As Control = CreateInputHost("Company Address", txtCompanyAddress)
            editor.Controls.Add(addressHost, 0, 2)
            editor.SetColumnSpan(addressHost, 2)

            Dim note As New Label With {
                .Dock = DockStyle.Fill,
                .Text = "These values are used across GST invoices, purchase printouts, PDF exports, and Excel template generation.",
                .ForeColor = ThemePalette.TextMuted,
                .Font = New Font("Segoe UI", 9.75F, FontStyle.Italic),
                .TextAlign = ContentAlignment.TopLeft
            }
            editor.Controls.Add(note, 0, 3)
            editor.SetColumnSpan(note, 2)

            panel.Controls.Add(UiStyler.CreateScrollableHost(editor))
            panel.Controls.Add(title)
            Return panel
        End Function

        Private Function BuildRightWorkspace() As Control
            Dim layout As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 2,
                .BackColor = ThemePalette.AppBackground
            }
            layout.RowStyles.Add(New RowStyle(SizeType.Percent, 50.0F))
            layout.RowStyles.Add(New RowStyle(SizeType.Percent, 50.0F))

            Dim systemCard As Control = BuildSystemCard()
            systemCard.Margin = New Padding(0, 0, 0, 10)

            Dim maintenanceCard As Control = BuildMaintenanceCard()
            maintenanceCard.Margin = New Padding(0, 10, 0, 0)

            layout.Controls.Add(systemCard, 0, 0)
            layout.Controls.Add(maintenanceCard, 0, 1)
            Return layout
        End Function

        Private Function BuildSystemCard() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)

            Dim title As New Label With {
                .Dock = DockStyle.Top,
                .Height = 30,
                .Text = "System configuration",
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
            editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 72))
            editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 72))
            editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 72))
            editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 110))
            editor.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            editor.Controls.Add(CreateInputHost("Invoice Prefix", txtInvoicePrefix), 0, 0)
            editor.Controls.Add(CreateInputHost("Purchase Prefix", txtPurchasePrefix), 1, 0)
            editor.Controls.Add(CreateInputHost("Receipt Prefix", txtReceiptPrefix), 0, 1)
            editor.Controls.Add(CreateInputHost("Supplier Payment Prefix", txtSupplierPaymentPrefix), 1, 1)
            editor.Controls.Add(CreateInputHost("Currency Symbol", txtCurrencySymbol), 0, 2)
            editor.Controls.Add(CreateInputHost("Low Stock Threshold", nudLowStockThreshold), 1, 2)

            Dim templateHost As New Panel With {.Dock = DockStyle.Fill, .Padding = New Padding(0, 0, 12, 10)}
            Dim templateLabel As New Label With {
                .Dock = DockStyle.Top,
                .Height = 22,
                .Text = "Invoice Template Path",
                .Font = New Font("Segoe UI Semibold", 9.5F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary
            }

            Dim templateLayout As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 3
            }
            templateLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            templateLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 90))
            templateLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 120))

            templateLayout.Controls.Add(txtInvoiceTemplatePath, 0, 0)

            UiStyler.StyleSecondaryButton(btnBrowseTemplate)
            btnBrowseTemplate.Text = "Browse"
            btnBrowseTemplate.Dock = DockStyle.Fill

            UiStyler.StyleSecondaryButton(btnResetTemplate)
            btnResetTemplate.Text = "Use Default"
            btnResetTemplate.Dock = DockStyle.Fill

            templateLayout.Controls.Add(btnBrowseTemplate, 1, 0)
            templateLayout.Controls.Add(btnResetTemplate, 2, 0)

            templateHost.Controls.Add(templateLayout)
            templateHost.Controls.Add(templateLabel)
            editor.Controls.Add(templateHost, 0, 3)
            editor.SetColumnSpan(templateHost, 2)

            Dim resolvedHost As New Panel With {.Dock = DockStyle.Fill, .Padding = New Padding(0, 0, 12, 10)}
            Dim resolvedLabel As New Label With {
                .Dock = DockStyle.Top,
                .Height = 22,
                .Text = "Resolved Template File",
                .Font = New Font("Segoe UI Semibold", 9.5F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary
            }
            lblResolvedTemplatePath.Dock = DockStyle.Fill
            lblResolvedTemplatePath.Font = New Font("Consolas", 9.5F, FontStyle.Regular)
            lblResolvedTemplatePath.ForeColor = ThemePalette.TextMuted
            lblResolvedTemplatePath.TextAlign = ContentAlignment.TopLeft

            resolvedHost.Controls.Add(lblResolvedTemplatePath)
            resolvedHost.Controls.Add(resolvedLabel)
            editor.Controls.Add(resolvedHost, 0, 4)
            editor.SetColumnSpan(resolvedHost, 2)

            panel.Controls.Add(UiStyler.CreateScrollableHost(editor))
            panel.Controls.Add(title)
            Return panel
        End Function

        Private Function BuildMaintenanceCard() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)

            Dim title As New Label With {
                .Dock = DockStyle.Top,
                .Height = 30,
                .Text = "Maintenance and storage",
                .Font = New Font("Segoe UI Semibold", 14.0F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary
            }

            Dim shell As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .Padding = New Padding(0, 14, 0, 0)
            }
            shell.RowStyles.Add(New RowStyle(SizeType.Absolute, 54))
            shell.RowStyles.Add(New RowStyle(SizeType.Absolute, 254))
            shell.RowStyles.Add(New RowStyle(SizeType.Absolute, 164))
            shell.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            Dim intro As New Label With {
                .Dock = DockStyle.Fill,
                .Text = "Create verified SQLite backups, restore a selected backup with restart protection, optimize the live database, and open the operational folders used by invoices, reports, and logs.",
                .ForeColor = ThemePalette.TextMuted,
                .Font = New Font("Segoe UI", 9.75F, FontStyle.Regular),
                .TextAlign = ContentAlignment.TopLeft
            }
            shell.Controls.Add(intro, 0, 0)

            Dim pathsTable As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2,
                .RowCount = 7,
                .BackColor = Color.Transparent
            }
            pathsTable.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 152))
            pathsTable.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))

            AddPathRow(pathsTable, 0, "Data Root", lblDataRootPath)
            AddPathRow(pathsTable, 1, "Database File", lblDatabasePath)
            AddPathRow(pathsTable, 2, "Backups Folder", lblBackupsPath)
            AddPathRow(pathsTable, 3, "Invoices Folder", lblInvoicesPath)
            AddPathRow(pathsTable, 4, "Reports Folder", lblReportsPath)
            AddPathRow(pathsTable, 5, "Logs Folder", lblLogsPath)
            AddPathRow(pathsTable, 6, "Templates Folder", lblTemplateFolderPath)
            shell.Controls.Add(pathsTable, 0, 1)

            Dim buttonGrid As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 3,
                .RowCount = 3,
                .Margin = New Padding(0, 6, 0, 0)
            }
            buttonGrid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 33.3333F))
            buttonGrid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 33.3333F))
            buttonGrid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 33.3333F))
            buttonGrid.RowStyles.Add(New RowStyle(SizeType.Absolute, 48))
            buttonGrid.RowStyles.Add(New RowStyle(SizeType.Absolute, 48))
            buttonGrid.RowStyles.Add(New RowStyle(SizeType.Absolute, 48))

            ConfigureActionButton(btnCreateBackup, "Create Backup", AddressOf btnCreateBackup_Click, True)
            ConfigureActionButton(btnRestoreBackup, "Restore Backup", AddressOf btnRestoreBackup_Click, False)
            ConfigureActionButton(btnOptimizeDatabase, "Optimize Database", AddressOf btnOptimizeDatabase_Click, False)
            ConfigureActionButton(btnOpenBackups, "Open Backups", AddressOf btnOpenBackups_Click, False)
            ConfigureActionButton(btnOpenInvoices, "Open Invoices", AddressOf btnOpenInvoices_Click, False)
            ConfigureActionButton(btnOpenReports, "Open Reports", AddressOf btnOpenReports_Click, False)
            ConfigureActionButton(btnOpenLogs, "Open Logs", AddressOf btnOpenLogs_Click, False)

            buttonGrid.Controls.Add(btnCreateBackup, 0, 0)
            buttonGrid.Controls.Add(btnRestoreBackup, 1, 0)
            buttonGrid.Controls.Add(btnOptimizeDatabase, 2, 0)
            buttonGrid.Controls.Add(btnOpenBackups, 0, 1)
            buttonGrid.Controls.Add(btnOpenInvoices, 1, 1)
            buttonGrid.Controls.Add(btnOpenReports, 2, 1)
            buttonGrid.Controls.Add(btnOpenLogs, 0, 2)
            shell.Controls.Add(buttonGrid, 0, 2)

            Dim footerNote As New Label With {
                .Dock = DockStyle.Fill,
                .Text = "Restoring a backup replaces the active SQLite database after validation and creates a pre-restore safeguard copy automatically.",
                .ForeColor = ThemePalette.TextMuted,
                .Font = New Font("Segoe UI", 9.25F, FontStyle.Italic),
                .TextAlign = ContentAlignment.TopLeft
            }
            shell.Controls.Add(footerNote, 0, 3)

            panel.Controls.Add(UiStyler.CreateScrollableHost(shell))
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

        Private Shared Function CreatePathValueLabel() As Label
            Return New Label With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(10, 7, 10, 7),
                .BorderStyle = BorderStyle.FixedSingle,
                .BackColor = Color.White,
                .ForeColor = ThemePalette.TextMuted,
                .Font = New Font("Consolas", 9.0F, FontStyle.Regular),
                .AutoEllipsis = True,
                .TextAlign = ContentAlignment.MiddleLeft
            }
        End Function

        Private Sub AddPathRow(pathsTable As TableLayoutPanel, rowIndex As Integer, caption As String, valueLabel As Label)
            pathsTable.RowStyles.Add(New RowStyle(SizeType.Absolute, 36))

            Dim captionLabel As New Label With {
                .Dock = DockStyle.Fill,
                .Text = caption,
                .Font = New Font("Segoe UI Semibold", 9.5F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary,
                .TextAlign = ContentAlignment.MiddleLeft
            }

            pathsTable.Controls.Add(captionLabel, 0, rowIndex)
            pathsTable.Controls.Add(valueLabel, 1, rowIndex)
        End Sub

        Private Sub ConfigureToolbarButton(button As Button, text As String, handler As EventHandler, isPrimary As Boolean, width As Integer)
            If isPrimary Then
                UiStyler.StylePrimaryButton(button)
            Else
                UiStyler.StyleSecondaryButton(button)
            End If

            button.Text = text
            button.Width = width
            AddHandler button.Click, handler
        End Sub

        Private Sub ConfigureActionButton(button As Button, text As String, handler As EventHandler, isPrimary As Boolean)
            If isPrimary Then
                UiStyler.StylePrimaryButton(button)
            Else
                UiStyler.StyleSecondaryButton(button)
            End If

            button.Text = text
            button.Dock = DockStyle.Fill
            button.Margin = New Padding(0, 0, 10, 10)
            AddHandler button.Click, handler
        End Sub

        Private Sub ConfigureEditors()
            For Each editorTextBox As TextBox In New TextBox() {
                txtCompanyName,
                txtCompanyAddress,
                txtCompanyPhone,
                txtCompanyGstin,
                txtCompanyDrugLicense,
                txtInvoicePrefix,
                txtPurchasePrefix,
                txtReceiptPrefix,
                txtSupplierPaymentPrefix,
                txtCurrencySymbol,
                txtInvoiceTemplatePath
            }
                editorTextBox.BorderStyle = BorderStyle.FixedSingle
                UiStyler.StyleInput(editorTextBox)
            Next

            nudLowStockThreshold.Maximum = 1000000D
            nudLowStockThreshold.Minimum = 0D
            nudLowStockThreshold.Increment = 1D
            nudLowStockThreshold.Font = New Font("Segoe UI", 10.5F, FontStyle.Regular)
            nudLowStockThreshold.ThousandsSeparator = True
        End Sub

        Private Sub WireEvents()
            AddHandler Load, AddressOf FrmSettings_Load
            AddHandler btnBrowseTemplate.Click, AddressOf btnBrowseTemplate_Click
            AddHandler btnResetTemplate.Click, AddressOf btnResetTemplate_Click
            AddHandler txtInvoiceTemplatePath.TextChanged, AddressOf txtInvoiceTemplatePath_TextChanged
        End Sub

        Private Async Sub FrmSettings_Load(sender As Object, e As EventArgs)
            UpdatePathLabels()
            Await LoadSettingsAsync()
        End Sub

        Private Async Function LoadSettingsAsync() As Task
            SetBusy(True, "Loading settings...")

            Try
                Dim profile As AppSettingsProfile = Await _settingsService.LoadAsync()
                BindProfile(profile)
                ShowStatus("Settings loaded successfully.", False)
            Catch ex As Exception
                AppLogger.Error("Settings load failed.", ex)
                ShowStatus("Settings could not be loaded.", True)
            Finally
                SetBusy(False)
            End Try
        End Function

        Private Sub BindProfile(profile As AppSettingsProfile)
            txtCompanyName.Text = profile.CompanyName
            txtCompanyAddress.Text = profile.CompanyAddress
            txtCompanyPhone.Text = profile.CompanyPhone
            txtCompanyGstin.Text = profile.CompanyGstin
            txtCompanyDrugLicense.Text = profile.CompanyDrugLicense
            txtInvoicePrefix.Text = profile.InvoicePrefix
            txtPurchasePrefix.Text = profile.PurchasePrefix
            txtReceiptPrefix.Text = profile.ReceiptPrefix
            txtSupplierPaymentPrefix.Text = profile.SupplierPaymentPrefix
            nudLowStockThreshold.Value = Math.Min(nudLowStockThreshold.Maximum, profile.LowStockThreshold)
            txtCurrencySymbol.Text = profile.CurrencySymbol
            txtInvoiceTemplatePath.Text = profile.InvoiceTemplatePath
            UpdateResolvedTemplateLabel()
        End Sub

        Private Function ReadProfileFromForm() As AppSettingsProfile
            Return New AppSettingsProfile With {
                .CompanyName = txtCompanyName.Text,
                .CompanyAddress = txtCompanyAddress.Text,
                .CompanyPhone = txtCompanyPhone.Text,
                .CompanyGstin = txtCompanyGstin.Text,
                .CompanyDrugLicense = txtCompanyDrugLicense.Text,
                .InvoicePrefix = txtInvoicePrefix.Text,
                .PurchasePrefix = txtPurchasePrefix.Text,
                .ReceiptPrefix = txtReceiptPrefix.Text,
                .SupplierPaymentPrefix = txtSupplierPaymentPrefix.Text,
                .LowStockThreshold = Decimal.ToInt32(nudLowStockThreshold.Value),
                .CurrencySymbol = txtCurrencySymbol.Text,
                .InvoiceTemplatePath = txtInvoiceTemplatePath.Text
            }
        End Function

        Private Async Sub btnReload_Click(sender As Object, e As EventArgs)
            Await LoadSettingsAsync()
        End Sub

        Private Async Sub btnSave_Click(sender As Object, e As EventArgs)
            Dim profile As AppSettingsProfile = ReadProfileFromForm()
            SetBusy(True, "Saving settings...")
            Dim result As EntityOperationResult = Await _settingsService.SaveAsync(profile)
            SetBusy(False)
            ShowStatus(result.Message, Not result.IsSuccess)

            If result.IsSuccess Then
                BindProfile(profile)
            End If
        End Sub

        Private Sub btnBrowseTemplate_Click(sender As Object, e As EventArgs)
            Dim currentResolvedPath As String = _settingsService.GetResolvedTemplatePath(ReadProfileFromForm())

            Using dialog As New SaveFileDialog()
                dialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
                dialog.FileName = Path.GetFileName(currentResolvedPath)
                dialog.InitialDirectory = Path.GetDirectoryName(currentResolvedPath)

                If dialog.ShowDialog(Me) = DialogResult.OK Then
                    txtInvoiceTemplatePath.Text = AppPaths.ToBaseRelativePath(dialog.FileName)
                    UpdateResolvedTemplateLabel()
                End If
            End Using
        End Sub

        Private Sub btnResetTemplate_Click(sender As Object, e As EventArgs)
            txtInvoiceTemplatePath.Text = _settingsService.GetDefaultTemplateSettingValue()
            UpdateResolvedTemplateLabel()
            ShowStatus("Template path reset to the default GST invoice template.", False)
        End Sub

        Private Sub txtInvoiceTemplatePath_TextChanged(sender As Object, e As EventArgs)
            UpdateResolvedTemplateLabel()
        End Sub

        Private Sub btnOpenTemplates_Click(sender As Object, e As EventArgs)
            Dim profile As AppSettingsProfile = ReadProfileFromForm()
            Dim resolvedPath As String = _settingsService.GetResolvedTemplatePath(profile)
            Dim folderPath As String = Path.GetDirectoryName(resolvedPath)
            OpenFolderPath(folderPath, "Opened template folder.", "Template folder could not be opened.")
        End Sub

        Private Sub btnOpenDataRoot_Click(sender As Object, e As EventArgs)
            OpenFolderPath(AppPaths.DataRootDirectory, "Opened application data folder.", "Application data folder could not be opened.")
        End Sub

        Private Async Sub btnCreateBackup_Click(sender As Object, e As EventArgs)
            SetBusy(True, "Creating database backup...")
            Dim result As FileOperationResult = Await _maintenanceService.CreateDatabaseBackupAsync()
            SetBusy(False)

            If result.IsSuccess Then
                ShowStatus($"{result.Message} File: {result.FilePath}", False)
                UpdatePathLabels()
            Else
                ShowStatus(result.Message, True)
            End If
        End Sub

        Private Async Sub btnRestoreBackup_Click(sender As Object, e As EventArgs)
            Using dialog As New OpenFileDialog()
                dialog.Filter = "SQLite Database (*.db)|*.db|All Files (*.*)|*.*"
                dialog.InitialDirectory = _maintenanceService.GetDefaultBackupDirectory()
                dialog.Multiselect = False

                If dialog.ShowDialog(Me) <> DialogResult.OK Then
                    Return
                End If

                Dim restorePath As String = dialog.FileName
                Dim confirmResult As DialogResult =
                    MessageBox.Show(
                        $"Restore the backup '{Path.GetFileName(restorePath)}'? The current database will be replaced and the application will need to restart.",
                        "Confirm Restore",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning)

                If confirmResult <> DialogResult.Yes Then
                    Return
                End If

                SetBusy(True, "Restoring database backup...")
                Dim result As EntityOperationResult = Await _maintenanceService.RestoreDatabaseBackupAsync(restorePath)
                SetBusy(False)
                ShowStatus(result.Message, Not result.IsSuccess)

                If result.IsSuccess Then
                    Dim restartResult As DialogResult =
                        MessageBox.Show(
                            result.Message & Environment.NewLine & Environment.NewLine & "Restart the application now?",
                            "Restart Required",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Information)

                    If restartResult = DialogResult.Yes Then
                        Application.Restart()
                        Environment.Exit(0)
                    End If
                End If
            End Using
        End Sub

        Private Async Sub btnOptimizeDatabase_Click(sender As Object, e As EventArgs)
            SetBusy(True, "Optimizing database...")
            Dim result As EntityOperationResult = Await _maintenanceService.OptimizeDatabaseAsync()
            SetBusy(False)
            ShowStatus(result.Message, Not result.IsSuccess)
        End Sub

        Private Sub btnOpenBackups_Click(sender As Object, e As EventArgs)
            OpenFolderPath(AppPaths.BackupsDirectory, "Opened backup folder.", "Backup folder could not be opened.")
        End Sub

        Private Sub btnOpenInvoices_Click(sender As Object, e As EventArgs)
            OpenFolderPath(AppPaths.GeneratedInvoicesDirectory, "Opened invoice folder.", "Invoice folder could not be opened.")
        End Sub

        Private Sub btnOpenReports_Click(sender As Object, e As EventArgs)
            OpenFolderPath(AppPaths.ReportsDirectory, "Opened report folder.", "Report folder could not be opened.")
        End Sub

        Private Sub btnOpenLogs_Click(sender As Object, e As EventArgs)
            OpenFolderPath(AppPaths.LogsDirectory, "Opened logs folder.", "Logs folder could not be opened.")
        End Sub

        Private Sub OpenFolderPath(folderPath As String, successMessage As String, errorMessage As String)
            Try
                Directory.CreateDirectory(folderPath)
                Process.Start(New ProcessStartInfo With {
                    .FileName = folderPath,
                    .UseShellExecute = True
                })
                ShowStatus(successMessage, False)
            Catch ex As Exception
                AppLogger.Error(errorMessage, ex)
                ShowStatus(errorMessage, True)
            End Try
        End Sub

        Private Sub btnClose_Click(sender As Object, e As EventArgs)
            Close()
        End Sub

        Private Sub SetBusy(isBusy As Boolean, Optional message As String = "")
            _isBusy = isBusy

            txtCompanyName.Enabled = Not isBusy
            txtCompanyAddress.Enabled = Not isBusy
            txtCompanyPhone.Enabled = Not isBusy
            txtCompanyGstin.Enabled = Not isBusy
            txtCompanyDrugLicense.Enabled = Not isBusy
            txtInvoicePrefix.Enabled = Not isBusy
            txtPurchasePrefix.Enabled = Not isBusy
            txtReceiptPrefix.Enabled = Not isBusy
            txtSupplierPaymentPrefix.Enabled = Not isBusy
            nudLowStockThreshold.Enabled = Not isBusy
            txtCurrencySymbol.Enabled = Not isBusy
            txtInvoiceTemplatePath.Enabled = Not isBusy
            btnBrowseTemplate.Enabled = Not isBusy
            btnResetTemplate.Enabled = Not isBusy
            btnReload.Enabled = Not isBusy
            btnSave.Enabled = Not isBusy
            btnOpenTemplates.Enabled = Not isBusy
            btnOpenDataRoot.Enabled = Not isBusy
            btnCreateBackup.Enabled = Not isBusy
            btnRestoreBackup.Enabled = Not isBusy
            btnOptimizeDatabase.Enabled = Not isBusy
            btnOpenBackups.Enabled = Not isBusy
            btnOpenInvoices.Enabled = Not isBusy
            btnOpenReports.Enabled = Not isBusy
            btnOpenLogs.Enabled = Not isBusy
            btnClose.Enabled = Not isBusy

            If isBusy Then
                lblStatus.ForeColor = ThemePalette.TextMuted
                lblStatus.Text = message
            End If
        End Sub

        Private Sub ShowStatus(message As String, isError As Boolean)
            lblStatus.ForeColor = If(isError, ThemePalette.DangerRed, ThemePalette.AccentGreen)
            lblStatus.Text = message
        End Sub

        Private Sub UpdateResolvedTemplateLabel()
            Dim profile As AppSettingsProfile = ReadProfileFromForm()
            lblResolvedTemplatePath.Text = _settingsService.GetResolvedTemplatePath(profile)
            UpdatePathLabels()
        End Sub

        Private Sub UpdatePathLabels()
            AppPaths.EnsureDirectories()

            lblDataRootPath.Text = AppPaths.DataRootDirectory
            lblDatabasePath.Text = AppPaths.DatabaseFilePath
            lblBackupsPath.Text = AppPaths.BackupsDirectory
            lblInvoicesPath.Text = AppPaths.GeneratedInvoicesDirectory
            lblReportsPath.Text = AppPaths.ReportsDirectory
            lblLogsPath.Text = AppPaths.LogsDirectory

            Dim profile As AppSettingsProfile = ReadProfileFromForm()
            Dim resolvedTemplatePath As String = _settingsService.GetResolvedTemplatePath(profile)
            lblTemplateFolderPath.Text = Path.GetDirectoryName(resolvedTemplatePath)
        End Sub

        Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
            Select Case keyData
                Case Keys.Control Or Keys.S
                    btnSave.PerformClick()
                    Return True
                Case Keys.Control Or Keys.B
                    btnCreateBackup.PerformClick()
                    Return True
                Case Keys.F5
                    btnReload.PerformClick()
                    Return True
                Case Keys.Escape
                    Close()
                    Return True
            End Select

            Return MyBase.ProcessCmdKey(msg, keyData)
        End Function

    End Class

End Namespace
