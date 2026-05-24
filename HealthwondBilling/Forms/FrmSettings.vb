Imports HealthwondBilling.Models
Imports HealthwondBilling.Services
Imports HealthwondBilling.Utilities
Imports System.Diagnostics
Imports System.IO

Namespace Forms

    Public Class FrmSettings
        Inherits Form

        Private ReadOnly _settingsService As SettingsService
        Private ReadOnly txtCompanyName As New TextBox()
        Private ReadOnly txtCompanyAddress As New TextBox()
        Private ReadOnly txtCompanyPhone As New TextBox()
        Private ReadOnly txtCompanyGstin As New TextBox()
        Private ReadOnly txtCompanyDrugLicense As New TextBox()
        Private ReadOnly txtInvoicePrefix As New TextBox()
        Private ReadOnly txtPurchasePrefix As New TextBox()
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
        Private ReadOnly lblStatus As New Label()
        Private ReadOnly lblResolvedTemplatePath As New Label()

        Private _isBusy As Boolean

        Public Sub New(settingsService As SettingsService)
            _settingsService = settingsService

            Text = "Healthwond Billing System - Settings"
            StartPosition = FormStartPosition.CenterParent
            Size = New Size(1340, 840)
            MinimumSize = New Size(1200, 760)
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
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 82))
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
                .Text = "Maintain company invoice identity, numbering prefixes, low stock threshold, currency display, and the GST invoice template used by export workflows.",
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
            split.Panel2.Controls.Add(BuildSystemCard())
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
                .Text = "These values are used by Excel invoice generation, PDF export, and print preview.",
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
            editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 120))
            editor.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            editor.Controls.Add(CreateInputHost("Invoice Prefix", txtInvoicePrefix), 0, 0)
            editor.Controls.Add(CreateInputHost("Purchase Prefix", txtPurchasePrefix), 1, 0)
            editor.Controls.Add(CreateInputHost("Currency Symbol", txtCurrencySymbol), 0, 1)
            editor.Controls.Add(CreateInputHost("Low Stock Threshold", nudLowStockThreshold), 1, 1)

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
            editor.Controls.Add(templateHost, 0, 2)
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
            editor.Controls.Add(resolvedHost, 0, 3)
            editor.SetColumnSpan(resolvedHost, 2)

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

        Private Sub ConfigureEditors()
            For Each editorTextBox As TextBox In New TextBox() {txtCompanyName, txtCompanyAddress, txtCompanyPhone, txtCompanyGstin, txtCompanyDrugLicense, txtInvoicePrefix, txtPurchasePrefix, txtCurrencySymbol, txtInvoiceTemplatePath}
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
            Try
                Dim profile As AppSettingsProfile = ReadProfileFromForm()
                Dim resolvedPath As String = _settingsService.GetResolvedTemplatePath(profile)
                Dim folderPath As String = Path.GetDirectoryName(resolvedPath)
                Directory.CreateDirectory(folderPath)
                Process.Start(New ProcessStartInfo With {
                    .FileName = folderPath,
                    .UseShellExecute = True
                })
                ShowStatus("Opened template folder.", False)
            Catch ex As Exception
                AppLogger.Error("Template folder open failed.", ex)
                ShowStatus("Template folder could not be opened.", True)
            End Try
        End Sub

        Private Sub btnOpenDataRoot_Click(sender As Object, e As EventArgs)
            Try
                Process.Start(New ProcessStartInfo With {
                    .FileName = AppPaths.DataRootDirectory,
                    .UseShellExecute = True
                })
                ShowStatus("Opened application data folder.", False)
            Catch ex As Exception
                AppLogger.Error("Application data folder open failed.", ex)
                ShowStatus("Application data folder could not be opened.", True)
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
            nudLowStockThreshold.Enabled = Not isBusy
            txtCurrencySymbol.Enabled = Not isBusy
            txtInvoiceTemplatePath.Enabled = Not isBusy
            btnBrowseTemplate.Enabled = Not isBusy
            btnResetTemplate.Enabled = Not isBusy
            btnReload.Enabled = Not isBusy
            btnSave.Enabled = Not isBusy
            btnOpenTemplates.Enabled = Not isBusy
            btnOpenDataRoot.Enabled = Not isBusy
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
        End Sub

        Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
            Select Case keyData
                Case Keys.Control Or Keys.S
                    btnSave.PerformClick()
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
