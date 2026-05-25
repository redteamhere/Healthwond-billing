Imports HealthwondBilling.Models
Imports HealthwondBilling.Services
Imports HealthwondBilling.Utilities
Imports System.Linq

Namespace Forms

    Public Class FrmLogin
        Inherits Form

        Private Const ActionCreateCompany As String = "create-company"
        Private Const ActionRestoreDefaultDemo As String = "restore-default-demo"
        Private Const ActionRestoreFromBackup As String = "restore-from-backup"
        Private Const ActionDeleteSavedPassword As String = "delete-saved-password"
        Private Const ActionChangeOperatorPowers As String = "change-operator-powers"
        Private Const ActionChangeErpVersion As String = "change-erp-version"

        Private ReadOnly _runtimeFactory As AppRuntimeFactory
        Private ReadOnly _companyWorkspaceService As CompanyWorkspaceService
        Private ReadOnly _savedCredentialService As SavedCredentialService
        Private ReadOnly _clockTimer As New Timer()
        Private ReadOnly _startupActions As IReadOnlyList(Of StartupActionItem)

        Private ReadOnly lstCompanies As New ListBox()
        Private ReadOnly lstStartupActions As New ListBox()
        Private ReadOnly txtUsername As New TextBox()
        Private ReadOnly txtPassword As New TextBox()
        Private ReadOnly chkShowPassword As New CheckBox()
        Private ReadOnly btnLogin As New Button()
        Private ReadOnly btnDemoAdmin As New Button()
        Private ReadOnly btnDemoStaff As New Button()
        Private ReadOnly btnRunAction As New Button()
        Private ReadOnly btnRefreshCompanies As New Button()
        Private ReadOnly btnExit As New Button()
        Private ReadOnly lblStatus As New Label()
        Private ReadOnly lblTopClock As New Label()
        Private ReadOnly lblTopTitle As New Label()
        Private ReadOnly lblActionDescription As New Label()
        Private ReadOnly lblSelectedCompany As New Label()
        Private ReadOnly lblSelectedCompanyValue As New Label()
        Private ReadOnly lblDateValue As New Label()
        Private ReadOnly lblDayValue As New Label()
        Private ReadOnly lblTimeValue As New Label()
        Private ReadOnly lblShortcutHints As New Label()
        Private ReadOnly lblEditionBadge As New Label()
        Private ReadOnly lblAuthorizedSummary As New Label()

        Private _workspaceShadow As Panel
        Private _workspaceCard As Panel
        Private _workspaceHost As Panel
        Private _startupProfile As New StartupShellProfile()
        Private _companies As New List(Of CompanyWorkspaceRecord)()

        Public Sub New(runtimeFactory As AppRuntimeFactory, companyWorkspaceService As CompanyWorkspaceService, savedCredentialService As SavedCredentialService)
            _runtimeFactory = runtimeFactory
            _companyWorkspaceService = companyWorkspaceService
            _savedCredentialService = savedCredentialService
            _startupActions = New List(Of StartupActionItem) From {
                New StartupActionItem(ActionCreateCompany, "CREATE  NEW  COMPANY", "Create a new company workspace with its own SQLite data file and seeded ERP setup."),
                New StartupActionItem(ActionRestoreDefaultDemo, "RESTORE  DEFAULT  DEMONSTRATION", "Reset the built-in demo company so training data and default operators are available again."),
                New StartupActionItem(ActionRestoreFromBackup, "RESTORE  FROM  MY  DEMONSTRATION", "Import a demonstration or backup SQLite database as a separate company workspace."),
                New StartupActionItem(ActionDeleteSavedPassword, "DELETE  SAVED  PASSWORD", "Clear the saved operator credentials for the currently selected company."),
                New StartupActionItem(ActionChangeOperatorPowers, "CHANGE  OPERATOR  POWERS", "Maintain operator roles, active status, and passwords for the selected company."),
                New StartupActionItem(ActionChangeErpVersion, "CHANGE  ERP  VERSION", "Update the startup shell edition label shown across the classic ERP-style landing screen.")
            }

            Text = "Healthwond Billing System - Login"
            StartPosition = FormStartPosition.CenterScreen
            MinimumSize = New Size(1420, 900)
            Size = New Size(1600, 940)
            BackColor = ThemePalette.ClassicShellBlue
            Font = New Font("Segoe UI", 10.0F, FontStyle.Regular)
            KeyPreview = True

            BuildLayout()
            ConfigureControls()
            WireEvents()

            AcceptButton = btnLogin
            CancelButton = btnExit
        End Sub

        Private Sub BuildLayout()
            Dim root As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 4,
                .BackColor = ThemePalette.ClassicShellBlue
            }
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 42))
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 34))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 194))

            root.Controls.Add(BuildTopBar(), 0, 0)
            root.Controls.Add(BuildWorkspaceHost(), 0, 1)
            root.Controls.Add(BuildHintBar(), 0, 2)
            root.Controls.Add(BuildFooterPanel(), 0, 3)
            Controls.Add(root)
        End Sub

        Private Function BuildTopBar() As Control
            Dim panel As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = ThemePalette.ClassicShellBlue,
                .Padding = New Padding(10, 4, 10, 4)
            }

            lblTopTitle.Dock = DockStyle.Left
            lblTopTitle.Width = 1080
            lblTopTitle.Font = New Font("Segoe UI", 9.75F, FontStyle.Regular)
            lblTopTitle.ForeColor = Color.Black
            lblTopTitle.TextAlign = ContentAlignment.MiddleLeft

            lblTopClock.Dock = DockStyle.Right
            lblTopClock.Width = 170
            lblTopClock.Font = New Font("Consolas", 12.0F, FontStyle.Bold)
            lblTopClock.ForeColor = Color.White
            lblTopClock.BackColor = ThemePalette.ClassicShellGreen
            lblTopClock.TextAlign = ContentAlignment.MiddleCenter

            panel.Controls.Add(lblTopClock)
            panel.Controls.Add(lblTopTitle)
            Return panel
        End Function

        Private Function BuildWorkspaceHost() As Control
            _workspaceHost = New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = ThemePalette.ClassicShellCream,
                .Padding = New Padding(18)
            }

            _workspaceShadow = New Panel With {
                .BackColor = ThemePalette.ClassicShellShadow,
                .Size = New Size(1500, 688)
            }

            _workspaceCard = BuildWorkspaceCard()
            _workspaceHost.Controls.Add(_workspaceShadow)
            _workspaceHost.Controls.Add(_workspaceCard)
            _workspaceShadow.SendToBack()
            _workspaceCard.BringToFront()
            AddHandler _workspaceHost.Resize, AddressOf WorkspaceHost_Resize
            Return _workspaceHost
        End Function

        Private Function BuildWorkspaceCard() As Panel
            Dim card As New Panel With {
                .BackColor = ThemePalette.ClassicShellCream,
                .Size = New Size(1500, 688),
                .Padding = New Padding(18),
                .BorderStyle = BorderStyle.FixedSingle
            }

            Dim headerBand As New Label With {
                .Dock = DockStyle.Top,
                .Height = 34,
                .BackColor = ThemePalette.ClassicShellGreen,
                .ForeColor = Color.White,
                .Font = New Font("Segoe UI", 18.0F, FontStyle.Regular),
                .Text = "LIST OF COMPANIES",
                .Padding = New Padding(16, 0, 0, 0),
                .TextAlign = ContentAlignment.MiddleLeft
            }

            Dim contentShell As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2,
                .RowCount = 2,
                .BackColor = ThemePalette.ClassicShellCream,
                .Padding = New Padding(18, 16, 18, 16)
            }
            contentShell.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 420))
            contentShell.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            contentShell.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            contentShell.RowStyles.Add(New RowStyle(SizeType.Absolute, 88))

            contentShell.Controls.Add(BuildCompanyListPanel(), 0, 0)
            contentShell.Controls.Add(BuildControlZone(), 1, 0)
            contentShell.Controls.Add(BuildSelectedCompanyPrompt(), 0, 1)
            contentShell.Controls.Add(BuildDateInfoPanel(), 1, 1)

            card.Controls.Add(contentShell)
            card.Controls.Add(headerBand)
            Return card
        End Function

        Private Function BuildCompanyListPanel() As Control
            Dim panel As New Panel With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(18),
                .BackColor = ThemePalette.ClassicShellCream
            }

            Dim titleBar As New Label With {
                .Dock = DockStyle.Top,
                .Height = 34,
                .BackColor = ThemePalette.ClassicShellDarkGreen,
                .ForeColor = Color.White,
                .Font = New Font("Consolas", 15.0F, FontStyle.Bold),
                .Text = " AVAILABLE COMPANIES ",
                .TextAlign = ContentAlignment.MiddleLeft
            }

            lstCompanies.Dock = DockStyle.Fill
            lstCompanies.BackColor = Color.FromArgb(250, 250, 242)
            lstCompanies.ForeColor = Color.Black
            lstCompanies.BorderStyle = BorderStyle.FixedSingle
            lstCompanies.Font = New Font("Consolas", 12.0F, FontStyle.Bold)
            lstCompanies.ItemHeight = 24

            Dim helpText As New Label With {
                .Dock = DockStyle.Bottom,
                .Height = 54,
                .ForeColor = ThemePalette.TextPrimary,
                .Font = New Font("Consolas", 10.0F, FontStyle.Regular),
                .Text = "Use Up/Down to choose the active company." & Environment.NewLine & "Saved credentials load automatically for the selected company.",
                .TextAlign = ContentAlignment.MiddleLeft
            }

            panel.Controls.Add(lstCompanies)
            panel.Controls.Add(helpText)
            panel.Controls.Add(titleBar)
            Return panel
        End Function

        Private Function BuildControlZone() As Control
            Dim layout As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 2,
                .BackColor = ThemePalette.ClassicShellCream
            }
            layout.RowStyles.Add(New RowStyle(SizeType.Percent, 58.0F))
            layout.RowStyles.Add(New RowStyle(SizeType.Percent, 42.0F))

            Dim startupCard As Control = BuildStartupActionCard()
            startupCard.Margin = New Padding(0, 0, 0, 12)

            Dim loginCard As Control = BuildLoginCard()
            loginCard.Margin = New Padding(0, 12, 0, 0)

            layout.Controls.Add(startupCard, 0, 0)
            layout.Controls.Add(loginCard, 0, 1)
            Return layout
        End Function

        Private Function BuildStartupActionCard() As Control
            Dim panel As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = ThemePalette.ClassicShellPaleBlue,
                .BorderStyle = BorderStyle.FixedSingle,
                .Padding = New Padding(18, 14, 18, 14)
            }

            Dim titleBar As New Label With {
                .Dock = DockStyle.Top,
                .Height = 32,
                .BackColor = ThemePalette.ClassicShellGreen,
                .ForeColor = Color.White,
                .Font = New Font("Segoe UI", 16.0F, FontStyle.Regular),
                .Padding = New Padding(18, 0, 0, 0),
                .Text = "COMPANY CREATION",
                .TextAlign = ContentAlignment.MiddleLeft
            }

            Dim body As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .Padding = New Padding(22, 18, 22, 10)
            }
            body.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            body.RowStyles.Add(New RowStyle(SizeType.Absolute, 60))
            body.RowStyles.Add(New RowStyle(SizeType.Absolute, 54))

            lstStartupActions.Dock = DockStyle.Fill
            lstStartupActions.BackColor = ThemePalette.ClassicShellPaleBlue
            lstStartupActions.BorderStyle = BorderStyle.None
            lstStartupActions.Font = New Font("Consolas", 18.0F, FontStyle.Regular)
            lstStartupActions.ForeColor = Color.Black
            lstStartupActions.ItemHeight = 34

            lblActionDescription.Dock = DockStyle.Fill
            lblActionDescription.Font = New Font("Segoe UI", 9.75F, FontStyle.Regular)
            lblActionDescription.ForeColor = ThemePalette.ClassicShellBorder
            lblActionDescription.TextAlign = ContentAlignment.MiddleLeft

            Dim actionButtons As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.LeftToRight,
                .WrapContents = False,
                .BackColor = Color.Transparent
            }

            ConfigureClassicActionButton(btnRunAction, "Run Selected", True, 146)
            ConfigureClassicActionButton(btnRefreshCompanies, "Refresh Companies", False, 170)
            actionButtons.Controls.Add(btnRunAction)
            actionButtons.Controls.Add(btnRefreshCompanies)

            body.Controls.Add(lstStartupActions, 0, 0)
            body.Controls.Add(lblActionDescription, 0, 1)
            body.Controls.Add(actionButtons, 0, 2)

            panel.Controls.Add(body)
            panel.Controls.Add(titleBar)
            Return panel
        End Function

        Private Function BuildLoginCard() As Control
            Dim panel As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = ThemePalette.ClassicShellPaleBlue,
                .BorderStyle = BorderStyle.FixedSingle,
                .Padding = New Padding(18, 14, 18, 14)
            }

            Dim titleBar As New Label With {
                .Dock = DockStyle.Top,
                .Height = 32,
                .BackColor = ThemePalette.ClassicShellGreen,
                .ForeColor = Color.White,
                .Font = New Font("Segoe UI", 16.0F, FontStyle.Regular),
                .Padding = New Padding(18, 0, 0, 0),
                .Text = "OPERATOR ACCESS",
                .TextAlign = ContentAlignment.MiddleLeft
            }

            Dim content As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2,
                .Padding = New Padding(18, 14, 18, 8)
            }
            content.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0F))
            content.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0F))
            content.RowStyles.Add(New RowStyle(SizeType.Absolute, 34))
            content.RowStyles.Add(New RowStyle(SizeType.Absolute, 66))
            content.RowStyles.Add(New RowStyle(SizeType.Absolute, 66))
            content.RowStyles.Add(New RowStyle(SizeType.Absolute, 32))
            content.RowStyles.Add(New RowStyle(SizeType.Absolute, 52))
            content.RowStyles.Add(New RowStyle(SizeType.Absolute, 32))

            lblSelectedCompany.Dock = DockStyle.Fill
            lblSelectedCompany.Font = New Font("Consolas", 11.0F, FontStyle.Regular)
            lblSelectedCompany.ForeColor = ThemePalette.ClassicShellBorder
            lblSelectedCompany.TextAlign = ContentAlignment.MiddleLeft
            lblSelectedCompany.Text = "Company :"

            Dim operatorHint As New Label With {
                .Dock = DockStyle.Fill,
                .Font = New Font("Consolas", 10.5F, FontStyle.Regular),
                .ForeColor = ThemePalette.ClassicShellBorder,
                .Text = "Use Demo Admin or Demo Staff for seeded startup access.",
                .TextAlign = ContentAlignment.MiddleRight
            }

            content.Controls.Add(lblSelectedCompany, 0, 0)
            content.Controls.Add(operatorHint, 1, 0)

            txtUsername.BorderStyle = BorderStyle.FixedSingle
            txtUsername.Font = New Font("Consolas", 13.0F, FontStyle.Bold)
            txtPassword.BorderStyle = BorderStyle.FixedSingle
            txtPassword.Font = New Font("Consolas", 13.0F, FontStyle.Bold)

            content.Controls.Add(CreateClassicInputHost("Operator Name", txtUsername), 0, 1)
            content.Controls.Add(CreateClassicInputHost("Password", txtPassword), 1, 1)

            Dim demoButtons As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.LeftToRight,
                .WrapContents = False
            }
            ConfigureClassicActionButton(btnDemoAdmin, "Demo Admin", False, 136)
            ConfigureClassicActionButton(btnDemoStaff, "Demo Staff", False, 136)
            demoButtons.Controls.Add(btnDemoAdmin)
            demoButtons.Controls.Add(btnDemoStaff)
            content.Controls.Add(demoButtons, 0, 2)

            chkShowPassword.Text = "Show Password"
            chkShowPassword.AutoSize = True
            chkShowPassword.Font = New Font("Consolas", 10.0F, FontStyle.Regular)
            chkShowPassword.ForeColor = ThemePalette.ClassicShellBorder
            chkShowPassword.BackColor = Color.Transparent
            chkShowPassword.Dock = DockStyle.Left
            content.Controls.Add(chkShowPassword, 1, 2)

            Dim buttonRow As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.RightToLeft,
                .WrapContents = False
            }
            ConfigureClassicActionButton(btnLogin, "Sign In", True, 140)
            ConfigureClassicActionButton(btnExit, "Exit", False, 110)
            buttonRow.Controls.Add(btnLogin)
            buttonRow.Controls.Add(btnExit)
            content.Controls.Add(buttonRow, 0, 4)
            content.SetColumnSpan(buttonRow, 2)

            Dim optionsLabel As New Label With {
                .Dock = DockStyle.Fill,
                .Font = New Font("Consolas", 10.5F, FontStyle.Regular),
                .ForeColor = ThemePalette.ClassicShellBorder,
                .Text = "Enter/F5 Sign In    F8 Show Password    Esc Exit",
                .TextAlign = ContentAlignment.MiddleLeft
            }
            content.Controls.Add(optionsLabel, 0, 5)
            content.SetColumnSpan(optionsLabel, 2)

            panel.Controls.Add(content)
            panel.Controls.Add(titleBar)
            Return panel
        End Function

        Private Function CreateClassicInputHost(labelText As String, editorControl As Control) As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .Padding = New Padding(0, 0, 10, 10)}

            Dim caption As New Label With {
                .Dock = DockStyle.Top,
                .Height = 24,
                .Font = New Font("Consolas", 11.0F, FontStyle.Bold),
                .ForeColor = ThemePalette.ClassicShellBorder,
                .Text = labelText
            }

            editorControl.Dock = DockStyle.Top
            editorControl.Height = 36

            panel.Controls.Add(editorControl)
            panel.Controls.Add(caption)
            Return panel
        End Function

        Private Function BuildSelectedCompanyPrompt() As Control
            Dim panel As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = ThemePalette.ClassicShellCream,
                .Padding = New Padding(18, 18, 18, 0)
            }

            Dim labelText As New Label With {
                .Dock = DockStyle.Left,
                .Width = 240,
                .Font = New Font("Consolas", 18.0F, FontStyle.Regular),
                .ForeColor = Color.Black,
                .Text = "Selected Company ?",
                .TextAlign = ContentAlignment.MiddleLeft
            }

            lblSelectedCompanyValue.Dock = DockStyle.Fill
            lblSelectedCompanyValue.Font = New Font("Consolas", 18.0F, FontStyle.Bold)
            lblSelectedCompanyValue.ForeColor = ThemePalette.ClassicShellDarkGreen
            lblSelectedCompanyValue.BorderStyle = BorderStyle.FixedSingle
            lblSelectedCompanyValue.Padding = New Padding(10, 0, 10, 0)
            lblSelectedCompanyValue.TextAlign = ContentAlignment.MiddleLeft

            panel.Controls.Add(lblSelectedCompanyValue)
            panel.Controls.Add(labelText)
            Return panel
        End Function

        Private Function BuildDateInfoPanel() As Control
            Dim panel As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = ThemePalette.ClassicShellCream,
                .Padding = New Padding(20, 12, 12, 0)
            }

            Dim table As New TableLayoutPanel With {
                .Dock = DockStyle.Right,
                .Width = 420,
                .ColumnCount = 2
            }
            table.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 120))
            table.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            table.RowStyles.Add(New RowStyle(SizeType.Absolute, 30))
            table.RowStyles.Add(New RowStyle(SizeType.Absolute, 30))
            table.RowStyles.Add(New RowStyle(SizeType.Absolute, 30))

            table.Controls.Add(CreateClassicInfoLabel("Date :"), 0, 0)
            table.Controls.Add(CreateClassicValueLabel(lblDateValue), 1, 0)
            table.Controls.Add(CreateClassicInfoLabel("Day  :"), 0, 1)
            table.Controls.Add(CreateClassicValueLabel(lblDayValue), 1, 1)
            table.Controls.Add(CreateClassicInfoLabel("Time :"), 0, 2)
            table.Controls.Add(CreateClassicValueLabel(lblTimeValue), 1, 2)

            panel.Controls.Add(table)
            Return panel
        End Function

        Private Function CreateClassicInfoLabel(text As String) As Label
            Return New Label With {
                .Dock = DockStyle.Fill,
                .Font = New Font("Segoe UI", 16.0F, FontStyle.Regular),
                .ForeColor = Color.Black,
                .Text = text,
                .TextAlign = ContentAlignment.MiddleLeft
            }
        End Function

        Private Function CreateClassicValueLabel(target As Label) As Label
            target.Dock = DockStyle.Fill
            target.Font = New Font("Segoe UI", 17.0F, FontStyle.Regular)
            target.ForeColor = Color.Black
            target.TextAlign = ContentAlignment.MiddleLeft
            Return target
        End Function

        Private Function BuildHintBar() As Control
            Dim panel As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = ThemePalette.ClassicShellHintBar,
                .Padding = New Padding(10, 0, 10, 0)
            }

            lblShortcutHints.Dock = DockStyle.Fill
            lblShortcutHints.Font = New Font("Segoe UI", 10.5F, FontStyle.Regular)
            lblShortcutHints.ForeColor = Color.White
            lblShortcutHints.Text = "F2-New Company   F3-Default Demo   F4-Import Demo   F5-Sign In   F6-Delete Saved Password   F7-Operator Powers   F8-Show Password   F9-ERP Version"
            lblShortcutHints.TextAlign = ContentAlignment.MiddleLeft

            panel.Controls.Add(lblShortcutHints)
            Return panel
        End Function

        Private Function BuildFooterPanel() As Control
            Dim panel As New Panel With {
                .Dock = DockStyle.Fill,
                .BackColor = ThemePalette.ClassicShellFooter,
                .Padding = New Padding(16, 10, 16, 10)
            }

            Dim layout As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 3
            }
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 42.0F))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 22.0F))
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 36.0F))

            Dim supportLabel As New Label With {
                .Dock = DockStyle.Fill,
                .Font = New Font("Segoe UI", 10.5F, FontStyle.Regular),
                .ForeColor = ThemePalette.TextPrimary,
                .Text = "Developed & Managed By :" & Environment.NewLine &
                        "HEALTHWOND BILLING SYSTEM" & Environment.NewLine &
                        "Support : support@healthwond.local" & Environment.NewLine &
                        "Desktop ERP Workspace for Pharmaceutical Billing",
                .TextAlign = ContentAlignment.TopLeft
            }

            Dim centerPanel As New Panel With {.Dock = DockStyle.Fill, .BackColor = ThemePalette.ClassicShellFooter}
            lblEditionBadge.Size = New Size(190, 116)
            lblEditionBadge.BorderStyle = BorderStyle.FixedSingle
            lblEditionBadge.BackColor = ThemePalette.ClassicShellCream
            lblEditionBadge.Font = New Font("Segoe UI", 13.0F, FontStyle.Bold)
            lblEditionBadge.ForeColor = ThemePalette.BrandBlue
            lblEditionBadge.TextAlign = ContentAlignment.MiddleCenter
            centerPanel.Controls.Add(lblEditionBadge)
            AddHandler centerPanel.Resize,
                Sub()
                    lblEditionBadge.Location = New Point(Math.Max((centerPanel.ClientSize.Width - lblEditionBadge.Width) \ 2, 0), Math.Max((centerPanel.ClientSize.Height - lblEditionBadge.Height) \ 2, 0))
                End Sub

            lblAuthorizedSummary.Dock = DockStyle.Fill
            lblAuthorizedSummary.Font = New Font("Segoe UI", 13.0F, FontStyle.Regular)
            lblAuthorizedSummary.ForeColor = ThemePalette.TextPrimary
            lblAuthorizedSummary.TextAlign = ContentAlignment.TopLeft

            layout.Controls.Add(supportLabel, 0, 0)
            layout.Controls.Add(centerPanel, 1, 0)
            layout.Controls.Add(lblAuthorizedSummary, 2, 0)
            panel.Controls.Add(layout)
            Return panel
        End Function

        Private Sub ConfigureControls()
            txtUsername.Text = "admin"
            txtPassword.UseSystemPasswordChar = True

            lstStartupActions.Items.Clear()
            For Each actionItem As StartupActionItem In _startupActions
                lstStartupActions.Items.Add(actionItem)
            Next
            If lstStartupActions.Items.Count > 0 Then
                lstStartupActions.SelectedIndex = 0
            End If

            _clockTimer.Interval = 1000
            UpdateClock()
        End Sub

        Private Sub WireEvents()
            AddHandler Load, AddressOf FrmLogin_Load
            AddHandler lstCompanies.SelectedIndexChanged, AddressOf lstCompanies_SelectedIndexChanged
            AddHandler lstStartupActions.SelectedIndexChanged, AddressOf lstStartupActions_SelectedIndexChanged
            AddHandler lstStartupActions.DoubleClick, AddressOf lstStartupActions_DoubleClick
            AddHandler btnRunAction.Click, AddressOf btnRunAction_Click
            AddHandler btnRefreshCompanies.Click, AddressOf btnRefreshCompanies_Click
            AddHandler btnLogin.Click, AddressOf btnLogin_Click
            AddHandler btnDemoAdmin.Click, AddressOf btnDemoAdmin_Click
            AddHandler btnDemoStaff.Click, AddressOf btnDemoStaff_Click
            AddHandler btnExit.Click, AddressOf btnExit_Click
            AddHandler chkShowPassword.CheckedChanged, AddressOf chkShowPassword_CheckedChanged
            AddHandler _clockTimer.Tick, AddressOf ClockTimer_Tick
        End Sub

        Private Async Sub FrmLogin_Load(sender As Object, e As EventArgs)
            _clockTimer.Start()
            WorkspaceHost_Resize(_workspaceHost, EventArgs.Empty)
            Await LoadStartupContextAsync()
        End Sub

        Private Async Function LoadStartupContextAsync(Optional preferredCompanyId As String = Nothing) As Task
            ToggleUi(False, "Loading startup workspace...")

            Try
                _startupProfile = Await Task.Run(Function() _companyWorkspaceService.LoadStartupProfile())
                _companies = (Await Task.Run(Function() _companyWorkspaceService.LoadCompanies())).ToList()

                If preferredCompanyId IsNot Nothing Then
                    _startupProfile.SelectedCompanyId = preferredCompanyId
                End If

                BindCompanies()
                ApplyStartupProfileVisuals()
                LoadSavedCredentialsForSelection()
                SetStatus("Startup workspace ready. Choose a company, run startup tools, or sign in.", ThemePalette.AccentGreen)
            Catch ex As Exception
                AppLogger.Error("Startup shell initialization failed.", ex)
                _companies.Clear()
                lstCompanies.Items.Clear()
                SetStatus("Startup shell could not be initialized.", ThemePalette.DangerRed)
            Finally
                ToggleUi(True)
            End Try
        End Function

        Private Sub BindCompanies()
            lstCompanies.BeginUpdate()
            lstCompanies.Items.Clear()
            For Each company As CompanyWorkspaceRecord In _companies
                lstCompanies.Items.Add(company)
            Next
            lstCompanies.EndUpdate()

            Dim selectedCompany As CompanyWorkspaceRecord =
                _companies.FirstOrDefault(Function(company) String.Equals(company.WorkspaceId, _startupProfile.SelectedCompanyId, StringComparison.OrdinalIgnoreCase))

            If selectedCompany Is Nothing AndAlso _companies.Count > 0 Then
                selectedCompany = _companies(0)
            End If

            If selectedCompany IsNot Nothing Then
                lstCompanies.SelectedItem = selectedCompany
            End If
        End Sub

        Private Sub ApplyStartupProfileVisuals()
            Dim editionLabel As String = If(_startupProfile.EditionLabel, "Desktop Edition").Trim()
            If editionLabel = String.Empty Then
                editionLabel = "Desktop Edition"
            End If

            lblTopTitle.Text = $"Healthwond ERP | {editionLabel} | Licensed Workstation: Sam"
            lblEditionBadge.Text = "Healthwond" & Environment.NewLine & "ERP Workspace" & Environment.NewLine & editionLabel
            lblAuthorizedSummary.Text = "Authorised Access :" & Environment.NewLine &
                                        GetSelectedCompanyName() & Environment.NewLine & Environment.NewLine &
                                        $"ERP Version : {editionLabel}"
        End Sub

        Private Sub LoadSavedCredentialsForSelection()
            Dim selectedCompany As CompanyWorkspaceRecord = GetSelectedCompanyRecord()
            lblSelectedCompany.Text = $"Company : {GetSelectedCompanyName()}"
            lblSelectedCompanyValue.Text = GetSelectedCompanyName()

            If selectedCompany Is Nothing Then
                txtUsername.Text = "admin"
                txtPassword.Clear()
                Return
            End If

            Dim savedCredentials As SavedCredentialRecord = _savedCredentialService.Load(selectedCompany.WorkspaceId)
            If savedCredentials Is Nothing Then
                txtUsername.Text = "admin"
                txtPassword.Clear()
                Return
            End If

            txtUsername.Text = savedCredentials.Username
            txtPassword.Text = savedCredentials.Password
        End Sub

        Private Sub WorkspaceHost_Resize(sender As Object, e As EventArgs)
            If _workspaceCard Is Nothing OrElse _workspaceShadow Is Nothing Then
                Return
            End If

            Dim cardWidth As Integer = Math.Min(Math.Max(_workspaceHost.ClientSize.Width - 72, 1260), 1500)
            Dim cardHeight As Integer = Math.Min(Math.Max(_workspaceHost.ClientSize.Height - 56, 612), 700)
            _workspaceCard.Size = New Size(cardWidth, cardHeight)
            _workspaceShadow.Size = _workspaceCard.Size

            Dim left As Integer = Math.Max((_workspaceHost.ClientSize.Width - _workspaceCard.Width) \ 2 - 8, 12)
            Dim top As Integer = Math.Max((_workspaceHost.ClientSize.Height - _workspaceCard.Height) \ 2 - 8, 12)

            _workspaceCard.Location = New Point(left, top)
            _workspaceShadow.Location = New Point(left + 18, top + 18)
            _workspaceShadow.SendToBack()
            _workspaceCard.BringToFront()
        End Sub

        Private Sub ConfigureClassicActionButton(button As Button, text As String, isPrimary As Boolean, width As Integer)
            button.Text = text
            button.Width = width
            button.Height = 38
            button.Cursor = Cursors.Hand
            button.FlatStyle = FlatStyle.Flat
            button.Font = New Font("Consolas", 10.5F, FontStyle.Bold)

            If isPrimary Then
                button.FlatAppearance.BorderSize = 1
                button.FlatAppearance.BorderColor = ThemePalette.ClassicShellBorder
                button.BackColor = ThemePalette.ClassicShellDarkGreen
                button.ForeColor = Color.White
            Else
                button.FlatAppearance.BorderSize = 1
                button.FlatAppearance.BorderColor = ThemePalette.ClassicShellBorder
                button.BackColor = ThemePalette.ClassicShellCream
                button.ForeColor = ThemePalette.ClassicShellBorder
            End If
        End Sub

        Private Sub lstCompanies_SelectedIndexChanged(sender As Object, e As EventArgs)
            Dim selectedCompany As CompanyWorkspaceRecord = GetSelectedCompanyRecord()
            If selectedCompany Is Nothing Then
                Return
            End If

            _companyWorkspaceService.SaveSelectedCompany(selectedCompany.WorkspaceId)
            _startupProfile.SelectedCompanyId = selectedCompany.WorkspaceId
            ApplyStartupProfileVisuals()
            LoadSavedCredentialsForSelection()
            SetStatus($"Selected company workspace: {selectedCompany.DisplayName}.", ThemePalette.BrandBlue)
        End Sub

        Private Sub lstStartupActions_SelectedIndexChanged(sender As Object, e As EventArgs)
            Dim selectedAction As StartupActionItem = TryCast(lstStartupActions.SelectedItem, StartupActionItem)
            If selectedAction Is Nothing Then
                lblActionDescription.Text = String.Empty
                Return
            End If

            lblActionDescription.Text = selectedAction.Description
        End Sub

        Private Async Sub lstStartupActions_DoubleClick(sender As Object, e As EventArgs)
            Await ExecuteSelectedStartupActionAsync()
        End Sub

        Private Async Sub btnRunAction_Click(sender As Object, e As EventArgs)
            Await ExecuteSelectedStartupActionAsync()
        End Sub

        Private Async Sub btnRefreshCompanies_Click(sender As Object, e As EventArgs)
            Await LoadStartupContextAsync(GetSelectedCompanyId())
        End Sub

        Private Sub btnDemoAdmin_Click(sender As Object, e As EventArgs)
            txtUsername.Text = "admin"
            txtPassword.Text = "Admin@123"
            SetStatus("Demo administrator credentials loaded. Press Sign In.", ThemePalette.BrandBlue)
            txtPassword.Focus()
            txtPassword.SelectionStart = txtPassword.TextLength
        End Sub

        Private Sub btnDemoStaff_Click(sender As Object, e As EventArgs)
            txtUsername.Text = "staff"
            txtPassword.Text = "Staff@123"
            SetStatus("Demo staff credentials loaded. Press Sign In.", ThemePalette.BrandBlue)
            txtPassword.Focus()
            txtPassword.SelectionStart = txtPassword.TextLength
        End Sub

        Private Async Sub btnLogin_Click(sender As Object, e As EventArgs)
            Await LoginAsync()
        End Sub

        Private Sub btnExit_Click(sender As Object, e As EventArgs)
            Close()
        End Sub

        Private Sub chkShowPassword_CheckedChanged(sender As Object, e As EventArgs)
            txtPassword.UseSystemPasswordChar = Not chkShowPassword.Checked
        End Sub

        Private Async Function ExecuteSelectedStartupActionAsync() As Task
            Dim selectedAction As StartupActionItem = TryCast(lstStartupActions.SelectedItem, StartupActionItem)
            If selectedAction Is Nothing Then
                SetStatus("Choose a startup action first.", ThemePalette.WarningAmber)
                Return
            End If

            Select Case selectedAction.Key
                Case ActionCreateCompany
                    Await CreateNewCompanyAsync()
                Case ActionRestoreDefaultDemo
                    Await RestoreDefaultDemoAsync()
                Case ActionRestoreFromBackup
                    Await RestoreFromBackupAsync()
                Case ActionDeleteSavedPassword
                    DeleteSavedPassword()
                Case ActionChangeOperatorPowers
                    Await OpenOperatorPowersAsync()
                Case ActionChangeErpVersion
                    ChangeErpVersion()
            End Select
        End Function

        Private Async Function CreateNewCompanyAsync() As Task
            Using prompt As New FrmTextPrompt("Create New Company", "Enter the company name for the new workspace.")
                If prompt.ShowDialog(Me) <> DialogResult.OK Then
                    Return
                End If

                ToggleUi(False, "Creating company workspace...")

                Try
                    Dim result As CompanyWorkspaceOperationResult =
                        Await Task.Run(Function() _companyWorkspaceService.CreateCompany(prompt.PromptValue))

                    SetStatus(result.Message, If(result.IsSuccess, ThemePalette.AccentGreen, ThemePalette.DangerRed))
                    If result.IsSuccess Then
                        Await LoadStartupContextAsync(result.Company.WorkspaceId)
                    End If
                Catch ex As Exception
                    AppLogger.Error("Create company startup action failed.", ex)
                    SetStatus("The company workspace could not be created.", ThemePalette.DangerRed)
                Finally
                    ToggleUi(True)
                End Try
            End Using
        End Function

        Private Async Function RestoreDefaultDemoAsync() As Task
            ToggleUi(False, "Restoring default demonstration...")

            Try
                Dim result As CompanyWorkspaceOperationResult =
                    Await Task.Run(Function() _companyWorkspaceService.RestoreDefaultDemonstration())

                SetStatus(result.Message, If(result.IsSuccess, ThemePalette.AccentGreen, ThemePalette.DangerRed))
                If result.IsSuccess Then
                    Await LoadStartupContextAsync(result.Company.WorkspaceId)
                End If
            Catch ex As Exception
                AppLogger.Error("Restore default demonstration action failed.", ex)
                SetStatus("The default demonstration could not be restored.", ThemePalette.DangerRed)
            Finally
                ToggleUi(True)
            End Try
        End Function

        Private Async Function RestoreFromBackupAsync() As Task
            Using dialog As New OpenFileDialog()
                dialog.Title = "Import Demonstration Database"
                dialog.Filter = "SQLite Database (*.db)|*.db|All Files (*.*)|*.*"
                dialog.InitialDirectory = AppPaths.DataRootDirectory

                If dialog.ShowDialog(Me) <> DialogResult.OK Then
                    Return
                End If

                Dim defaultDisplayName As String = IO.Path.GetFileNameWithoutExtension(dialog.FileName)
                Using prompt As New FrmTextPrompt("Workspace Name", "Enter the company name to use for the imported demonstration.", defaultDisplayName)
                    If prompt.ShowDialog(Me) <> DialogResult.OK Then
                        Return
                    End If

                    ToggleUi(False, "Importing demonstration workspace...")

                    Try
                        Dim result As CompanyWorkspaceOperationResult =
                            Await Task.Run(Function() _companyWorkspaceService.RestoreFromBackup(dialog.FileName, prompt.PromptValue))

                        SetStatus(result.Message, If(result.IsSuccess, ThemePalette.AccentGreen, ThemePalette.DangerRed))
                        If result.IsSuccess Then
                            Await LoadStartupContextAsync(result.Company.WorkspaceId)
                        End If
                    Catch ex As Exception
                        AppLogger.Error("Restore from demonstration action failed.", ex)
                        SetStatus("The selected demonstration could not be imported.", ThemePalette.DangerRed)
                    Finally
                        ToggleUi(True)
                    End Try
                End Using
            End Using
        End Function

        Private Sub DeleteSavedPassword()
            Dim selectedCompany As CompanyWorkspaceRecord = GetSelectedCompanyRecord()
            If selectedCompany Is Nothing Then
                SetStatus("Select a company before clearing saved credentials.", ThemePalette.WarningAmber)
                Return
            End If

            Dim removed As Boolean = _savedCredentialService.Clear(selectedCompany.WorkspaceId)
            txtPassword.Clear()

            If removed Then
                SetStatus("Saved credentials were removed for the selected company.", ThemePalette.AccentGreen)
            Else
                SetStatus("No saved credentials were found for the selected company.", ThemePalette.WarningAmber)
            End If
        End Sub

        Private Async Function OpenOperatorPowersAsync() As Task
            Dim selectedCompany As CompanyWorkspaceRecord = GetSelectedCompanyRecord()
            If selectedCompany Is Nothing Then
                SetStatus("Select a company before changing operator powers.", ThemePalette.WarningAmber)
                Return
            End If

            ToggleUi(False, "Opening operator powers...")

            Try
                Dim runtime As AppRuntimeContext = Await Task.Run(Function() _runtimeFactory.CreateRuntime(selectedCompany))
                Using operatorForm As New FrmOperatorPowers(runtime.UserAdministrationService, selectedCompany.DisplayName)
                    operatorForm.ShowDialog(Me)
                End Using

                SetStatus("Operator powers window closed.", ThemePalette.BrandBlue)
            Catch ex As Exception
                AppLogger.Error("Operator powers startup action failed.", ex)
                SetStatus("Operator powers could not be opened.", ThemePalette.DangerRed)
            Finally
                ToggleUi(True)
            End Try
        End Function

        Private Sub ChangeErpVersion()
            Using prompt As New FrmTextPrompt("ERP Version", "Enter the startup shell ERP version or edition label.", _startupProfile.EditionLabel)
                If prompt.ShowDialog(Me) <> DialogResult.OK Then
                    Return
                End If

                Dim result As EntityOperationResult = _companyWorkspaceService.UpdateEditionLabel(prompt.PromptValue)
                SetStatus(result.Message, If(result.IsSuccess, ThemePalette.AccentGreen, ThemePalette.DangerRed))

                If result.IsSuccess Then
                    _startupProfile.EditionLabel = prompt.PromptValue
                    ApplyStartupProfileVisuals()
                End If
            End Using
        End Sub

        Private Async Function LoginAsync() As Task
            Dim selectedCompany As CompanyWorkspaceRecord = GetSelectedCompanyRecord()
            If selectedCompany Is Nothing Then
                SetStatus("Select a company before signing in.", ThemePalette.WarningAmber)
                Return
            End If

            ToggleUi(False, "Opening company workspace...")

            Try
                Dim runtime As AppRuntimeContext = Await Task.Run(Function() _runtimeFactory.CreateRuntime(selectedCompany))
                Dim result As LoginResult = Await runtime.AuthService.AuthenticateAsync(txtUsername.Text, txtPassword.Text)

                If Not result.IsSuccess Then
                    SetStatus(result.Message, ThemePalette.DangerRed)
                    Return
                End If

                _savedCredentialService.Save(selectedCompany.WorkspaceId, txtUsername.Text.Trim(), txtPassword.Text)
                _companyWorkspaceService.SaveSelectedCompany(selectedCompany.WorkspaceId)
                SessionManager.StartSession(result.User)

                Using dashboard As New FrmDashboard(
                    runtime.DashboardService,
                    runtime.ProductService,
                    runtime.CustomerService,
                    runtime.SupplierService,
                    runtime.BillingService,
                    runtime.PurchaseService,
                    runtime.PurchasePrintService,
                    runtime.InvoiceExportService,
                    runtime.ReportService,
                    runtime.InventoryService,
                    runtime.StockOperationService,
                    runtime.SettlementService,
                    runtime.SettingsService,
                    runtime.MaintenanceService,
                    runtime.AccountingService)

                    Hide()
                    dashboard.ShowDialog(Me)

                    If dashboard.RequestedLogout Then
                        SessionManager.Clear()
                        txtPassword.Clear()
                        LoadSavedCredentialsForSelection()
                        Show()
                        Activate()
                        txtPassword.Focus()
                        SetStatus("You have been signed out.", ThemePalette.AccentGreen)
                    Else
                        Close()
                    End If
                End Using
            Catch ex As Exception
                AppLogger.Error("Login workflow failed.", ex)
                SetStatus("Unable to complete the sign-in request.", ThemePalette.DangerRed)
            Finally
                ToggleUi(True)
            End Try
        End Function

        Private Sub ToggleUi(isEnabled As Boolean, Optional busyMessage As String = Nothing)
            lstCompanies.Enabled = isEnabled
            lstStartupActions.Enabled = isEnabled
            btnRunAction.Enabled = isEnabled
            btnRefreshCompanies.Enabled = isEnabled
            txtUsername.Enabled = isEnabled
            txtPassword.Enabled = isEnabled
            chkShowPassword.Enabled = isEnabled
            btnLogin.Enabled = isEnabled
            btnDemoAdmin.Enabled = isEnabled
            btnDemoStaff.Enabled = isEnabled
            btnExit.Enabled = isEnabled
            btnLogin.Text = If(isEnabled, "Sign In", "Signing In...")
            btnRunAction.Text = If(isEnabled, "Run Selected", "Working...")

            If Not isEnabled AndAlso busyMessage IsNot Nothing Then
                SetStatus(busyMessage, ThemePalette.BrandBlue)
            End If
        End Sub

        Private Sub ClockTimer_Tick(sender As Object, e As EventArgs)
            UpdateClock()
        End Sub

        Private Sub UpdateClock()
            Dim nowValue As DateTime = DateTime.Now
            lblTopClock.Text = nowValue.ToString("HH:mm:ss")
            lblDateValue.Text = nowValue.ToString("dd MMM, yyyy")
            lblDayValue.Text = nowValue.ToString("dddd")
            lblTimeValue.Text = nowValue.ToString("hh:mm:ss tt")
        End Sub

        Private Function GetSelectedCompanyRecord() As CompanyWorkspaceRecord
            Return TryCast(lstCompanies.SelectedItem, CompanyWorkspaceRecord)
        End Function

        Private Function GetSelectedCompanyId() As String
            Dim company As CompanyWorkspaceRecord = GetSelectedCompanyRecord()
            If company Is Nothing Then
                Return String.Empty
            End If

            Return company.WorkspaceId
        End Function

        Private Function GetSelectedCompanyName() As String
            Dim company As CompanyWorkspaceRecord = GetSelectedCompanyRecord()
            If company Is Nothing Then
                Return "No company selected"
            End If

            Return company.DisplayName
        End Function

        Private Sub SetStatus(message As String, colorValue As Color)
            lblStatus.Text = message
            lblStatus.ForeColor = colorValue
        End Sub

        Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
            Select Case keyData
                Case Keys.F2
                    RunActionByKey(ActionCreateCompany)
                    Return True
                Case Keys.F3
                    RunActionByKey(ActionRestoreDefaultDemo)
                    Return True
                Case Keys.F4
                    RunActionByKey(ActionRestoreFromBackup)
                    Return True
                Case Keys.F5
                    btnLogin.PerformClick()
                    Return True
                Case Keys.F6
                    RunActionByKey(ActionDeleteSavedPassword)
                    Return True
                Case Keys.F7
                    RunActionByKey(ActionChangeOperatorPowers)
                    Return True
                Case Keys.F8
                    chkShowPassword.Checked = Not chkShowPassword.Checked
                    Return True
                Case Keys.F9
                    RunActionByKey(ActionChangeErpVersion)
                    Return True
                Case Keys.Escape
                    Close()
                    Return True
            End Select

            Return MyBase.ProcessCmdKey(msg, keyData)
        End Function

        Private Sub RunActionByKey(actionKey As String)
            For index As Integer = 0 To lstStartupActions.Items.Count - 1
                Dim actionItem As StartupActionItem = TryCast(lstStartupActions.Items(index), StartupActionItem)
                If actionItem IsNot Nothing AndAlso String.Equals(actionItem.Key, actionKey, StringComparison.OrdinalIgnoreCase) Then
                    lstStartupActions.SelectedIndex = index
                    btnRunAction.PerformClick()
                    Exit For
                End If
            Next
        End Sub

        Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
            _clockTimer.Stop()
            MyBase.OnFormClosing(e)
        End Sub

        Private Class StartupActionItem

            Public Sub New(key As String, title As String, description As String)
                Me.Key = key
                Me.Title = title
                Me.Description = description
            End Sub

            Public ReadOnly Property Key As String
            Public ReadOnly Property Title As String
            Public ReadOnly Property Description As String

            Public Overrides Function ToString() As String
                Return Title
            End Function

        End Class

    End Class

End Namespace
