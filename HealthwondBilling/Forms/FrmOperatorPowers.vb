Imports HealthwondBilling.Models
Imports HealthwondBilling.Services
Imports HealthwondBilling.Utilities
Imports System.Globalization
Imports System.Linq

Namespace Forms

    Public Class FrmOperatorPowers
        Inherits Form

        Private ReadOnly _userAdministrationService As UserAdministrationService
        Private ReadOnly grdUsers As New DataGridView()
        Private ReadOnly txtUsername As New TextBox()
        Private ReadOnly txtFullName As New TextBox()
        Private ReadOnly cmbRole As New ComboBox()
        Private ReadOnly chkIsActive As New CheckBox()
        Private ReadOnly btnNew As New Button()
        Private ReadOnly btnSave As New Button()
        Private ReadOnly btnResetPassword As New Button()
        Private ReadOnly btnReload As New Button()
        Private ReadOnly btnClose As New Button()
        Private ReadOnly lblStatus As New Label()
        Private ReadOnly lblCompanyName As New Label()

        Private _users As New List(Of UserAccount)()
        Private _editingUserId As Integer
        Private _isLoading As Boolean

        Public Sub New(userAdministrationService As UserAdministrationService, companyName As String)
            _userAdministrationService = userAdministrationService

            Text = "Healthwond Billing System - Operator Powers"
            StartPosition = FormStartPosition.CenterParent
            Size = New Size(1100, 760)
            MinimumSize = New Size(980, 680)
            BackColor = ThemePalette.AppBackground
            Font = New Font("Segoe UI", 10.0F, FontStyle.Regular)
            KeyPreview = True

            lblCompanyName.Text = companyName

            BuildLayout()
            ConfigureGrid()
            ConfigureEditor()
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
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 98))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 70))
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 34))

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
                .Height = 38,
                .Text = "Operator Powers",
                .Font = New Font("Segoe UI Semibold", 24.0F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary
            }

            Dim subtitle As New Label With {
                .Dock = DockStyle.Top,
                .Height = 26,
                .Text = "Maintain usernames, operator roles, active status, and password resets for the selected company workspace.",
                .Font = New Font("Segoe UI", 10.5F, FontStyle.Regular),
                .ForeColor = ThemePalette.TextMuted
            }

            lblCompanyName.Dock = DockStyle.Top
            lblCompanyName.Height = 26
            lblCompanyName.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
            lblCompanyName.ForeColor = ThemePalette.AccentGreen

            panel.Controls.Add(lblCompanyName)
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

            ConfigureToolbarButton(btnNew, "New Operator", True, 126)
            ConfigureToolbarButton(btnSave, "Save Changes", True, 126)
            ConfigureToolbarButton(btnResetPassword, "Reset Password", False, 132)
            ConfigureToolbarButton(btnReload, "Reload", False, 96)
            ConfigureToolbarButton(btnClose, "Close", False, 96)

            flow.Controls.AddRange(New Control() {btnNew, btnSave, btnResetPassword, btnReload, btnClose})
            panel.Controls.Add(flow)
            Return panel
        End Function

        Private Function BuildMainPanel() As Control
            Dim split As New SplitContainer With {
                .Dock = DockStyle.Fill,
                .SplitterDistance = 620,
                .BackColor = ThemePalette.AppBackground
            }

            split.Panel1.Controls.Add(BuildUserListCard())
            split.Panel2.Controls.Add(BuildEditorCard())
            Return split
        End Function

        Private Function BuildUserListCard() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)

            Dim title As New Label With {
                .Dock = DockStyle.Top,
                .Height = 32,
                .Text = "Current operators",
                .Font = New Font("Segoe UI Semibold", 14.0F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary
            }

            grdUsers.Dock = DockStyle.Fill
            panel.Controls.Add(grdUsers)
            panel.Controls.Add(title)
            Return panel
        End Function

        Private Function BuildEditorCard() As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.White}
            UiStyler.StyleCard(panel)

            Dim title As New Label With {
                .Dock = DockStyle.Top,
                .Height = 32,
                .Text = "Operator details",
                .Font = New Font("Segoe UI Semibold", 14.0F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary
            }

            Dim editor As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .Padding = New Padding(0, 16, 0, 0)
            }
            editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 74))
            editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 74))
            editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 74))
            editor.RowStyles.Add(New RowStyle(SizeType.Absolute, 44))
            editor.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            editor.Controls.Add(CreateInputHost("Username", txtUsername), 0, 0)
            editor.Controls.Add(CreateInputHost("Full Name", txtFullName), 0, 1)

            Dim roleHost As New Panel With {.Dock = DockStyle.Fill}
            Dim roleCaption As New Label With {
                .Dock = DockStyle.Top,
                .Height = 24,
                .Text = "Role",
                .Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary
            }
            cmbRole.Dock = DockStyle.Top
            cmbRole.Height = 36
            roleHost.Controls.Add(cmbRole)
            roleHost.Controls.Add(roleCaption)
            editor.Controls.Add(roleHost, 0, 2)

            chkIsActive.Dock = DockStyle.Top
            chkIsActive.Text = "Operator account is active"
            chkIsActive.Font = New Font("Segoe UI", 10.0F, FontStyle.Regular)
            chkIsActive.ForeColor = ThemePalette.TextPrimary
            editor.Controls.Add(chkIsActive, 0, 3)

            Dim note As New Label With {
                .Dock = DockStyle.Fill,
                .ForeColor = ThemePalette.TextMuted,
                .Text = "Use Save Changes to add a new operator or update the selected operator. Password changes are handled separately through Reset Password.",
                .Font = New Font("Segoe UI", 9.75F, FontStyle.Italic),
                .TextAlign = ContentAlignment.TopLeft
            }
            editor.Controls.Add(note, 0, 4)

            panel.Controls.Add(editor)
            panel.Controls.Add(title)
            Return panel
        End Function

        Private Function CreateInputHost(labelText As String, editorControl As Control) As Control
            Dim panel As New Panel With {.Dock = DockStyle.Fill}

            Dim caption As New Label With {
                .Dock = DockStyle.Top,
                .Height = 24,
                .Text = labelText,
                .Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold),
                .ForeColor = ThemePalette.TextPrimary
            }

            editorControl.Dock = DockStyle.Top
            editorControl.Height = 36
            panel.Controls.Add(editorControl)
            panel.Controls.Add(caption)
            Return panel
        End Function

        Private Sub ConfigureToolbarButton(button As Button, text As String, isPrimary As Boolean, width As Integer)
            button.Text = text
            button.Width = width
            button.Height = 40
            button.FlatStyle = FlatStyle.Flat
            button.Cursor = Cursors.Hand
            button.Font = New Font("Segoe UI Semibold", 9.5F, FontStyle.Bold)

            If isPrimary Then
                button.FlatAppearance.BorderSize = 0
                button.BackColor = ThemePalette.BrandBlue
                button.ForeColor = Color.White
            Else
                button.FlatAppearance.BorderColor = ThemePalette.TextMuted
                button.FlatAppearance.BorderSize = 1
                button.BackColor = Color.White
                button.ForeColor = ThemePalette.TextPrimary
            End If
        End Sub

        Private Sub ConfigureGrid()
            grdUsers.AllowUserToAddRows = False
            grdUsers.AllowUserToDeleteRows = False
            grdUsers.AllowUserToResizeRows = False
            grdUsers.AutoGenerateColumns = False
            grdUsers.BackgroundColor = Color.White
            grdUsers.BorderStyle = BorderStyle.None
            grdUsers.MultiSelect = False
            grdUsers.ReadOnly = True
            grdUsers.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            grdUsers.RowHeadersVisible = False
            grdUsers.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            grdUsers.ColumnHeadersDefaultCellStyle.BackColor = ThemePalette.BrandBlue
            grdUsers.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
            grdUsers.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold)
            grdUsers.EnableHeadersVisualStyles = False

            grdUsers.Columns.Add(New DataGridViewTextBoxColumn With {
                .Name = "Id",
                .DataPropertyName = "Id",
                .Visible = False
            })
            grdUsers.Columns.Add(New DataGridViewTextBoxColumn With {
                .HeaderText = "Username",
                .DataPropertyName = "Username",
                .FillWeight = 23.0F
            })
            grdUsers.Columns.Add(New DataGridViewTextBoxColumn With {
                .HeaderText = "Full Name",
                .DataPropertyName = "FullName",
                .FillWeight = 28.0F
            })
            grdUsers.Columns.Add(New DataGridViewTextBoxColumn With {
                .HeaderText = "Role",
                .DataPropertyName = "RoleText",
                .FillWeight = 16.0F
            })
            grdUsers.Columns.Add(New DataGridViewCheckBoxColumn With {
                .HeaderText = "Active",
                .DataPropertyName = "IsActive",
                .FillWeight = 12.0F
            })
            grdUsers.Columns.Add(New DataGridViewTextBoxColumn With {
                .HeaderText = "Last Login",
                .DataPropertyName = "LastLoginText",
                .FillWeight = 21.0F
            })
        End Sub

        Private Sub ConfigureEditor()
            txtUsername.BorderStyle = BorderStyle.FixedSingle
            txtFullName.BorderStyle = BorderStyle.FixedSingle
            cmbRole.DropDownStyle = ComboBoxStyle.DropDownList
            cmbRole.Items.Add(UserRole.Admin)
            cmbRole.Items.Add(UserRole.Staff)
            cmbRole.SelectedIndex = 0
            chkIsActive.Checked = True
        End Sub

        Private Sub WireEvents()
            AddHandler Load, AddressOf FrmOperatorPowers_Load
            AddHandler btnNew.Click, AddressOf btnNew_Click
            AddHandler btnSave.Click, AddressOf btnSave_Click
            AddHandler btnResetPassword.Click, AddressOf btnResetPassword_Click
            AddHandler btnReload.Click, AddressOf btnReload_Click
            AddHandler btnClose.Click, AddressOf btnClose_Click
            AddHandler grdUsers.SelectionChanged, AddressOf grdUsers_SelectionChanged
        End Sub

        Private Async Sub FrmOperatorPowers_Load(sender As Object, e As EventArgs)
            Await LoadUsersAsync(0)
            BeginNewUser()
        End Sub

        Private Async Function LoadUsersAsync(selectUserId As Integer) As Task
            _isLoading = True
            Try
                _users = (Await _userAdministrationService.LoadUsersAsync()).OrderBy(Function(user) user.Username).ToList()

                Dim rows As List(Of OperatorGridRow) =
                    _users.Select(
                        Function(user) New OperatorGridRow With {
                            .Id = user.Id,
                            .Username = user.Username,
                            .FullName = user.FullName,
                            .RoleText = user.Role.ToFriendlyText(),
                            .IsActive = user.IsActive,
                            .LastLoginText = If(user.LastLoginAt.HasValue, user.LastLoginAt.Value.ToString("dd-MMM-yyyy HH:mm", CultureInfo.InvariantCulture), "Never")
                        }).ToList()

                grdUsers.DataSource = rows

                If selectUserId > 0 Then
                    For Each row As DataGridViewRow In grdUsers.Rows
                        If Convert.ToInt32(row.Cells("Id").Value, CultureInfo.InvariantCulture) = selectUserId Then
                            row.Selected = True
                            grdUsers.CurrentCell = row.Cells("Username")
                            PopulateEditorFromUserId(selectUserId)
                            Exit For
                        End If
                    Next
                ElseIf grdUsers.Rows.Count > 0 Then
                    grdUsers.Rows(0).Selected = True
                    grdUsers.CurrentCell = grdUsers.Rows(0).Cells("Username")
                    PopulateEditorFromUserId(Convert.ToInt32(grdUsers.Rows(0).Cells("Id").Value, CultureInfo.InvariantCulture))
                Else
                    BeginNewUser()
                End If

                SetStatus("Operator powers loaded.", ThemePalette.AccentGreen)
            Catch ex As Exception
                AppLogger.Error("Operator powers could not be loaded.", ex)
                SetStatus("Operator records could not be loaded.", ThemePalette.DangerRed)
            Finally
                _isLoading = False
            End Try
        End Function

        Private Sub PopulateEditorFromUserId(userId As Integer)
            Dim selectedUser As UserAccount = _users.FirstOrDefault(Function(user) user.Id = userId)
            If selectedUser Is Nothing Then
                Return
            End If

            _editingUserId = selectedUser.Id
            txtUsername.Text = selectedUser.Username
            txtFullName.Text = selectedUser.FullName
            cmbRole.SelectedItem = selectedUser.Role
            chkIsActive.Checked = selectedUser.IsActive
        End Sub

        Private Sub BeginNewUser()
            _editingUserId = 0
            txtUsername.Clear()
            txtFullName.Clear()
            cmbRole.SelectedItem = UserRole.Staff
            chkIsActive.Checked = True
            txtUsername.Focus()
            SetStatus("Entering a new operator record.", ThemePalette.BrandBlue)
        End Sub

        Private Async Sub btnSave_Click(sender As Object, e As EventArgs)
            Dim user As New UserAccount With {
                .Id = _editingUserId,
                .Username = txtUsername.Text.Trim(),
                .FullName = txtFullName.Text.Trim(),
                .Role = CType(cmbRole.SelectedItem, UserRole),
                .IsActive = chkIsActive.Checked
            }

            Dim initialPassword As String = Nothing
            If user.Id <= 0 Then
                Using prompt As New FrmTextPrompt("Initial Password", "Enter the initial password for the new operator.", "", True)
                    If prompt.ShowDialog(Me) <> DialogResult.OK Then
                        Return
                    End If

                    initialPassword = prompt.PromptValue
                End Using
            End If

            ToggleUi(False)
            Try
                Dim result As EntityOperationResult = Await _userAdministrationService.SaveUserAsync(user, initialPassword)
                SetStatus(result.Message, If(result.IsSuccess, ThemePalette.AccentGreen, ThemePalette.DangerRed))

                If result.IsSuccess Then
                    Await LoadUsersAsync(result.EntityId)
                End If
            Catch ex As Exception
                AppLogger.Error("Operator powers save failed.", ex)
                SetStatus("Operator powers could not be saved.", ThemePalette.DangerRed)
            Finally
                ToggleUi(True)
            End Try
        End Sub

        Private Async Sub btnResetPassword_Click(sender As Object, e As EventArgs)
            If _editingUserId <= 0 Then
                SetStatus("Select an existing operator before resetting the password.", ThemePalette.WarningAmber)
                Return
            End If

            Using prompt As New FrmTextPrompt("Reset Password", "Enter a replacement password for the selected operator.", "", True)
                If prompt.ShowDialog(Me) <> DialogResult.OK Then
                    Return
                End If

                ToggleUi(False)
                Try
                    Dim result As EntityOperationResult = Await _userAdministrationService.ResetPasswordAsync(_editingUserId, prompt.PromptValue)
                    SetStatus(result.Message, If(result.IsSuccess, ThemePalette.AccentGreen, ThemePalette.DangerRed))
                Catch ex As Exception
                    AppLogger.Error("Operator password reset failed.", ex)
                    SetStatus("Operator password could not be reset.", ThemePalette.DangerRed)
                Finally
                    ToggleUi(True)
                End Try
            End Using
        End Sub

        Private Async Sub btnReload_Click(sender As Object, e As EventArgs)
            Await LoadUsersAsync(_editingUserId)
        End Sub

        Private Sub btnNew_Click(sender As Object, e As EventArgs)
            grdUsers.ClearSelection()
            BeginNewUser()
        End Sub

        Private Sub btnClose_Click(sender As Object, e As EventArgs)
            Close()
        End Sub

        Private Sub grdUsers_SelectionChanged(sender As Object, e As EventArgs)
            If _isLoading OrElse grdUsers.SelectedRows.Count = 0 Then
                Return
            End If

            Dim userId As Integer = Convert.ToInt32(grdUsers.SelectedRows(0).Cells("Id").Value, CultureInfo.InvariantCulture)
            PopulateEditorFromUserId(userId)
        End Sub

        Private Sub ToggleUi(isEnabled As Boolean)
            grdUsers.Enabled = isEnabled
            txtUsername.Enabled = isEnabled
            txtFullName.Enabled = isEnabled
            cmbRole.Enabled = isEnabled
            chkIsActive.Enabled = isEnabled
            btnNew.Enabled = isEnabled
            btnSave.Enabled = isEnabled
            btnResetPassword.Enabled = isEnabled
            btnReload.Enabled = isEnabled
            btnClose.Enabled = isEnabled
        End Sub

        Private Sub SetStatus(message As String, colorValue As Color)
            lblStatus.Text = message
            lblStatus.ForeColor = colorValue
        End Sub

        Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
            If keyData = Keys.Escape Then
                Close()
                Return True
            End If

            If keyData = Keys.F2 Then
                BeginNewUser()
                Return True
            End If

            If keyData = Keys.F5 Then
                btnSave.PerformClick()
                Return True
            End If

            Return MyBase.ProcessCmdKey(msg, keyData)
        End Function

        Private Class OperatorGridRow

            Public Property Id As Integer
            Public Property Username As String = String.Empty
            Public Property FullName As String = String.Empty
            Public Property RoleText As String = String.Empty
            Public Property IsActive As Boolean
            Public Property LastLoginText As String = String.Empty

        End Class

    End Class

End Namespace
