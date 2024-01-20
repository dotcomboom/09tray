Imports MessengerContentInstallerLibrary
Imports MessengerAPI

Public Class frmMain
    Shared api As New Messenger
    Shared process As Process

    Private Sub SetIconFromImage(image As System.Drawing.Bitmap)
        NotifyIcon1.Icon = Icon.FromHandle(image.GetHicon())
    End Sub

    Private Function MessengerRunning() As Boolean
        Try
            ' process = System.Diagnostics.Process.GetProcessesByName("msnmsgr")(0)
            modWinman.HideHiddenWindow()
            api = New Messenger
            Dim window = api.Window
            Return True
        Catch ext As System.Runtime.InteropServices.COMException
            Return False
        Catch ex As IndexOutOfRangeException
            Return False
        End Try
    End Function

    Private Sub StartMessenger()
        If Not MessengerRunning() Then
            process = process.Start("msnmsgr")
        End If
    End Sub

    Private Function Window() As IMessengerWindow
        Return api.Window
    End Function

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim w As IMessengerWindow = api.Window
        w.Show()
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Dim w As IMessengerWindow = api.Window
        w.Close()
    End Sub

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.Hide()

        process = System.Diagnostics.Process.GetProcessesByName("msnmsgr")(0)

        SetIconFromImage(My.Resources.disconnected)
        RefreshContextStatus()

        'For Each c As IMessengerContact In api.MyContacts
        '    ListBox1.Items.Add(c.FriendlyName)
        'Next
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        PropertyGrid1.SelectedObject = api
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        Dim contact As IMessengerContact = api.MyContacts(ListBox1.SelectedIndex)
        PropertyGrid1.SelectedObject = contact
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        Dim contact As IMessengerContact = api.MyContacts(ListBox1.SelectedIndex)
        api.InstantMessage(contact)
    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        Dim contact As IMessengerContact = api.MyContacts(ListBox1.SelectedIndex)
        'api.Page(contact)
    End Sub

    Private away_statuses() As MISTATUS = {MISTATUS.MISTATUS_AWAY, MISTATUS.MISTATUS_BE_RIGHT_BACK, MISTATUS.MISTATUS_IDLE}
    Private busy_statuses() As MISTATUS = {MISTATUS.MISTATUS_BUSY, MISTATUS.MISTATUS_ON_THE_PHONE}
    Private offline_statuses() As MISTATUS = {MISTATUS.MISTATUS_OFFLINE, MISTATUS.MISTATUS_UNKNOWN}

    Private Function CheckStatus(status As String)
        Dim state As MISTATUS

        If Not MessengerRunning() Then
            ' Messenger is not open!
            SetIconFromImage(My.Resources.wlm)
            SetNotifyStatus("Start Messenger")
            Return False
        End If

        state = api.MyStatus

        If status = "online" And state = MISTATUS.MISTATUS_ONLINE Then
            SetIconFromImage(My.Resources.online)
            SetNotifyStatus("Online")
            Return True
        End If
        If status = "away" And away_statuses.Contains(state) Then
            SetIconFromImage(My.Resources.away)
            SetNotifyStatus("Away")
            Return True
        End If
        If status = "busy" And busy_statuses.Contains(state) Then
            SetIconFromImage(My.Resources.busy)
            SetNotifyStatus("Busy")
            Return True
        End If
        If status = "invisible" And state = MISTATUS.MISTATUS_INVISIBLE Then
            SetIconFromImage(My.Resources.offline)
            SetNotifyStatus("Invisible")
            Return True
        End If
        If status = "offline" And offline_statuses.Contains(state) Then
            SetIconFromImage(My.Resources.disconnected)
            SetNotifyStatus("Disconnected")
            Return True
        End If
        Return False
    End Function

    Private Sub ChangeStatus(status As MISTATUS)
        api.MyStatus = status
        RefreshContextStatus()
    End Sub

    Private Sub RefreshContextStatus()
        Dim ready = MessengerRunning()

        QuitMessengerToolStripMenuItem1.Enabled = ready
        SignInToolStripMenuItem.Enabled = CheckStatus("offline")
        SignOuitToolStripMenuItem.Enabled = ready And Not CheckStatus("offline")

        itmOnline.Checked = CheckStatus("online")
        itmAway.Checked = CheckStatus("away")
        itmBusy.Checked = CheckStatus("busy")
        imInvisible.Checked = CheckStatus("invisible")

        itmOnline.Enabled = ready And Not CheckStatus("offline")
        itmAway.Enabled = ready And Not CheckStatus("offline")
        itmBusy.Enabled = ready And Not CheckStatus("offline")
        imInvisible.Enabled = ready And Not CheckStatus("offline")
    End Sub
    Private Sub SetupContextImages()
        itmOnline.Image = My.Resources.online_state
        itmAway.Image = My.Resources.away_state
        itmBusy.Image = My.Resources.busy_state
        imInvisible.Image = My.Resources.offline_state
    End Sub

    Private Sub ContextMenuStrip1_Opened(sender As System.Object, e As System.EventArgs) Handles ContextMenuStrip1.Opening
        RefreshContextStatus()
        SetupContextImages()
    End Sub

    Private Sub QuitMessengerToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs)
        modWinman.CloseHiddenWindow()
        process = System.Diagnostics.Process.GetProcessesByName("msnmsgr")(0)
        process.Kill()
    End Sub

    Private Sub OpenMessenegrToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)
        Window.Show()
    End Sub

    Private Sub StopMessenger()
        For Each process As Process In System.Diagnostics.Process.GetProcessesByName("msnmsgr")
            process.Kill()
        Next

        SetIconFromImage(My.Resources.wlm)
        SetNotifyStatus("Start Messenger")
    End Sub

    Private Sub QuitMessengerToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles QuitMessengerToolStripMenuItem1.Click
        StopMessenger()
    End Sub

    Private Sub NotifyIcon1_DoubleClick(sender As System.Object, e As System.EventArgs) Handles NotifyIcon1.DoubleClick
        If MessengerRunning() Then
            If Window.IsClosed Then
                Window.Show()
            Else
                Window.Close()
            End If
        Else
            StartMessenger()
        End If
    End Sub

    Private Sub itmOnline_Click(sender As System.Object, e As System.EventArgs) Handles itmOnline.Click
        ChangeStatus(MISTATUS.MISTATUS_ONLINE)
    End Sub

    Private Sub itmAway_Click(sender As System.Object, e As System.EventArgs) Handles itmAway.Click
        ChangeStatus(MISTATUS.MISTATUS_AWAY)
    End Sub

    Private Sub itmBusy_Click(sender As System.Object, e As System.EventArgs) Handles itmBusy.Click
        ChangeStatus(MISTATUS.MISTATUS_BUSY)
    End Sub

    Private Sub imInvisible_Click(sender As System.Object, e As System.EventArgs) Handles imInvisible.Click
        ChangeStatus(MISTATUS.MISTATUS_INVISIBLE)
    End Sub

    Private Sub SetNotifyStatus(p1 As String)
        NotifyIcon1.Text = Me.Text & " - " & p1
    End Sub

    Private Sub SignOuitToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SignOuitToolStripMenuItem.Click
        api.Signout()
    End Sub

    Private Sub SignInToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SignInToolStripMenuItem.Click
        Window.Show()
    End Sub

    Private Sub OpenMessengerToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles OpenMessengerToolStripMenuItem.Click
        If MessengerRunning() Then
            Window.Show()
        Else
            StartMessenger()
        End If
    End Sub

    Private Sub QuitToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles QuitToolStripMenuItem.Click
        StopMessenger()
        Me.Close()
    End Sub

    Private Sub GToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles GToolStripMenuItem.Click
        modWinman.ShowHiddenWindow()
        Me.Close()
    End Sub

    Private Sub SetStatusToolStripMenuItem1_DropDownOpening(sender As System.Object, e As System.EventArgs) Handles SetStatusToolStripMenuItem1.DropDownOpening
        ' Clear existing items in the dropdown menu
        SetStatusToolStripMenuItem1.DropDownItems.Clear()

        SetStatusToolStripMenuItem1.DropDownItems.Add(MusicToolStripMenuItem)
        SetStatusToolStripMenuItem1.DropDownItems.Add(GameToolStripMenuItem)
        SetStatusToolStripMenuItem1.DropDownItems.Add(OfficeToolStripMenuItem)
        SetStatusToolStripMenuItem1.DropDownItems.Add(ToolStripSeparator6)

        ' Get all running processes
        Dim processes() As Process = process.GetProcesses()

        Dim blocked_processes() As String = {"msnmsgr", "TextInputHost"}

        ' Loop through each process
        For Each process As Process In processes
            ' Check if the process has a main window title
            If Not String.IsNullOrEmpty(process.MainWindowTitle) Then
                ' Create a new dropdown item
                Dim newItem As New ToolStripMenuItem()

                ' Set the text of the dropdown item to the window title
                newItem.Text = process.MainWindowTitle

                ' Set the tag of the dropdown item to the process name
                newItem.Tag = process.ProcessName

                If newItem.Tag = monitoredProcess And beamingActivity Then
                    newItem.Checked = True
                End If

                Try

                    newItem.Image = Icon.ExtractAssociatedIcon(process.Modules(0).FileName).ToBitmap()
                Catch ex As Exception

                End Try

                ' Add a click event handler to the dropdown item
                AddHandler newItem.Click, AddressOf ToolStripMenuItem_Click

                ' Add the dropdown item to the menu
                SetStatusToolStripMenuItem1.DropDownItems.Add(newItem)
            End If
        Next

    End Sub

    Private beamingActivity As Boolean = False
    Private monitoredProcess = ""
    Private activityKind = "Music"

    Private Sub ToolStripMenuItem_Click(sender As Object, e As EventArgs)
        ' Handle the click event here
        Dim clickedItem As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)

        ' Access the process name from the Tag property
        Dim processName As String = clickedItem.Tag.ToString()
        monitoredProcess = clickedItem.Tag.ToString()
        beamingActivity = True
        updateActivity()

        'modWinman.setActivity("Music", clickedItem.Text)


        ' You can use the processName as needed (e.g., launch another process, etc.)
        'MessageBox.Show("Selected Process: " & processName)

        ' Add a click event handler to the dropdown item
        'AddHandler music.Click, AddressOf ToolStripMenuItem_Click
        'AddHandler game.Click, AddressOf ToolStripMenuItem_Click
        'AddHandler office.Click, AddressOf ToolStripMenuItem_Click

        'clickedItem.DropDownItems.Add(music)
        'clickedItem.DropDownItems.Add(game)
        'clickedItem.DropDownItems.Add(office)
    End Sub

    Private Sub updateActivity()
        Dim activity As MessengerDotNet.MessengerActivityType = MessengerDotNet.MessengerActivityType.Music
        If activityKind = "Games" Then
            activity = MessengerDotNet.MessengerActivityType.Games
        ElseIf activityKind = "Office" Then
            activity = MessengerDotNet.MessengerActivityType.Office
        End If

        Dim procs = process.GetProcessesByName(monitoredProcess)
        If procs.Count > 0 Then
            Dim proc As Process = procs(0)
            MessengerDotNet.MessengerActivity.SetActivity(activity, beamingActivity, proc.MainWindowTitle)
        Else
            MessengerDotNet.MessengerActivity.SetActivity(MessengerDotNet.MessengerActivityType.Music, False, "dis")
        End If
        Timer1.Enabled = beamingActivity
    End Sub

    Private Sub ContextMenuStrip1_Opened(sender As System.Object, e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening

    End Sub

    Private Sub MusicToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles MusicToolStripMenuItem.Click
        activityKind = "Music"
        updateActivity()
    End Sub

    Private Sub GameToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles GameToolStripMenuItem.Click
        activityKind = "Games"
        updateActivity()
    End Sub

    Private Sub OfficeToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles OfficeToolStripMenuItem.Click
        activityKind = "Office"
        updateActivity()
    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Timer1.Enabled = beamingActivity
        updateActivity()
    End Sub
End Class
