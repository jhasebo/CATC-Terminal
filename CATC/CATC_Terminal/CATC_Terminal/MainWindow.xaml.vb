Option Explicit On
Imports System
Imports System.IO
Imports System.Data
Imports System.Data.SqlServerCe
Imports System.Data.SqlClient
Imports System.Windows.Media.Animation
Imports System.Windows.Threading
Imports System.Windows.Controls.Primitives
Imports Microsoft.Win32
Imports Microsoft.Office.Interop
Class MainWindow
    Dim con As New SqlCeConnection(ConString)
    Dim command As SqlCeCommand = con.CreateCommand
    Public Shared timer As DispatcherTimer = New DispatcherTimer
    'Onload Events
    Private Sub MainWindow_Loaded(sender As Object, e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        timer.Interval = TimeSpan.FromSeconds(10)
        AddHandler timer.Tick, AddressOf timer_tick
        timer.Start()
        fillreferrals()
        Dim adapter As New SqlCeDataAdapter("SELECT Referral.Date as [Date], Referral.TraceNo as [Ref No], StudentList.SecCode as [Section], StudentList.StudentNo as [Student Number], Student.LastName as [Last Name], Student.FirstName as [First Name], Student.MiddleName as [Middle Name] FROM Referral INNER JOIN StudentList ON Referral.SLRefNum=StudentList.SLRefNum INNER JOIN Student ON StudentList.StudentNo=Student.StudentNo INNER JOIN Section ON StudentList.SecCode=Section.SecCode WHERE Student.College='" & ActiveDeptCode & "' ORDER BY Referral.Date DESC", con)
        Dim cbuilder As New SqlCeCommandBuilder(adapter)
        fillreportpicker(adapter)
    End Sub


    'Transition Events
    Private Sub setClickable()
        Dim da As New DoubleAnimation
        With da
            .To = 100
            .From = 0
            .Duration = New Duration(TimeSpan.FromSeconds(0))
        End With
        Clickable.BeginAnimation(Grid.WidthProperty, da)
    End Sub
    Private Sub btnUser_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles btnUser.Click
        ActiveTab = 1
    End Sub

    Private Sub btnReferral_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles btnReferral.Click
        ActiveTab = 2
    End Sub

    Private Sub btnReports_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles btnReports.Click
        ActiveTab = 3
        dpFrom.SelectedDate = Date.Now
        dpTo.SelectedDate = Date.Now
        expDateFilter.IsExpanded = True
    End Sub

    Private Sub gridMenu_LeftButtonUp(sender As System.Object, e As System.Windows.Input.MouseButtonEventArgs) Handles gridMenu.MouseLeftButtonUp
        Dim daW As New DoubleAnimation
        With daW
            .From = 800
            .To = 0
            .Duration = New Duration(TimeSpan.FromSeconds(0.1))
        End With
        Dim daFW As New DoubleAnimation
        With daFW
            .From = 700
            .To = 0
            .Duration = New Duration(TimeSpan.FromSeconds(0.1))
        End With
        Dim daOP As New DoubleAnimation
        With daOP
            .From = 0.6
            .To = 1
            .Duration = New Duration(TimeSpan.FromSeconds(0.1))
        End With
        If ActiveTab = 1 Then
            gridContentHolder.BeginAnimation(Grid.WidthProperty, daW)
            gridMenu.BeginAnimation(Grid.OpacityProperty, daOP)
        ElseIf ActiveTab = 2 Then
            gridReferrals.BeginAnimation(Grid.WidthProperty, daW)
            gridMenu.BeginAnimation(Grid.OpacityProperty, daOP)
        ElseIf ActiveTab = 3 Then
            gridReports.BeginAnimation(Grid.WidthProperty, daW)
            gridMenu.BeginAnimation(Grid.OpacityProperty, daOP)
        ElseIf ActiveTab = 4 Then
            gridReferrals.BeginAnimation(Grid.WidthProperty, daW)
            gridReferralDetails.BeginAnimation(Grid.WidthProperty, daFW)
            gridMenu.BeginAnimation(Grid.OpacityProperty, daOP)
        ElseIf ActiveTab = 5 Then
            gridReports.BeginAnimation(Grid.WidthProperty, daW)
            gridReportDetails.BeginAnimation(Grid.WidthProperty, daFW)
            gridMenu.BeginAnimation(Grid.OpacityProperty, daOP)
        End If
        ActiveTab = 0
    End Sub

    'Code for Profile Tab
    Private Sub btnUpdateContact_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles btnUpdateContact.Click
        Dim daUC As DoubleAnimation = New DoubleAnimation()
        With daUC
            .From = 0
            .To = 120
            .Duration = New Duration(TimeSpan.FromSeconds(0.25))
        End With
        tbUpdateContact.BeginAnimation(TextBox.WidthProperty, daUC)
        setContact.Width = 50
        CancelC.Width = 50
        btnUpdateContact.Width = 0
        tbUpdateContact.Focus()
    End Sub

    Private Sub setContact_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles setContact.Click
        timer.Stop()
        Try
            con.Close()
            con.Open()
            With command
                .CommandText = "Update CATC Set Contact = '" & tbUpdateContact.Text & "' Where EmpNum='" & ActiveUser & "'"
                .ExecuteNonQuery()
            End With
            lblContact.Content = tbUpdateContact.Text()

        Catch ex As Exception
            MessageBox.Show("Unable to Update Contact: " & ex.ToString)
        Finally
            con.Close()
            tbUpdateContact.Text = String.Empty
            Dim daUC As DoubleAnimation = New DoubleAnimation()
            With daUC
                .From = 120
                .To = 0
                .Duration = New Duration(TimeSpan.FromSeconds(0.1))
            End With
            tbUpdateContact.BeginAnimation(TextBox.WidthProperty, daUC)
            btnUpdateContact.Width = 62
            setContact.Width = 0
            CancelC.Width = 0
            btnUpdateContact.Focus()
        End Try
        timer.Start()
    End Sub

    Private Sub CancelC_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles CancelC.Click
        Dim daUC As DoubleAnimation = New DoubleAnimation()
        With daUC
            .From = 120
            .To = 0
            .Duration = New Duration(TimeSpan.FromSeconds(0.1))
        End With
        tbUpdateContact.BeginAnimation(TextBox.WidthProperty, daUC)
        btnUpdateContact.Width = 62
        setContact.Width = 0
        CancelC.Width = 0
        btnUpdateContact.Focus()
    End Sub

    Private Sub btnUpdateEmail_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles btnUpdateEmail.Click
        Dim daUC As DoubleAnimation = New DoubleAnimation()
        With daUC
            .From = 0
            .To = 120
            .Duration = New Duration(TimeSpan.FromSeconds(0.25))
        End With
        tbUpdateEmail.BeginAnimation(TextBox.WidthProperty, daUC)
        setEmail.Width = 50
        CancelE.Width = 50
        btnUpdateEmail.Width = 0
        tbUpdateEmail.Focus()
    End Sub

    Private Sub setEmail_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles setEmail.Click
        timer.Stop()
        Try
            con.Close()
            con.Open()
            With command
                .CommandText = "Update CATC Set Email = '" & tbUpdateEmail.Text & "' Where EmpNum='" & ActiveUser & "'"
                .ExecuteNonQuery()
            End With
            lblEmail.Content = tbUpdateEmail.Text.Replace("_", "__")

        Catch ex As Exception
            MessageBox.Show("Unable to Update Email: " & ex.ToString)
        Finally
            con.Close()
            tbUpdateEmail.Text = String.Empty
            Dim daUC As DoubleAnimation = New DoubleAnimation()
            With daUC
                .From = 120
                .To = 0
                .Duration = New Duration(TimeSpan.FromSeconds(0.1))
            End With
            tbUpdateEmail.BeginAnimation(TextBox.WidthProperty, daUC)
            btnUpdateEmail.Width = 62
            setEmail.Width = 0
            CancelE.Width = 0
            btnUpdateEmail.Focus()
        End Try
        timer.Start()
    End Sub

    Private Sub CancelE_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles CancelE.Click
        Dim daUC As DoubleAnimation = New DoubleAnimation()
        With daUC
            .From = 120
            .To = 0
            .Duration = New Duration(TimeSpan.FromSeconds(0.1))
        End With
        tbUpdateEmail.BeginAnimation(TextBox.WidthProperty, daUC)
        btnUpdateEmail.Width = 62
        setEmail.Width = 0
        CancelE.Width = 0
        btnUpdateEmail.Focus()
    End Sub

    Private Sub btnPicture_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles btnPicture.Click
        timer.Stop()
        Dim openpic As New Microsoft.Win32.OpenFileDialog()
        openpic.FileName = "Profile Photo"
        openpic.DefaultExt = ".jpg"
        openpic.Filter = "Image Files(*.jpg,*.png,*.bmp)|*.jpg;*.png;*.bmp|All Files (*.*)|*.*"
        Dim result As Boolean = openpic.ShowDialog()
        If result = True Then
            Try
                con.Close()
                con.Open()
                With command
                    .CommandText = "Update CATC Set PhotoFilePath = '" & openpic.FileName & "' Where EmpNum = '" & ActiveUser & "'"
                    .ExecuteNonQuery()
                End With
                imgProfile.Source = New BitmapImage(New Uri(openpic.FileName))
            Catch ex As Exception
                MessageBox.Show("Unable to Update Profile Image: " & ex.ToString)
            Finally
                con.Close()
            End Try
        End If
        timer.Start()
    End Sub

    'Code for Settings
    Private Sub btnSettings_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles btnSettings.Click
        Dim da As New DoubleAnimation
        With da
            .From = 0
            .To = 90
            .Duration = New Duration(TimeSpan.FromSeconds(0.25))
        End With
        settings.BeginAnimation(StackPanel.HeightProperty, da)
    End Sub

    Private Sub btnChangePassword_MouseEnter(sender As Object, e As System.Windows.Input.MouseEventArgs) Handles btnChangePassword.MouseEnter
        btnChangePassword.Foreground = System.Windows.Media.Brushes.Black
    End Sub

    Private Sub btnChangePassword_MouseLeave(sender As Object, e As System.Windows.Input.MouseEventArgs) Handles btnChangePassword.MouseLeave
        btnChangePassword.Foreground = System.Windows.Media.Brushes.White
    End Sub

    Private Sub btnAbout_MouseEnter(sender As Object, e As System.Windows.Input.MouseEventArgs) Handles btnAbout.MouseEnter
        btnAbout.Foreground = System.Windows.Media.Brushes.Black
    End Sub

    Private Sub btnAbout_MouseLeave(sender As Object, e As System.Windows.Input.MouseEventArgs) Handles btnAbout.MouseLeave
        btnAbout.Foreground = System.Windows.Media.Brushes.White
    End Sub

    Private Sub btnHelp_MouseEnter(sender As Object, e As System.Windows.Input.MouseEventArgs) Handles btnHelp.MouseEnter
        btnHelp.Foreground = System.Windows.Media.Brushes.Black
    End Sub

    Private Sub btnHelp_MouseLeave(sender As Object, e As System.Windows.Input.MouseEventArgs) Handles btnHelp.MouseLeave
        btnHelp.Foreground = System.Windows.Media.Brushes.White
    End Sub


    Private Sub settings_MouseLeave(sender As Object, e As System.Windows.Input.MouseEventArgs) Handles settings.MouseLeave
        Dim da As New DoubleAnimation
        With da
            .From = 90
            .To = 0
            .Duration = New Duration(TimeSpan.FromSeconds(0.1))
        End With
        settings.BeginAnimation(StackPanel.HeightProperty, da)
    End Sub
    'Code For Referrals
    Private Sub timer_tick()
        command.CommandText = "Select Count(*) FROM Referral INNER JOIN StudentList ON Referral.SLRefNum=StudentList.SLRefNum INNER JOIN Student ON StudentList.StudentNo=Student.StudentNo WHERE Referral.ActionTaken is null AND Student.College='" & ActiveDeptCode & "'"
        con.Close()
        con.Open()
        Dim ctr As Integer = command.ExecuteScalar()
        con.Close()
        If ctr = 0 Then
            cNotify.Visibility = Windows.Visibility.Hidden
            dgReferrals.Visibility = Windows.Visibility.Hidden
            tbNoNewNotif.Visibility = Windows.Visibility.Visible
        Else
            cNotify.Visibility = Windows.Visibility.Visible
            tbDisplayNotifCtr.Text = ctr
            dgReferrals.Visibility = Windows.Visibility.Visible
            tbNoNewNotif.Visibility = Windows.Visibility.Hidden
            fillreferrals()
        End If
    End Sub

    Private Sub fillreferrals()
        timer.Stop()
        con.Close()
        con.Open()
        Dim ds As New DataSet
        Dim da As New SqlCeDataAdapter("SELECT Referral.Date as [Date], Referral.TraceNo as [Ref No], StudentList.SecCode as [Section], StudentList.StudentNo as [Student Number], Student.LastName as [Last Name], Student.FirstName as [First Name], Student.MiddleName as [Middle Name] FROM Referral INNER JOIN StudentList ON Referral.SLRefNum=StudentList.SLRefNum INNER JOIN Student ON StudentList.StudentNo=Student.StudentNo INNER JOIN Section ON StudentList.SecCode=Section.SecCode WHERE (Referral.ActionTaken is null and Student.College='" & ActiveDeptCode & "')", con)
        Dim cb As New SqlCeCommandBuilder(da)
        da.Fill(ds, "Referral")
        dgReferrals.ItemsSource = ds.Tables(0).DefaultView
        dgReferrals.DataContext = ds.Tables(0)
        con.Close()
        timer.Start()
    End Sub

    Private Sub dgReferrals_AutoGeneratingColumn(sender As Object, e As System.Windows.Controls.DataGridAutoGeneratingColumnEventArgs) Handles dgReferrals.AutoGeneratingColumn
        If e.PropertyType = GetType(System.DateTime) Then
            TryCast(e.Column, DataGridTextColumn).Binding.StringFormat = "dd/MM/yyyy"
        End If
    End Sub

    Private Sub dgReferrals_SelectionChanged(sender As System.Object, e As System.Windows.Controls.SelectionChangedEventArgs) Handles dgReferrals.SelectionChanged
        timer.Stop()
        ActiveTab = 4
        setClickable()
        con.Close()
        con.Open()
        If dgReferrals.SelectedItems.Count > 0 Then
            Dim row As DataRowView = dgReferrals.SelectedItems(0)
            lblRefRefNum.Content = row("Ref No")
            lblRefRefNum.Foreground = Brushes.White
            Dim ds As New DataSet
            Dim dsd As New DataSet
            Dim adapter As New SqlCeDataAdapter("Select Student.LastName as [SLastName], Student.FirstName as [SFirstName], Student.MiddleName as [SMidName], StudentList.StudentNo, Section.SubjCode, Subject.Description, StudentList.SecCode, Referral.Concerns, Student.Guardian, Student.GAddr, Student.GTel, Student.GMobile, Faculty.LastName as [FLastName], Faculty.FirstName as [FFirstName], Faculty.MiddleName as [FMidName], Faculty.Contact as [FContact] FROM Referral INNER JOIN StudentList ON Referral.SLRefNum=StudentList.SLRefNum INNER JOIN Student ON StudentList.StudentNo=Student.StudentNo INNER JOIN Section on StudentList.SecCode=Section.SecCode INNER JOIN Subject on Section.SubjCode=Subject.SubjCode INNER JOIN Faculty on Section.EmpNum=Faculty.EmpNum WHERE Referral.TraceNo=" & lblRefRefNum.Content, con)
            Dim adapterd As New SqlCeDataAdapter("Select Attendance.Date as [Date] FROM ReferralDates INNER JOIN Attendance ON ReferralDates.AttendanceRefNum=Attendance.AttendanceRefNum WHERE TraceNo=" & lblRefRefNum.Content & "ORDER BY Attendance.Date ASC", con)
            Dim cbuilder As New SqlCeCommandBuilder(adapter)
            Dim cbuilderd As New SqlCeCommandBuilder(adapterd)
            adapter.Fill(ds, "Referral")
            adapterd.Fill(dsd, "ReferralDates")
            Dim i As Integer = 0
            Dim dateholder As String
            tbRefDates.Text = String.Empty
            While i < dsd.Tables(0).Rows.Count
                dateholder = dsd.Tables(0).Rows(i).Item(0).ToString
                tbRefDates.Text = tbRefDates.Text & Date.Parse(dateholder).ToString("dd/MM/yyyy") & ", "
                i = i + 1
            End While
            With ds.Tables(0).Rows(0)
                lblRefName.Content = .Item(0).ToString & ", " & .Item(1).ToString & " " & .Item(2).ToString
                lblRefStudentNo.Content = .Item(3).ToString
                lblRefSubject.Content = .Item(4).ToString & " - " & .Item(5).ToString
                lblRefSection.Content = .Item(6).ToString
                lblRefGuardianName.Content = .Item(8).ToString
                tbRefAddress.Text = .Item(9).ToString
                lblRefTelNum.Content = .Item(10).ToString
                lblRefMobNum.Content = .Item(11).ToString
                lblRefby.Content = .Item(12).ToString & ", " & .Item(13).ToString & " " & .Item(14).ToString
                lblRefbyContact.Content = .Item(15).ToString
                If .Item("Concerns").ToString().Equals(String.Empty) Then
                    tbRefConcerns.Text = """No faculty concerns were set..."""
                Else
                    tbRefConcerns.Text = .Item("Concerns").ToString()
                End If
            End With
            Dim da As New DoubleAnimation
            With da
                .To = 700
                .From = 0
                .Duration = New Duration(TimeSpan.FromSeconds(0.25))
            End With
            gridReferralDetails.BeginAnimation(Grid.WidthProperty, da)
            expNotifFeedback.IsExpanded = True
            con.Close()
        End If
    End Sub

    Private Sub btnRefClear_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles btnRefClear.Click
        tbFeedback.Text = String.Empty
    End Sub

    Private Sub tbFeedback_TextChanged(sender As System.Object, e As System.Windows.Controls.TextChangedEventArgs) Handles tbFeedback.TextChanged
        tbFeedback.Text = tbFeedback.Text.Replace("'", "")
        If tbFeedback.Text.Length > 200 Then
            tbFeedback.Text = tbFeedback.Text.Remove(200)
        End If
        lblRefFeedCtr.Content = Val(200 - tbFeedback.Text.Length).ToString
        If Val(lblRefFeedCtr.Content) < 20 Then
            lblRefFeedCtr.Foreground = System.Windows.Media.Brushes.Red
        Else
            lblRefFeedCtr.Foreground = System.Windows.Media.Brushes.White
        End If
        tbFeedback.Select(tbFeedback.Text.Length, 0)
    End Sub

    Private Sub btnRefSubmit_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles btnRefSubmit.Click
        con.Close()
        con.Open()
        command.CommandText = "UPDATE Referral SET ActionTaken = '" & tbFeedback.Text & "' WHERE TraceNo=" & lblRefRefNum.Content
        command.ExecuteNonQuery()
        command.CommandText = "UPDATE Referral SET ATby = '" & ActiveUser & "' WHERE TraceNo =" & lblRefRefNum.Content
        command.ExecuteNonQuery()
        MessageBox.Show("Successfully taken action on Referral..." & lblRefRefNum.Content & ".", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
        command.CommandText = "INSERT INTO FeedbackRequest(TraceNo, Finished) Values(" & lblRefRefNum.Content & ", null)"
        command.ExecuteNonQuery()
        con.Close()
        fillreferrals()
    End Sub


    Private Sub Clickable_LeftButtonUp(sender As System.Object, e As System.Windows.Input.MouseButtonEventArgs) Handles Clickable.MouseLeftButtonUp
        If ActiveTab = 4 Then
            Dim dafw As New DoubleAnimation
            With dafw
                .From = 700
                .To = 0
                .Duration = New Duration(TimeSpan.FromSeconds(0.1))
            End With
            Dim da As New DoubleAnimation
            With da
                .To = 0
                .From = 70
                .Duration = New Duration(TimeSpan.FromSeconds(0))
            End With
            Clickable.BeginAnimation(Grid.WidthProperty, da)
            gridReferralDetails.BeginAnimation(Grid.WidthProperty, dafw)
            ActiveTab = 2
            timer.Start()
        ElseIf ActiveTab = 5 Then
            Dim dafw As New DoubleAnimation
            With dafw
                .From = 700
                .To = 0
                .Duration = New Duration(TimeSpan.FromSeconds(0.1))
            End With
            Dim da As New DoubleAnimation
            With da
                .To = 0
                .From = 70
                .Duration = New Duration(TimeSpan.FromSeconds(0))
            End With
            Clickable.BeginAnimation(Grid.WidthProperty, da)
            gridReportDetails.BeginAnimation(Grid.WidthProperty, dafw)
            ActiveTab = 3
            timer.Start()
        End If
    End Sub

    Private Sub expNotifActionTaken_Expanded(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles expNotifActionTaken.Expanded
        expNotifFeedback.IsExpanded = False
    End Sub

    Private Sub expNotifFeedback_Expanded(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles expNotifFeedback.Expanded
        expNotifActionTaken.IsExpanded = False
    End Sub

    Private Sub fillreportpicker(adapter As SqlCeDataAdapter)
        Dim dataset As New DataSet
        adapter.Fill(dataset)
        dgReportPicker.ItemsSource = dataset.Tables(0).DefaultView
        dgReportPicker.DataContext = dataset.Tables(0)
    End Sub

    Private Sub dgReportPicker_AutoGeneratingColumn(sender As Object, e As System.Windows.Controls.DataGridAutoGeneratingColumnEventArgs) Handles dgReportPicker.AutoGeneratingColumn
        If e.PropertyType = GetType(System.DateTime) Then
            TryCast(e.Column, DataGridTextColumn).Binding.StringFormat = "dd/MM/yyyy"
        End If
    End Sub

    Private Sub tbReportSearch_TextChanged(sender As System.Object, e As System.Windows.Controls.TextChangedEventArgs) Handles tbReportSearch.TextChanged
        If dpFrom.SelectedDate.HasValue = True And dpTo.SelectedDate.HasValue = True Then
            Dim adapter As New SqlCeDataAdapter("SELECT Referral.Date as [Date], Referral.TraceNo as [Ref No], StudentList.SecCode as [Section], StudentList.StudentNo as [Student Number], Student.LastName as [Last Name], Student.FirstName as [First Name], Student.MiddleName as [Middle Name] FROM Referral INNER JOIN StudentList ON Referral.SLRefNum=StudentList.SLRefNum INNER JOIN Student ON StudentList.StudentNo=Student.StudentNo INNER JOIN Section ON StudentList.SecCode=Section.SecCode WHERE (Student.College='" & ActiveDeptCode & "' AND (Referral.Date BETWEEN {d '" & dpFrom.SelectedDate.Value.ToString("yyyy-MM-dd") & "'} AND {d '" & dpTo.SelectedDate.Value.ToString("yyyy-MM-dd") & "'}) AND (Referral.TraceNo LIKE '%" & tbReportSearch.Text & "%' OR Student.StudentNo LIKE '%" & tbReportSearch.Text & "%' OR Student.LastName LIKE '%" & tbReportSearch.Text & "%' OR Student.FirstName Like '%" & tbReportSearch.Text & "%' OR StudentList.SecCode Like '%" & tbReportSearch.Text & "%')) ORDER BY Referral.Date DESC", con)
            Dim cbuilder As New SqlCeCommandBuilder(adapter)
            fillreportpicker(adapter)
        End If
    End Sub

    Private Sub dpFrom_SelectedDateChanged(sender As Object, e As System.Windows.Controls.SelectionChangedEventArgs) Handles dpFrom.SelectedDateChanged
        If dpFrom.SelectedDate.HasValue = True And dpTo.SelectedDate.HasValue = True Then
            Dim adapter As New SqlCeDataAdapter("SELECT Referral.Date as [Date], Referral.TraceNo as [Ref No], StudentList.SecCode as [Section], StudentList.StudentNo as [Student Number], Student.LastName as [Last Name], Student.FirstName as [First Name], Student.MiddleName as [Middle Name] FROM Referral INNER JOIN StudentList ON Referral.SLRefNum=StudentList.SLRefNum INNER JOIN Student ON StudentList.StudentNo=Student.StudentNo INNER JOIN Section ON StudentList.SecCode=Section.SecCode WHERE (Student.College='" & ActiveDeptCode & "' AND (Referral.Date BETWEEN {d '" & dpFrom.SelectedDate.Value.ToString("yyyy-MM-dd") & "'} AND {d '" & dpTo.SelectedDate.Value.ToString("yyyy-MM-dd") & "'}) AND (Referral.TraceNo LIKE '%" & tbReportSearch.Text & "%' OR Student.StudentNo LIKE '%" & tbReportSearch.Text & "%' OR Student.LastName LIKE '%" & tbReportSearch.Text & "%' OR Student.FirstName Like '%" & tbReportSearch.Text & "%' OR StudentList.SecCode Like '%" & tbReportSearch.Text & "%')) ORDER BY Referral.Date DESC", con)
            Dim cbuilder As New SqlCeCommandBuilder(adapter)
            fillreportpicker(adapter)
        End If
    End Sub
    Private Sub dpTo_SelectedDateChanged(sender As System.Object, e As System.Windows.Controls.SelectionChangedEventArgs) Handles dpTo.SelectedDateChanged
        If dpFrom.SelectedDate.HasValue = True And dpTo.SelectedDate.HasValue = True Then
            Dim adapter As New SqlCeDataAdapter("SELECT Referral.Date as [Date], Referral.TraceNo as [Ref No], StudentList.SecCode as [Section], StudentList.StudentNo as [Student Number], Student.LastName as [Last Name], Student.FirstName as [First Name], Student.MiddleName as [Middle Name] FROM Referral INNER JOIN StudentList ON Referral.SLRefNum=StudentList.SLRefNum INNER JOIN Student ON StudentList.StudentNo=Student.StudentNo INNER JOIN Section ON StudentList.SecCode=Section.SecCode WHERE (Student.College='" & ActiveDeptCode & "' AND (Referral.Date BETWEEN {d '" & dpFrom.SelectedDate.Value.ToString("yyyy-MM-dd") & "'} AND {d '" & dpTo.SelectedDate.Value.ToString("yyyy-MM-dd") & "'}) AND (Referral.TraceNo LIKE '%" & tbReportSearch.Text & "%' OR Student.StudentNo LIKE '%" & tbReportSearch.Text & "%' OR Student.LastName LIKE '%" & tbReportSearch.Text & "%' OR Student.FirstName Like '%" & tbReportSearch.Text & "%' OR StudentList.SecCode Like '%" & tbReportSearch.Text & "%')) ORDER BY Referral.Date DESC", con)
            Dim cbuilder As New SqlCeCommandBuilder(adapter)
            fillreportpicker(adapter)
        End If
    End Sub
    Private Sub DateFilter_MouseLeave(sender As System.Object, e As System.Windows.Input.MouseEventArgs) Handles Grid1.MouseLeave
        expDateFilter.IsExpanded = False
    End Sub


    Private Sub dgReportPicker_SelectionChanged(sender As Object, e As System.Windows.Controls.SelectionChangedEventArgs) Handles dgReportPicker.SelectionChanged
        timer.Stop()
        ActiveTab = 5
        setClickable()
        con.Close()
        con.Open()
        If dgReportPicker.SelectedItems.Count > 0 Then
            Dim row As DataRowView = dgReportPicker.SelectedItem
            lblRefRefNum.Content = row("Ref No")
            lblRefRefNum.Foreground = Brushes.White
            command.CommandText = "Select Date from Referral WHERE TraceNo=" & lblRefRefNum.Content
            ActiveReferralDate = Date.Parse(command.ExecuteScalar().ToString).ToString("MM/dd/yyyy")
            Dim ds As New DataSet
            Dim dsd As New DataSet
            Dim adapter As New SqlCeDataAdapter("Select Student.LastName as [SLastName], Student.FirstName as [SFirstName], Student.MiddleName as [SMidName], StudentList.StudentNo, Section.SubjCode, Subject.Description, StudentList.SecCode, Referral.Concerns, Referral.ActionTaken, Referral.Feedback, Faculty.LastName as [FLastName], Faculty.FirstName as [FFirstName], Faculty.MiddleName as [FMidName], Faculty.Contact as [FContact] FROM Referral INNER JOIN StudentList ON Referral.SLRefNum=StudentList.SLRefNum INNER JOIN Student ON StudentList.StudentNo=Student.StudentNo INNER JOIN Section on StudentList.SecCode=Section.SecCode INNER JOIN Subject on Section.SubjCode=Subject.SubjCode INNER JOIN Faculty on Section.EmpNum=Faculty.EmpNum WHERE Referral.TraceNo=" & lblRefRefNum.Content, con)
            Dim adapterd As New SqlCeDataAdapter("Select Attendance.Date as [Date] FROM ReferralDates INNER JOIN Attendance ON ReferralDates.AttendanceRefNum=Attendance.AttendanceRefNum WHERE TraceNo=" & lblRefRefNum.Content & "ORDER BY Attendance.Date ASC", con)
            Dim cbuilder As New SqlCeCommandBuilder(adapter)
            Dim cbuilderd As New SqlCeCommandBuilder(adapterd)
            adapter.Fill(ds, "Referral")
            adapterd.Fill(dsd, "ReferralDates")
            Dim i As Integer = 0
            Dim dateholder As String
            tbRepDates.Text = String.Empty
            While i < dsd.Tables(0).Rows.Count
                dateholder = dsd.Tables(0).Rows(i).Item(0).ToString
                tbRepDates.Text = tbRepDates.Text & Date.Parse(dateholder).ToString("dd/MM/yyyy") & ", "
                i = i + 1
            End While
            With ds.Tables(0).Rows(0)
                lblRepName.Content = .Item(0).ToString & ", " & .Item(1).ToString & " " & .Item(2).ToString
                lblRepStudentNo.Content = .Item(3).ToString
                lblRepSubject.Content = .Item(4).ToString & " - " & .Item(5).ToString
                lblRepSection.Content = .Item(6).ToString
                tbRepActionTaken.Text = .Item(8).ToString
                tbRepFeedback.Text = .Item(9).ToString
                lblRepby.Content = .Item(10).ToString & ", " & .Item(11).ToString & " " & .Item(12).ToString
                If .Item(7).ToString().Equals(String.Empty) Then
                    tbRepConcerns.Text = """No faculty concerns were set..."""
                Else
                    tbRepConcerns.Text = .Item("Concerns").ToString()
                End If
            End With
            Dim da As New DoubleAnimation
            With da
                .To = 700
                .From = 0
                .Duration = New Duration(TimeSpan.FromSeconds(0.25))
            End With
            gridReportDetails.BeginAnimation(Grid.WidthProperty, da)
            con.Close()
        End If
    End Sub

    Private Sub btnExport_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles btnExport.Click
        Dim xlApp As Microsoft.Office.Interop.Excel.Application
        Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook
        Dim xlWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        xlApp = New Microsoft.Office.Interop.Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")

        'Layout Proper
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(1, 10)))
        xlWorkSheet.Cells(1, 1) = "Referral Record"
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(2, 1), xlWorkSheet.Cells(2, 2)))
        xlWorkSheet.Cells(2, 1) = "Referral Date"
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(2, 3), xlWorkSheet.Cells(2, 4)))
        xlWorkSheet.Cells(2, 3).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        xlWorkSheet.Cells(2, 3) = ActiveReferralDate
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(3, 1), xlWorkSheet.Cells(3, 2)))
        xlWorkSheet.Cells(3, 1).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
        xlWorkSheet.Cells(3, 1) = "Name of Student"
        xlWorkSheet.Cells(3, 3) = lblRepName.Content
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(3, 6), xlWorkSheet.Cells(3, 7)))
        xlWorkSheet.Cells(3, 6).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
        xlWorkSheet.Cells(3, 6) = "Student No."
        xlWorkSheet.Cells(3, 8) = lblRepStudentNo.Content
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(5, 1), xlWorkSheet.Cells(5, 2)))
        xlWorkSheet.Cells(5, 1) = "Subject"
        xlWorkSheet.Cells(5, 3) = lblRepSubject.Content
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(5, 6), xlWorkSheet.Cells(5, 7)))
        xlWorkSheet.Cells(5, 6).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        xlWorkSheet.Cells(5, 6) = "Section"
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(5, 8), xlWorkSheet.Cells(5, 10)))
        xlWorkSheet.Cells(5, 8).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        xlWorkSheet.Cells(5, 8) = lblRepSection.Content
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(6, 1), xlWorkSheet.Cells(6, 2)))
        xlWorkSheet.Cells(6, 2).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        xlWorkSheet.Cells(6, 1) = "Faculty Concerns"
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(7, 2), xlWorkSheet.Cells(11, 5)))
        xlWorkSheet.Cells(7, 2).HorizontalAlignment = Excel.XlHAlign.xlHAlignJustify
        xlWorkSheet.Cells(7, 2) = tbRepConcerns.Text
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(6, 6), xlWorkSheet.Cells(6, 7)))
        xlWorkSheet.Cells(6, 6).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        xlWorkSheet.Cells(6, 6) = "Dates Absent/Late"
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(7, 7), xlWorkSheet.Cells(11, 10)))
        xlWorkSheet.Cells(7, 7).HorizontalAlignment = Excel.XlHAlign.xlHAlignJustify
        xlWorkSheet.Cells(7, 7) = tbRepDates.Text
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(12, 1), xlWorkSheet.Cells(12, 2)))
        xlWorkSheet.Cells(12, 2).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
        xlWorkSheet.Cells(12, 1) = "Reffered by"
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(12, 3), xlWorkSheet.Cells(12, 8)))
        xlWorkSheet.Cells(12, 3).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        xlWorkSheet.Cells(12, 3) = lblRepby.Content
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(13, 3), xlWorkSheet.Cells(13, 8)))
        xlWorkSheet.Cells(13, 3).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        xlWorkSheet.Cells(13, 3) = lblaffiliation.Content
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(14, 3), xlWorkSheet.Cells(14, 7)))
        xlWorkSheet.Cells(14, 3).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        xlWorkSheet.Cells(14, 3) = "Faculty/Department/contact number"
        'contact
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(15, 1), xlWorkSheet.Cells(15, 2)))
        xlWorkSheet.Cells(15, 1).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        xlWorkSheet.Cells(15, 1) = "Name of Parent"
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(15, 3), xlWorkSheet.Cells(15, 6)))
        xlWorkSheet.Cells(15, 3).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        'xlWorkSheet.Cells(15, 3) =
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(16, 1), xlWorkSheet.Cells(16, 2)))
        xlWorkSheet.Cells(16, 1).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        xlWorkSheet.Cells(16, 1) = "Address"
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(16, 3), xlWorkSheet.Cells(16, 6)))
        xlWorkSheet.Cells(16, 3).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        'xlWorkSheet.Cells(16, 3) =
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(17, 1), xlWorkSheet.Cells(17, 2)))
        xlWorkSheet.Cells(17, 1).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        xlWorkSheet.Cells(17, 1) = "Mobile Number"
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(17, 3), xlWorkSheet.Cells(17, 6)))
        xlWorkSheet.Cells(17, 3).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        'xlWorkSheet.Cells(17, 3) = 
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(18, 1), xlWorkSheet.Cells(18, 2)))
        xlWorkSheet.Cells(18, 2).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        xlWorkSheet.Cells(18, 1) = "Action Taken"
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(19, 2), xlWorkSheet.Cells(23, 9)))
        xlWorkSheet.Cells(19, 2).HorizontalAlignment = Excel.XlHAlign.xlHAlignJustify
        xlWorkSheet.Cells(19, 2) = tbRepActionTaken.Text
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(24, 1), xlWorkSheet.Cells(24, 2)))
        xlWorkSheet.Cells(24, 2).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        xlWorkSheet.Cells(24, 1) = "Faculty Feedback"
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(25, 2), xlWorkSheet.Cells(29, 9)))
        xlWorkSheet.Cells(25, 2).HorizontalAlignment = Excel.XlHAlign.xlHAlignJustify
        xlWorkSheet.Cells(25, 2) = tbRepFeedback.Text

        Dim exportSaveFileDialog As New SaveFileDialog()
        exportSaveFileDialog.Title = "Select Excel File"
        exportSaveFileDialog.Filter = "Microsoft Office Excel Workbook(*.xlsx)|*.xlsx"
        Dim result As Nullable(Of Boolean) = exportSaveFileDialog.ShowDialog()
        If result = True Then
            Dim fullFileName As String = exportSaveFileDialog.FileName
            xlWorkBook.SaveAs(fullFileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, System.Reflection.Missing.Value, misValue, False, False, Excel.XlSaveAsAccessMode.xlShared, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlUserResolution, True, misValue, misValue)
            xlWorkBook.Saved = True
            MessageBox.Show("Exported successfully", "Exported to Excel", MessageBoxButton.OK, MessageBoxImage.Information)
        End If

        xlWorkBook.Close()
        xlApp.Quit()
    End Sub

    Private Sub btnPrint_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles btnPrint.Click
        Dim xlApp As Microsoft.Office.Interop.Excel.Application
        Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook
        Dim xlWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        xlApp = New Microsoft.Office.Interop.Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")

        'Layout Proper
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(1, 10)))
        xlWorkSheet.Cells(1, 1) = "Referral Record"
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(2, 1), xlWorkSheet.Cells(2, 2)))
        xlWorkSheet.Cells(2, 1) = "Referral Date"
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(2, 3), xlWorkSheet.Cells(2, 4)))
        xlWorkSheet.Cells(2, 3).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        xlWorkSheet.Cells(2, 3) = ActiveReferralDate
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(3, 1), xlWorkSheet.Cells(3, 2)))
        xlWorkSheet.Cells(3, 1).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
        xlWorkSheet.Cells(3, 1) = "Name of Student"
        xlWorkSheet.Cells(3, 3) = lblRepName.Content
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(3, 6), xlWorkSheet.Cells(3, 7)))
        xlWorkSheet.Cells(3, 6).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
        xlWorkSheet.Cells(3, 6) = "Student No."
        xlWorkSheet.Cells(3, 8) = lblRepStudentNo.Content
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(5, 1), xlWorkSheet.Cells(5, 2)))
        xlWorkSheet.Cells(5, 1) = "Subject"
        xlWorkSheet.Cells(5, 3) = lblRepSubject.Content
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(5, 6), xlWorkSheet.Cells(5, 7)))
        xlWorkSheet.Cells(5, 6).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        xlWorkSheet.Cells(5, 6) = "Section"
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(5, 8), xlWorkSheet.Cells(5, 10)))
        xlWorkSheet.Cells(5, 8).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        xlWorkSheet.Cells(5, 8) = lblRepSection.Content
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(6, 1), xlWorkSheet.Cells(6, 2)))
        xlWorkSheet.Cells(6, 2).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        xlWorkSheet.Cells(6, 1) = "Faculty Concerns"
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(7, 2), xlWorkSheet.Cells(11, 5)))
        xlWorkSheet.Cells(7, 2).HorizontalAlignment = Excel.XlHAlign.xlHAlignJustify
        xlWorkSheet.Cells(7, 2) = tbRepConcerns.Text
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(6, 6), xlWorkSheet.Cells(6, 7)))
        xlWorkSheet.Cells(6, 6).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        xlWorkSheet.Cells(6, 6) = "Dates Absent/Late"
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(7, 7), xlWorkSheet.Cells(11, 10)))
        xlWorkSheet.Cells(7, 7).HorizontalAlignment = Excel.XlHAlign.xlHAlignJustify
        xlWorkSheet.Cells(7, 7) = tbRepDates.Text
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(12, 1), xlWorkSheet.Cells(12, 2)))
        xlWorkSheet.Cells(12, 2).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
        xlWorkSheet.Cells(12, 1) = "Reffered by"
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(12, 3), xlWorkSheet.Cells(12, 8)))
        xlWorkSheet.Cells(12, 3).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        xlWorkSheet.Cells(12, 3) = lblRepby.Content
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(13, 3), xlWorkSheet.Cells(13, 8)))
        xlWorkSheet.Cells(13, 3).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        xlWorkSheet.Cells(13, 3) = lblaffiliation.Content
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(14, 3), xlWorkSheet.Cells(14, 7)))
        xlWorkSheet.Cells(14, 3).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        xlWorkSheet.Cells(14, 3) = "Faculty/Department/contact number"
        'contact
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(15, 1), xlWorkSheet.Cells(15, 2)))
        xlWorkSheet.Cells(15, 1).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        xlWorkSheet.Cells(15, 1) = "Name of Parent"
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(15, 3), xlWorkSheet.Cells(15, 6)))
        xlWorkSheet.Cells(15, 3).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        'xlWorkSheet.Cells(15, 3) =
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(16, 1), xlWorkSheet.Cells(16, 2)))
        xlWorkSheet.Cells(16, 1).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        xlWorkSheet.Cells(16, 1) = "Address"
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(16, 3), xlWorkSheet.Cells(16, 6)))
        xlWorkSheet.Cells(16, 3).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        'xlWorkSheet.Cells(16, 3) =
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(17, 1), xlWorkSheet.Cells(17, 2)))
        xlWorkSheet.Cells(17, 1).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        xlWorkSheet.Cells(17, 1) = "Mobile Number"
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(17, 3), xlWorkSheet.Cells(17, 6)))
        xlWorkSheet.Cells(17, 3).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        'xlWorkSheet.Cells(17, 3) = 
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(18, 1), xlWorkSheet.Cells(18, 2)))
        xlWorkSheet.Cells(18, 2).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        xlWorkSheet.Cells(18, 1) = "Action Taken"
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(19, 2), xlWorkSheet.Cells(23, 9)))
        xlWorkSheet.Cells(19, 2).HorizontalAlignment = Excel.XlHAlign.xlHAlignJustify
        xlWorkSheet.Cells(19, 2) = tbRepActionTaken.Text
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(24, 1), xlWorkSheet.Cells(24, 2)))
        xlWorkSheet.Cells(24, 2).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        xlWorkSheet.Cells(24, 1) = "Faculty Feedback"
        MergeAndCenter(xlWorkSheet.Range(xlWorkSheet.Cells(25, 2), xlWorkSheet.Cells(29, 9)))
        xlWorkSheet.Cells(25, 2).HorizontalAlignment = Excel.XlHAlign.xlHAlignJustify
        xlWorkSheet.Cells(25, 2) = tbRepFeedback.Text


        xlWorkSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait
        xlWorkSheet.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperLegal
        xlApp.Visible = True
        xlApp.Dialogs(Excel.XlBuiltInDialog.xlDialogPrint).Show()
    End Sub

    Public Sub MergeAndCenter(ByVal MergeRange As Microsoft.Office.Interop.Excel.Range)
        MergeRange.[Select]()
        MergeRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        MergeRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
        MergeRange.WrapText = False
        MergeRange.Orientation = 0
        MergeRange.AddIndent = False
        MergeRange.IndentLevel = 0
        MergeRange.ShrinkToFit = False
        MergeRange.ReadingOrder = CInt(Excel.Constants.xlContext)
        MergeRange.MergeCells = False

        MergeRange.Merge(System.Type.Missing)
    End Sub

    Private Sub expReportActionTaken_Expanded(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles expReportActionTaken.Expanded
        expReportFeedback.IsExpanded = False
    End Sub
End Class
