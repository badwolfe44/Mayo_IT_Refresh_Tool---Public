' Developed by: Jacob C. Baker - baker.jacob@mayo.edu/Badwolfebw@gmail.com
' I've added lots of comments to help whoever takes over understand this
' I'd recommend changing your comment color for ease of reading
' And yes, I know the code is kinda jank, this was started as a small project and grew far far bigger than anticipated,
' so its not the easiest to read. I've done my best to clean it,
' but this is built outside my job parameters and I have limited time to work on itcd

Imports Microsoft.Office.Interop
Imports System.Data.SqlClient
Imports System.Text
Imports System.Security.Principal
Imports System.IO
Imports System.Runtime.InteropServices
Imports File = System.IO.File
Imports System.Threading
Imports System.ComponentModel
Imports System.Globalization
Imports System.Windows.Forms.VisualStyles
Imports Microsoft.SharePoint.Client.Discovery

Public Class Form1

    Inherits Form

    Public firstName As String
    Public lastName As String
    Public lanid As String

    Private myConn As SqlConnection
    Private myCmd As SqlCommand
    Private myReader As SqlDataReader
    Private results As String

    Private itemList As New List(Of String)()
    Private totalItemCount As Integer = 0
    Private failedCount As Integer = 0
    Private successfulCount As Integer = 0

    Private domain As String = "blank"
    Private userName As String = "blank"
    Private password As String = "blank"

    Private connectionString As String = "Data Source=blank;Initial Catalog=master;Integrated Security=True;"

    Dim oldRsStr As String()
    Dim newRsStr As String()
    Dim buildingsStr As String()
    Dim floorsStr As String()
    Dim roomsStr As String()

    Dim control As Integer

    Private lastDoubleClickTime As DateTime = DateTime.MinValue

    Private WithEvents backgroundWorker As New BackgroundWorker()

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        SetDefaultText(searchSNTxt, "Device ID")
        SetDefaultText(peopleSearchTxt, "LAN ID")

        RefreshCheckbox()

    End Sub

    Public Function RefreshCheckbox()
        CheckedListBox1.Items.Clear()
        Dim path As String = "blankEmailRecipients.txt"
        Dim applicationsToCheck As New List(Of String)

        Try
            applicationsToCheck = File.ReadAllLines(path).ToList()
        Catch ex As Exception
            MessageBox.Show("An error occurred: " & ex.Message)
        End Try

        For Each email As String In applicationsToCheck
            CheckedListBox1.Items.Add(email)
        Next
    End Function


    ' Handles the click event of the AutoDiscoverBtn button.
    Private Sub AutoDiscoverBtn_Click(sender As Object, e As EventArgs) Handles AutoDiscoverBtn.Click

        ' Extracting the text from AutoDiscoverTxt and splitting it into lines.
        Dim s As String = AutoDiscoverTxt.Text
        Dim words As String() = s.Replace(Chr(13), "").Split(Chr(10))

        ' Launching Chrome browser and pausing execution for 1 second to allow Chrome to open.
        Process.Start("chrome")
        Threading.Thread.Sleep(1000)

        ' Iterating through each line from the text box.
        For Each word As String In words
            ' For lines with 2 or more characters, open a new Chrome tab with a specific URL, including the line text.
            If word.Length >= 2 Then
                Process.Start("chrome", "blank" + word)
                Threading.Thread.Sleep(1000)
            End If
        Next

        ' Additional pause after processing all lines.
        Threading.Thread.Sleep(1000)
        ' Using keyboard shortcuts to switch to the next tab and then close the current tab in Chrome.
        SendKeys.Send("^{TAB}")
        SendKeys.Send("^w")

    End Sub


    Private Sub AutoDiscoverClear_Click(sender As Object, e As EventArgs) Handles AutoDiscoverClear.Click

        ' Clear AutoDiscoverTxt
        AutoDiscoverTxt.Clear()

    End Sub

    ' Handles the click event of the PingBtn button.
    Private Sub PingBtn_Click(sender As Object, e As EventArgs) Handles PingBtn.Click

        ' Extracts text from PingTxt, removes carriage returns, and splits it into individual IP addresses.
        Dim ipList As String = PingTxt.Text
        Dim ips As String() = ipList.Replace(Chr(13), "").Split(Chr(10))

        ' Iterates through each IP address.
        For Each ip As String In ips
            ip = ip.Replace(" ", "")
            Try
                ' Checks if the IP address string is valid (at least 2 characters long).
                If ip.Length >= 2 Then
                    ' Pings the IP address with a timeout of 100 milliseconds.
                    If My.Computer.Network.Ping(ip, 100) Then
                        ' If ping is successful, adds the IP with success status and its resolved address to PingList.
                        ' Increments the count of successful pings.
                        PingList.Items.Add(ip + " - Success - IP:" + Net.Dns.GetHostEntry(ip).AddressList.GetValue(0).ToString)
                        successfulCount += 1
                    Else
                        ' If ping fails (timeout or unreachable), adds the IP with failure status and its resolved address to PingList.
                        ' Increments the count of failed pings.
                        PingList.Items.Add(ip + " - Timed Out/Unreachable - IP:" + Net.Dns.GetHostEntry(ip).AddressList.GetValue(0).ToString)
                        failedCount += 1
                    End If
                End If
            Catch ex As Exception
                ' In case of any exception (like DNS resolution failure), adds the IP with fail status to PingList.
                ' Increments the count of failed pings.
                PingList.Items.Add(ip + " - Fail - IP:Unavailable")
                failedCount += 1
            End Try
        Next

        ' Updates labels with the total, successful, and failed ping counts.
        TotalPingLbl.Text = "Total: " + PingList.Items.Count.ToString()
        SuccessfulPingLbl.Text = "Successful: " + successfulCount.ToString()
        FailedPingLbl.Text = "Failed: " + failedCount.ToString()

    End Sub


    ' Handles the click event of the RePingBtn button.
    Private Sub RePingBtn_Click(sender As Object, e As EventArgs) Handles RePingBtn.Click

        ' Clears the PingList of previous results.
        PingList.Items.Clear()

        ' Extracts text from PingTxt, removes carriage returns, and splits it into individual IP addresses.
        Dim ipList As String = PingTxt.Text
        Dim ips As String() = ipList.Replace(Chr(13), "").Split(Chr(10))

        ' Resets the counters for successful and failed ping attempts.
        successfulCount = 0
        failedCount = 0
        ' Resets the label displaying the count of selected IPs.
        PingSelectedLbl.Text = "Selected: 0"

        ' Iterates through each IP address.
        For Each ip As String In ips
            ip = ip.Replace(" ", "")
            Try
                ' Checks if the IP address string is valid (at least 2 characters long).
                If ip.Length >= 2 Then
                    ' Pings the IP address with a timeout of 100 milliseconds.
                    If My.Computer.Network.Ping(ip, 100) Then
                        ' If ping is successful, adds the IP with success status and its resolved address to PingList.
                        ' Increments the count of successful pings.
                        PingList.Items.Add(ip + " - Success - IP:" + Net.Dns.GetHostEntry(ip).AddressList.GetValue(0).ToString)
                        successfulCount += 1
                    Else
                        ' If ping fails (timeout or unreachable), adds the IP with failure status and its resolved address to PingList.
                        ' Increments the count of failed pings.
                        PingList.Items.Add(ip + " - Timed Out/Unreachable - IP:" + Net.Dns.GetHostEntry(ip).AddressList.GetValue(0).ToString)
                        failedCount += 1
                    End If
                End If
            Catch ex As Exception
                ' In case of any exception (like DNS resolution failure), adds the IP with fail status to PingList.
                ' Increments the count of failed pings.
                PingList.Items.Add(ip + " - Fail - IP:Unavailable")
                failedCount += 1
            End Try
        Next

        ' Updates labels with the total, successful, and failed ping counts.
        TotalPingLbl.Text = "Total: " + PingList.Items.Count.ToString()
        SuccessfulPingLbl.Text = "Successful: " + successfulCount.ToString()
        FailedPingLbl.Text = "Failed: " + failedCount.ToString()

    End Sub


    ' Handles the click event of the DiscoverSelectedBtn button.
    Private Sub DiscoverSelectedBtn_Click(sender As Object, e As EventArgs) Handles DiscoverSelectedBtn.Click

        ' Starts a new instance of the Chrome web browser.
        Process.Start("chrome")
        ' Pauses execution for 1 second to allow Chrome to open.
        Threading.Thread.Sleep(1000)

        ' Iterates through each selected item in the PingList control.
        For Each itm In PingList.SelectedItems
            ' Opens a new Chrome tab for each selected item, using a substring of the item's text as a parameter in the URL.
            ' Assumes that the relevant part of the itm string (device ID) is in the first 8 characters.
            Process.Start("chrome", "blank" + itm.Substring(0, 8))
            ' Pauses for 1 second between opening each tab.
            Threading.Thread.Sleep(1000)
        Next
        ' Additional pause after processing all selected items.
        Threading.Thread.Sleep(1000)
        ' Uses keyboard shortcuts to switch to the next tab and then close the current tab in Chrome.
        SendKeys.Send("^{TAB}")
        SendKeys.Send("^w")

    End Sub


    ' Handles the MouseDown event on the PingList control.
    Private Sub PingList_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PingList.MouseDown

        ' Checks if the right mouse button was clicked and if at most one item is selected in the PingList.
        If e.Button = MouseButtons.Right And PingList.SelectedItems.Count <= 1 Then
            ' Clears any currently selected items in the PingList.
            PingList.SelectedItems.Clear()
            ' Selects the item in the PingList where the right-click occurred.
            ' The IndexFromPoint method is used to get the index of the item at the mouse cursor's position.
            PingList.SelectedIndex = PingList.IndexFromPoint(e.X, e.Y)
        End If

    End Sub


    ' Handles the click event of the CopyRsMenuItem1 menu item.
    Private Sub CopyRsMenuItem1_Click(sender As Object, e As EventArgs) Handles CopyRsMenuItem1.Click

        ' Initializes a new list to store the selected items' text.
        Dim selectedItems As New List(Of String)

        ' Iterates through each selected item in the PingList.
        For Each item In PingList.SelectedItems
            ' Adds the first 8 characters (or less, if the string is shorter) of each selected item's text to the selectedItems list.
            ' The Math.Min function ensures that no more than the string's length is attempted to be taken, avoiding an error.
            selectedItems.Add(item.ToString().Substring(0, Math.Min(item.ToString().Length, 8)))
        Next

        ' Joins the list of selected items into a single string, separated by new lines.
        Dim result As String = String.Join(Environment.NewLine, selectedItems)

        ' Copies the resulting string to the clipboard.
        My.Computer.Clipboard.SetText(result)

    End Sub


    ' Handles the click event of the FailedPingsBtn button.
    Private Sub FailedPingsBtn_Click(sender As Object, e As EventArgs) Handles FailedPingsBtn.Click

        PingList.Items.Clear()

        ' Extracts text from PingTxt, removes carriage returns, and splits it into individual IP addresses.
        Dim ipList As String = PingTxt.Text
        Dim ips As String() = ipList.Replace(Chr(13), "").Split(Chr(10))

        ' Resets the count for failed ping attempts.
        failedCount = 0

        ' Iterates through each IP address in the list.
        For Each ip As String In ips
            ip = ip.Replace(" ", "")
            Try
                ' Checks if the IP address string is valid (at least 2 characters long).
                If ip.Length >= 2 Then
                    ' Pings the IP address with a timeout of 100 milliseconds.
                    If My.Computer.Network.Ping(ip, 100) Then
                        ' If ping is successful, nothing is added to the PingList (commented out code).
                    Else
                        ' If ping fails (timeout or unreachable), adds the IP with failure status and its resolved address to PingList.
                        ' Increments the count of failed pings.
                        PingList.Items.Add(ip + " - Timed Out/Unreachable - IP:" + Net.Dns.GetHostEntry(ip).AddressList.GetValue(0).ToString)
                        failedCount += 1
                    End If
                End If
            Catch ex As Exception
                ' In case of any exception (like DNS resolution failure), adds the IP with fail status to PingList.
                ' Increments the count of failed pings.
                PingList.Items.Add(ip + " - Fail - IP:Unavailable")
                failedCount += 1
            End Try
        Next

        ' Updates labels with the total, successful, and failed ping counts.
        TotalPingLbl.Text = "Total: " + PingList.Items.Count.ToString()
        SuccessfulPingLbl.Text = "Successful: " + successfulCount.ToString()
        FailedPingLbl.Text = "Failed: " + failedCount.ToString()

    End Sub


    ' Handles the click event of the GoodPingsBtn button.
    Private Sub GoodPingsBtn_Click(sender As Object, e As EventArgs) Handles GoodPingsBtn.Click

        PingList.Items.Clear()

        ' Extracts text from PingTxt, removes carriage returns, and splits it into individual IP addresses.
        Dim ipList As String = PingTxt.Text
        Dim ips As String() = ipList.Replace(Chr(13), "").Split(Chr(10))

        ' Resets the count for successful ping attempts.
        successfulCount = 0

        ' Iterates through each IP address in the list.
        For Each ip As String In ips
            ip = ip.Replace(" ", "")
            Try
                ' Checks if the IP address string is valid (at least 2 characters long).
                If ip.Length >= 2 Then
                    ' Pings the IP address with a timeout of 100 milliseconds.
                    If My.Computer.Network.Ping(ip, 100) Then
                        ' If ping is successful, adds the IP with success status and its resolved address to PingList.
                        ' Increments the count of successful pings.
                        PingList.Items.Add(ip + " - Success - IP:" + Net.Dns.GetHostEntry(ip).AddressList.GetValue(0).ToString)
                        successfulCount += 1
                    Else
                        ' If ping fails (timeout or unreachable), nothing is added to the PingList (commented out code).
                    End If
                End If
            Catch ex As Exception
                ' In case of any exception, nothing is added to the PingList (commented out code).
            End Try
        Next

        ' Updates labels with the total, successful, and failed ping counts.
        TotalPingLbl.Text = "Total: " + PingList.Items.Count.ToString()
        SuccessfulPingLbl.Text = "Successful: " + successfulCount.ToString()
        FailedPingLbl.Text = "Failed: " + failedCount.ToString()

    End Sub


    ' Handles the SelectedIndexChanged event of the PingList control.
    Private Sub PingList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles PingList.SelectedIndexChanged
        ' Counts the number of items currently selected in the PingList.
        Dim selectedCount As Integer = PingList.SelectedItems.Count
        ' Updates the PingSelectedLbl label with the count of selected items.
        PingSelectedLbl.Text = "Selected: " & selectedCount.ToString()
    End Sub


    ' Handles the click event of the ClearPingBtn button.
    Private Sub ClearPingBtn_Click(sender As Object, e As EventArgs) Handles ClearPingBtn.Click

        ' Clears the text from the PingTxt textbox.
        PingTxt.Clear()
        ' Clears all items from the PingList listbox.
        PingList.Items.Clear()
        ' Resets the counts for failed and successful ping attempts.
        failedCount = 0
        successfulCount = 0

        ' Updates the labels to reflect the reset state.
        ' TotalPingLbl shows the current total count of items in PingList (which should be 0 after clearing).
        TotalPingLbl.Text = "Total: " + PingList.Items.Count.ToString()
        ' SuccessfulPingLbl and FailedPingLbl show the counts of successful and failed pings, respectively.
        SuccessfulPingLbl.Text = "Successful: " + successfulCount.ToString()
        FailedPingLbl.Text = "Failed: " + failedCount.ToString()
        ' PingSelectedLbl is updated to show zero selected items.
        PingSelectedLbl.Text = "Selected: 0"

    End Sub


    ' Handles the click event of the GenCSVBtn button.
    Private Sub GenCSVBtn_Click(sender As Object, e As EventArgs) Handles GenCSVBtn.Click

        ' Gets the current date and time, and sets a start time 10 minutes from now.
        Dim now As DateTime = DateTime.Now
        Dim start As DateTime = now.AddMinutes(10)
        Console.WriteLine(start.ToString("MM/dd/yyyy HH:mm:ss"))

        ' Checks if both BackupOldTxt and BackupNewTxt textboxes have content.
        If BackupOldTxt.Text.Length & BackupNewTxt.Text.Length > 0 Then

            ' Processes the text from BackupOldTxt into an array of strings.
            Dim oldList As String = BackupOldTxt.Text
            Dim oldR() As String = oldList.Replace(Chr(13), "").Split(Chr(10))

            ' Processes the text from BackupNewTxt into an array of strings.
            Dim newList As String = BackupNewTxt.Text
            Dim newR() As String = newList.Replace(Chr(13), "").Split(Chr(10))

            ' Initializes the CSV string with headers.
            Dim csv As String = "HOSTNAME,BACKUPNAME,IPADDRESS,SCHEDULEDATE,SERVER,SELECTED" & vbCrLf

            ' Determines the smaller size between the two arrays.
            Dim minCount As Integer
            minCount = IIf(UBound(oldR) > UBound(newR), UBound(newR), UBound(oldR))

            ' Gets the current culture settings of the user.
            Dim userCulture As CultureInfo = CultureInfo.CurrentCulture

            ' Iterates through the arrays and constructs CSV entries.
            For i = 0 To minCount
                ' Every 5th iteration, increments the start time by 15 minutes.
                If i Mod 5 = 0 And i > 0 Then
                    start = start.AddMinutes(15)
                End If
                ' Adds an entry to the CSV string.
                csv = csv & oldR(i) & "," & newR(i) & ",,True," + start.ToString(userCulture) + "blank ,False" & vbCrLf
            Next

            ' Configures the SaveFileDialog to save a CSV file.
            SaveFileDialog1.DefaultExt = "csv"
            SaveFileDialog1.Filter = "CSV|*.csv"
            SaveFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)

            ' Opens the SaveFileDialog and writes the CSV string to the selected file.
            If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
                If Not String.IsNullOrEmpty(SaveFileDialog1.FileName) Then
                    System.IO.File.WriteAllText(SaveFileDialog1.FileName, csv)
                End If
            End If

        End If

    End Sub


    ' Handles the click event of the BackupCSVClearBtn button.
    Private Sub BackupCSVClearBtn_Click(sender As Object, e As EventArgs) Handles BackupCSVClearBtn.Click

        ' Clears the text from the BackupOldTxt textbox.
        BackupOldTxt.Clear()
        ' Clears the text from the BackupNewTxt textbox.
        BackupNewTxt.Clear()

    End Sub


    ' Handles the click event of the GenerateAvayaBtn button.
    Private Sub GenerateAvayaBtn_Click(sender As Object, e As EventArgs) Handles GenerateAvayaBtn.Click

        ' Splits the text from AvayaRsTxt into an array of computer names, removing carriage returns.
        Dim computerNames As String() = AvayaRsTxt.Text.Replace(Chr(13), "").Split(Chr(10))

        ' Iterates through each computer name in the array.
        For Each computerName As String In computerNames
            ' Calls GetLANByComputerName function for each computer name with "Avaya" as an argument.
            Dim username As String = GetLANByComputerName(computerName, "Avaya")
        Next

    End Sub


    Private Sub AvayaClearBtn_Click(sender As Object, e As EventArgs) Handles AvayaClearBtn.Click

        AvayaRsTxt.Clear()

    End Sub

    ' Handles the click event of the GenerateLaptopBtn button.
    Private Sub GenerateLaptopBtn_Click(sender As Object, e As EventArgs) Handles GenerateLaptopBtn.Click

        ' Splits the text from LaptopRsTxt into an array of computer names, removing carriage returns.
        Dim computerNames As String() = LaptopRsTxt.Text.Replace(Chr(13), "").Split(Chr(10))

        ' Iterates through each computer name in the array.
        For Each computerName As String In computerNames
            ' Calls the GetLANByComputerName function for each computer name with "Laptop" as an argument.
            Dim username As String = GetLANByComputerName(computerName, "Laptop")
        Next

    End Sub


    Private Sub LaptopClearBtn_Click(sender As Object, e As EventArgs) Handles LaptopClearBtn.Click

        LaptopRsTxt.Clear()

    End Sub

    ' Handles the click event of the GenerateSnBtn button.
    Private Sub GenerateSnBtn_Click(sender As Object, e As EventArgs) Handles GenerateSnBtn.Click

        ' Initializing lists to store old and new serial numbers or data.
        Dim oldsList As New List(Of String)
        Dim newsList As New List(Of String)

        ' Populating the lists from the textboxes, removing carriage returns and splitting by line feeds.
        oldsList = SnOldTxt.Text.Replace(Chr(13), "").Split(Chr(10)).ToList
        newsList = SnNewTxt.Text.Replace(Chr(13), "").Split(Chr(10)).ToList

        ' Ensuring both lists have the same number of elements by padding the shorter list with empty strings.
        If oldsList.Count < newsList.Count Then
            Dim diff As Integer = newsList.Count - oldsList.Count
            oldsList.AddRange(Enumerable.Repeat("", diff))
        ElseIf oldsList.Count > newsList.Count Then
            Dim diff As Integer = oldsList.Count - newsList.Count
            newsList.AddRange(Enumerable.Repeat("", diff))
        End If

        ' Iterating through the lists and adding rows to SnDataGrid.
        For i As Integer = 0 To oldsList.Count - 1
            ' Adds rows to the grid based on the presence of data in the old and new lists.
            If oldsList(i).Length > 0 And newsList(i).Length > 0 Then
                SnDataGrid.Rows.Add(oldsList(i) + " - " + newsList(i), "test")
            ElseIf oldsList(i).Length > 0 And newsList(i).Length < 1 Then
                SnDataGrid.Rows.Add(oldsList(i))
            ElseIf newsList(i).Length > 0 And oldsList(i).Length < 1 Then
                SnDataGrid.Rows.Add(newsList(i))
            End If
        Next

    End Sub


    ' Handles the CellMouseClick event of the SnDataGrid DataGridView.
    Private Sub SnDataGrid_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles SnDataGrid.CellMouseClick

        ' Checks if the right mouse button was clicked.
        If e.Button = MouseButtons.Right Then

            ' Checks if the background color of the clicked cell is already green.
            If SnDataGrid.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.BackColor = Color.Green Then
                ' If the cell is green, clears the background color to its default state.
                SnDataGrid.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.BackColor = Color.Empty
            Else
                ' If the cell is not green, changes the background color to green.
                SnDataGrid.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.BackColor = Color.Green
            End If

        End If

    End Sub


    ' Handles the CellDoubleClick event of the SnDataGrid DataGridView.
    Private Sub SnDataGrid_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles SnDataGrid.CellDoubleClick
        ' Sets a time interval to differentiate between double-click events.
        Dim doubleClickInterval As TimeSpan = TimeSpan.FromSeconds(3)

        ' Records the current time and calculates the time elapsed since the last double-click.
        Dim currentTime As DateTime = DateTime.Now
        Dim timeSinceLastDoubleClick As TimeSpan = currentTime - lastDoubleClickTime

        ' Checks if the time elapsed since the last double-click exceeds the set interval.
        If timeSinceLastDoubleClick >= doubleClickInterval Then
            ' Updates the time of the last double-click.
            lastDoubleClickTime = currentTime

            ' Initializes URL variables.
            Dim url As String
            Dim url2 As String = "blank"

            ' Determines the URL based on the state of the printerCheckSnUpdate checkbox.
            If printerCheckSnUpdate.Checked Then
                url = "blank"
            Else
                url = "blank"
            End If

            ' Processes the selected cell's value to extract R#s.
            If SnDataGrid.SelectedCells(0).Value IsNot Nothing AndAlso SnDataGrid.SelectedCells(0).Value.ToString().Contains(" - ") Then
                Dim rs As String = SnDataGrid.SelectedCells(0).Value.ToString()
                Dim rsList As String() = rs.Replace(" ", "").Split("-")
                Dim rList As New List(Of String)(rsList)

                ' Opens Chrome and navigates to specified URLs based on the extracted R#s.
                Process.Start("chrome")
                Threading.Thread.Sleep(1000)
                Process.Start("chrome", url + rList.Item(0))
                Threading.Thread.Sleep(1000)
                Process.Start("chrome", url2)
                Threading.Thread.Sleep(1000)
                ' Simulates keyboard shortcuts to manage browser tabs.
                SendKeys.Send("^{TAB}")
                SendKeys.Send("^w")
                Threading.Thread.Sleep(1000)

                ' Copies the second R#s to the clipboard.
                My.Computer.Clipboard.SetText(rList.Item(1))
            Else
                ' Handles the scenario when the selected cell's value does not contain a hyphen.
                Dim rs As String = SnDataGrid.SelectedCells(0).Value.ToString()

                ' Similar process of opening Chrome and navigating to URLs based on the single R#.
                Process.Start("chrome")
                Threading.Thread.Sleep(1000)
                Process.Start("chrome", url + rs)
                Threading.Thread.Sleep(1000)
                Process.Start("chrome", url2)
                Threading.Thread.Sleep(1000)
                SendKeys.Send("^{TAB}")
                SendKeys.Send("^w")
                Threading.Thread.Sleep(1000)

                ' Copies the R#s to the clipboard.
                My.Computer.Clipboard.SetText(rs)
            End If
        End If
    End Sub


    Private Sub OpenSnBtn_Click(sender As Object, e As EventArgs) Handles OpenSnBtn.Click

        Dim url As String
        Dim url2 As String = "blank"

        If printerCheckSnUpdate.Checked Then

            url = "blank"

        Else

            url = "blank"

        End If

        If SnDataGrid.SelectedCells(0).Value.ToString().Contains(" - ") Then

            Dim rs As String = SnDataGrid.SelectedCells(0).Value.ToString()
            Dim rsList As String() = rs.Replace(" ", "").Split("-")
            Dim rList As New List(Of String)

            For Each r As String In rsList
                rList.Add(r)
            Next

            Process.Start("chrome")
            Threading.Thread.Sleep(1000)
            Process.Start("chrome", url + rList.Item(0))
            Threading.Thread.Sleep(1000)
            Process.Start("chrome", url2)
            Threading.Thread.Sleep(1000)
            SendKeys.Send("^{TAB}")
            SendKeys.Send("^w")
            Threading.Thread.Sleep(1000)

            My.Computer.Clipboard.SetText(rList.Item(1))

        Else

            Dim rs As String = SnDataGrid.SelectedCells(0).Value.ToString()

            Process.Start("chrome")
            Threading.Thread.Sleep(1000)
            If printerCheckSnUpdate.Checked Then

                Process.Start("chrome", "blank" + rs.ToString)

            Else

                Process.Start("chrome", "blank" + rs.ToString)

            End If
            Threading.Thread.Sleep(1000)
            Process.Start("chrome", url2)
            Threading.Thread.Sleep(1000)
            SendKeys.Send("^{TAB}")
            SendKeys.Send("^w")
            Threading.Thread.Sleep(1000)

            My.Computer.Clipboard.SetText(rs.ToString)

        End If

    End Sub

    Private Sub ClearSnBtn_Click(sender As Object, e As EventArgs) Handles ClearSnBtn.Click

        SnDataGrid.Rows.Clear()
        SnOldTxt.Clear()
        SnNewTxt.Clear()

    End Sub

    Private Sub printLblBtn_Click(sender As Object, e As EventArgs) Handles printLblBtn.Click
        'Set up print document

        oldRsStr = oldRsLblTxt.Text.Replace(Chr(13), "").Split(Chr(10))
        newRsStr = newRsLblTxt.Text.Replace(Chr(13), "").Split(Chr(10))
        buildingsStr = buildingsLblTxt.Text.Replace(Chr(13), "").Split(Chr(10))
        floorsStr = floorsLblTxt.Text.Replace(Chr(13), "").Split(Chr(10))
        roomsStr = roomsLblTxt.Text.Replace(Chr(13), "").Split(Chr(10))

        control = 0

        If oldRsStr(oldRsStr.Length - 1) = "" Then

            Array.Resize(oldRsStr, oldRsStr.Length - 1)

        End If

        If newRsStr(newRsStr.Length - 1) = "" Then

            Array.Resize(newRsStr, newRsStr.Length - 1)

        End If

        If buildingsStr(buildingsStr.Length - 1) = "" Then

            Array.Resize(buildingsStr, buildingsStr.Length - 1)

        End If

        If floorsStr(floorsStr.Length - 1) = "" Then

            Array.Resize(floorsStr, floorsStr.Length - 1)

        End If

        If roomsStr(roomsStr.Length - 1) = "" Then

            Array.Resize(roomsStr, roomsStr.Length - 1)

        End If

        ' Find the maximum length among the string arrays
        Dim maxLength As Integer = Math.Max(oldRsStr.Length, Math.Max(newRsStr.Length, Math.Max(buildingsStr.Length, floorsStr.Length)))

        ' Resize each string array if its length is less than maxLength
        If oldRsStr.Length < maxLength Then
            Array.Resize(oldRsStr, maxLength)
            For i As Integer = oldRsStr.GetLowerBound(0) To oldRsStr.GetUpperBound(0)
                If oldRsStr(i) Is Nothing Then
                    oldRsStr(i) = "0"
                End If
            Next
        End If

        If newRsStr.Length < maxLength Then
            Array.Resize(newRsStr, maxLength)
            For i As Integer = newRsStr.GetLowerBound(0) To newRsStr.GetUpperBound(0)
                If newRsStr(i) Is Nothing Then
                    newRsStr(i) = "0"
                End If
            Next
        End If

        If buildingsStr.Length < maxLength Then
            Array.Resize(buildingsStr, maxLength)
            For i As Integer = buildingsStr.GetLowerBound(0) To buildingsStr.GetUpperBound(0)
                If buildingsStr(i) Is Nothing Then
                    If CheckBox2.Checked = True Then
                        buildingsStr(i) = buildingsStr(0)
                    Else
                        buildingsStr(i) = "0"
                    End If
                End If
            Next
        End If

        If floorsStr.Length < maxLength Then
            Array.Resize(floorsStr, maxLength)
            For i As Integer = floorsStr.GetLowerBound(0) To floorsStr.GetUpperBound(0)
                If floorsStr(i) Is Nothing Then
                    If CheckBox3.Checked = True Then
                        floorsStr(i) = floorsStr(0)
                    Else
                        floorsStr(i) = "0"
                    End If
                End If
            Next
        End If

        If roomsStr.Length < maxLength Then
            Array.Resize(roomsStr, maxLength)
            For i As Integer = roomsStr.GetLowerBound(0) To roomsStr.GetUpperBound(0)
                If roomsStr(i) Is Nothing Then
                    roomsStr(i) = "0"
                End If
            Next
        End If



        If oldRsStr.Length Mod 25 <> 0 Or oldRsStr.Length = 0 Then
            Array.Resize(oldRsStr, oldRsStr.Length + (25 - (oldRsStr.Length Mod 25))) 'If the length of strArray isn't a multiple of 25, resize it to add "0" to the end until it is
            For i As Integer = oldRsStr.GetLowerBound(0) To oldRsStr.GetUpperBound(0)
                If oldRsStr(i) Is Nothing Then
                    oldRsStr(i) = "0" 'Fill the newly added elements with "0"
                End If
            Next
        End If

        If newRsStr.Length Mod 25 <> 0 Or newRsStr.Length = 0 Then
            Array.Resize(newRsStr, newRsStr.Length + (25 - (newRsStr.Length Mod 25))) 'If the length of strArray isn't a multiple of 25, resize it to add "0" to the end until it is
            For i As Integer = newRsStr.GetLowerBound(0) To newRsStr.GetUpperBound(0)
                If newRsStr(i) Is Nothing Then
                    newRsStr(i) = "0" 'Fill the newly added elements with "0"
                End If
            Next
        End If

        If buildingsStr.Length Mod 25 <> 0 Or buildingsStr.Length = 0 Then
            Array.Resize(buildingsStr, buildingsStr.Length + (25 - (buildingsStr.Length Mod 25))) 'If the length of strArray isn't a multiple of 25, resize it to add "0" to the end until it is
            For i As Integer = buildingsStr.GetLowerBound(0) To buildingsStr.GetUpperBound(0)
                If buildingsStr(i) Is Nothing Then
                    If CheckBox2.Checked = True Then
                        buildingsStr(i) = buildingsStr(0)
                    Else
                        buildingsStr(i) = "0"
                    End If
                End If
            Next
        End If

        If floorsStr.Length Mod 25 <> 0 Or floorsStr.Length = 0 Then
            Array.Resize(floorsStr, floorsStr.Length + (25 - (floorsStr.Length Mod 25))) 'If the length of strArray isn't a multiple of 25, resize it to add "0" to the end until it is
            For i As Integer = floorsStr.GetLowerBound(0) To floorsStr.GetUpperBound(0)
                If floorsStr(i) Is Nothing Then
                    If CheckBox3.Checked = True Then
                        floorsStr(i) = floorsStr(0)
                    Else
                        floorsStr(i) = "0"
                    End If
                End If
            Next
        End If

        If roomsStr.Length Mod 25 <> 0 Or roomsStr.Length = 0 Then
            Array.Resize(roomsStr, roomsStr.Length + (25 - (roomsStr.Length Mod 25))) 'If the length of strArray isn't a multiple of 25, resize it to add "0" to the end until it is
            For i As Integer = roomsStr.GetLowerBound(0) To roomsStr.GetUpperBound(0)
                If roomsStr(i) Is Nothing Then
                    roomsStr(i) = "0"
                End If
            Next
        End If

        Dim lines As Integer = oldRsStr.Length()
        Dim result As Integer = lines / 25
        Dim rounded As Integer = lines Mod 25

        If result = 0 Then

            result = +1

        End If

        System.Console.WriteLine(result)

        If CheckBox1.Checked = True Then

            If result = 1 Then

                Dim pd As New Printing.PrintDocument()
                pd.DefaultPageSettings.Landscape = True
                AddHandler pd.PrintPage, AddressOf pd_PrintPage

                'Set up print preview dialog
                Dim ppd As New PrintPreviewDialog()
                ppd.Document = pd
                ppd.PrintPreviewControl.Zoom = 1.0

                'Show print preview dialog
                ppd.ShowDialog()

            Else

                For i As Integer = 0 To result - 1

                    control += 1

                    Dim pd As New Printing.PrintDocument()
                    pd.DefaultPageSettings.Landscape = True
                    AddHandler pd.PrintPage, AddressOf pd_PrintPage

                    'Set up print preview dialog
                    Dim ppd As New PrintPreviewDialog()
                    ppd.Document = pd
                    ppd.PrintPreviewControl.Zoom = 1.0

                    'Show print preview dialog
                    ppd.ShowDialog()
                Next

            End If

        Else

            If result = 1 Then

                Dim pd As New Printing.PrintDocument()
                pd.DefaultPageSettings.Landscape = True
                AddHandler pd.PrintPage, AddressOf pd_PrintPage

                'Print the document without preview
                pd.Print()

            Else

                For i As Integer = 0 To result - 1

                    control += 1

                    Dim pd As New Printing.PrintDocument()
                    pd.DefaultPageSettings.Landscape = True
                    AddHandler pd.PrintPage, AddressOf pd_PrintPage

                    'Print the document without preview
                    pd.Print()
                Next

            End If

        End If

    End Sub

    Private Sub clearPrintBtn_Click(sender As Object, e As EventArgs) Handles clearPrintBtn.Click
        oldRsLblTxt.Clear()
        newRsLblTxt.Clear()
        buildingsLblTxt.Clear()
        floorsLblTxt.Clear()
        roomsLblTxt.Clear()
    End Sub

    Private Sub autoAppBtn_Click(sender As Object, e As EventArgs) Handles autoAppBtn.Click

        actionDataList.Rows.Clear()
        noActionDataList.Rows.Clear()

        autoAppBtn.Enabled = False

        backgroundWorker.RunWorkerAsync()

    End Sub

    Private Sub moveActionToNoActionBtn_Click(sender As Object, e As EventArgs) Handles moveActionToNoActionBtn.Click

        Dim currentUser As String = System.Environment.UserName

        If currentUser = "blank" OrElse currentUser = "blank" Then

            If actionDataList.RowCount > 1 Then
                noActionDataList.Rows.Add(actionDataList.CurrentCell.Value.ToString())
            End If


            If actionDataList.CurrentCell IsNot Nothing And actionDataList.RowCount > 1 Then
                'Dim selectedCellValue As String = actionDataList.CurrentCell.Value.ToString()
                'Dim appDataPath As String = System.IO.Path.Combine(My.Computer.FileSystem.SpecialDirectories.MyDocuments, "MayoITRefreshTool")
                'Dim fileName As String = "NWWI_Action_Softwares.txt"
                'Dim filePathRemove As String = Path.Combine(appDataPath, fileName)
                'Dim fileName2 As String = "NWWI_NonAction_Softwares.txt"
                'Dim filePathAdd As String = Path.Combine(appDataPath, fileName2)

                Dim selectedCellValue As String = actionDataList.CurrentCell.Value.ToString()
                Dim appDataPath As String = System.IO.Path.Combine(My.Computer.FileSystem.SpecialDirectories.MyDocuments, "MayoITRefreshTool")
                Dim fileName As String = "NWWI_Action_Softwares.txt"
                Dim filePathRemove As String = Path.Combine("blank", fileName)
                Dim fileName2 As String = "NWWI_NonAction_Softwares.txt"
                Dim filePathAdd As String = Path.Combine("blank", fileName2)

                Try
                    ' Remove from file
                    Dim lines As List(Of String) = File.ReadAllLines(filePathRemove).ToList()
                    lines.RemoveAll(Function(line) line.Trim() = selectedCellValue)
                    File.WriteAllLines(filePathRemove, lines)

                    ' Add to another file
                    File.AppendAllText(filePathAdd, Environment.NewLine & selectedCellValue)

                    If actionDataList.SelectedCells.Count > 0 Then
                        Dim selectedRowIndex As Integer = actionDataList.SelectedCells(0).RowIndex
                        actionDataList.Rows.RemoveAt(selectedRowIndex)
                    End If
                Catch ex As Exception
                    MessageBox.Show("Error modifying the files: " & ex.Message)
                End Try
            End If
        Else
            MessageBox.Show("You are not authorized to perform this action.""" + currentUser + """")
        End If

    End Sub

    Private Sub moveNoActionToActionBtn_Click(sender As Object, e As EventArgs) Handles moveNoActionToActionBtn.Click

        Dim currentUser As String = System.Environment.UserName

        ' Check if the user is "M289928"
        If currentUser = "blank" OrElse currentUser = "blank" Then
            If noActionDataList.RowCount > 1 Then
                actionDataList.Rows.Add(noActionDataList.CurrentCell.Value.ToString())
            End If

            If noActionDataList.CurrentCell IsNot Nothing And noActionDataList.RowCount > 1 Then
                'Dim selectedCellValue As String = noActionDataList.CurrentCell.Value.ToString()
                'Dim appDataPath As String = System.IO.Path.Combine(My.Computer.FileSystem.SpecialDirectories.MyDocuments, "MayoITRefreshTool")
                'Dim fileName As String = "NWWI_Action_Softwares.txt"
                'Dim filePathAdd As String = Path.Combine(appDataPath, fileName)
                'Dim fileName2 As String = "NWWI_NonAction_Softwares.txt"
                'Dim filePathRemove As String = Path.Combine(appDataPath, fileName2)

                Dim selectedCellValue As String = noActionDataList.CurrentCell.Value.ToString()
                Dim appDataPath As String = System.IO.Path.Combine(My.Computer.FileSystem.SpecialDirectories.MyDocuments, "MayoITRefreshTool")
                Dim fileName As String = "NWWI_Action_Softwares.txt"
                Dim filePathAdd As String = Path.Combine("blank", fileName)
                Dim fileName2 As String = "NWWI_NonAction_Softwares.txt"
                Dim filePathRemove As String = Path.Combine("blank", fileName2)

                Try
                    ' Remove from file
                    Dim lines As List(Of String) = File.ReadAllLines(filePathRemove).ToList()
                    lines.RemoveAll(Function(line) line.Trim() = selectedCellValue)
                    File.WriteAllLines(filePathRemove, lines)

                    ' Add to another file
                    File.AppendAllText(filePathAdd, Environment.NewLine & selectedCellValue)
                    If noActionDataList.SelectedCells.Count > 0 Then
                        Dim selectedRowIndex As Integer = noActionDataList.SelectedCells(0).RowIndex
                        noActionDataList.Rows.RemoveAt(selectedRowIndex)
                    End If
                Catch ex As Exception
                    MessageBox.Show("Error modifying the files: " & ex.Message)
                End Try
            End If
        Else
            MessageBox.Show("You are not authorized to perform this action.""" + currentUser + """")
        End If

    End Sub

    Private Sub clearAutoAppBtn_Click(sender As Object, e As EventArgs) Handles clearAutoAppBtn.Click

        AutoAppRs.Clear()
        actionDataList.Rows.Clear()
        noActionDataList.Rows.Clear()

    End Sub

    Private Sub searchSNBtn_Click(sender As Object, e As EventArgs) Handles searchSNBtn.Click

        'Dim computerName As String = searchSNTxt.Text

        'resultsSNSearchTxt.Clear()

        'Using New ImpersonationHelper(domain, userName, password)
        '    Using connection As New SqlConnection(connectionString)
        '        Try
        '            connection.Open()

        '            Dim cmd As New SqlCommand("SELECT TOP 1 CIs.*, EI.*, SP.* FROM blank CIs INNER JOIN blank EI ON CIs.Device_ID = EI.Computer_Name INNER JOIN blank SP ON EI.Previous_Username = SP.LANID WHERE CIs.Device_ID = '" + computerName + "'", connection) ' Change YourTableName to your actual table name and modify the query as needed.
        '            Dim reader As SqlDataReader = cmd.ExecuteReader()



        '            While reader.Read()
        '                If reader("FirstName").ToString.Length() < 1 Then
        '                    resultsSNSearchTxt.AppendText("Computer Name: " + reader("Device_ID").ToString() + vbCrLf + "Last User: " + reader("Previous_Username") + vbCrLf + "Status: " + reader("Status").ToString() + vbCrLf + "Model: " + reader("Model").ToString() + vbCrLf + "Serial Number: " + reader("Serial_Number").ToString() + vbCrLf + "Location: " + reader("Location_Description").ToString + vbCrLf + "Location Notes: " + reader("Location_Notes").ToString() + vbCrLf + "Epic Department: " + reader("Epic_Department").ToString() + vbCrLf + "Epic User Type: " + reader("Epic_User_Type").ToString() + vbCrLf + "Created On: " + reader("Created_On").ToString()) ' Change YourColumnName to your actual column name.
        '                ElseIf reader("Previous_Username").ToString.Length() < 1 Then
        '                    resultsSNSearchTxt.AppendText("Computer Name: " + reader("Device_ID").ToString() + vbCrLf + "Last User: N/A" + vbCrLf + "Status: " + reader("Status").ToString() + vbCrLf + "Model: " + reader("Model").ToString() + vbCrLf + "Serial Number: " + reader("Serial_Number").ToString() + vbCrLf + "Location: " + reader("Location_Description").ToString + vbCrLf + "Location Notes: " + reader("Location_Notes").ToString() + vbCrLf + "Epic Department: " + reader("Epic_Department").ToString() + vbCrLf + "Epic User Type: " + reader("Epic_User_Type").ToString() + vbCrLf + "Created On: " + reader("Created_On").ToString())
        '                Else
        '                    resultsSNSearchTxt.AppendText("Computer Name: " + reader("Device_ID").ToString() + vbCrLf + "Last User: " + reader("Previous_Username") + " - " + reader("FirstName") + " " + reader("LastName") + vbCrLf + "Status: " + reader("Status").ToString() + vbCrLf + "Model: " + reader("Model").ToString() + vbCrLf + "Serial Number: " + reader("Serial_Number").ToString() + vbCrLf + "Location: " + reader("Location_Description").ToString + vbCrLf + "Location Notes: " + reader("Location_Notes").ToString() + vbCrLf + "Epic Department: " + reader("Epic_Department").ToString() + vbCrLf + "Epic User Type: " + reader("Epic_User_Type").ToString() + vbCrLf + "Created On: " + reader("Created_On").ToString()) ' Change YourColumnName to your actual column name.
        '                End If
        '            End While

        '            reader.Close()
        '        Catch ex As Exception
        '            MessageBox.Show("Error reading from database" + vbCrLf + "Please contact:" + vbCrLf + "Jacob Baker @ baker.jacob@mayo.edu" + vbCrLf + "Craig Zank @ zank.craig@mayo.edu")
        '        End Try
        '    End Using
        'End Using

        Dim computerName As String = searchSNTxt.Text

        resultsSNSearchTxt.Clear()
        CIAppsTxt.Clear()
        DevicesTxt.Clear()

        Using New ImpersonationHelper(domain, userName, password)
            Using connection As New SqlConnection(connectionString)
                Try
                    connection.Open()

                    ' Fetch data from EndpointInventory
                    Dim cmdEI As New SqlCommand("SELECT TOP 1 * FROM blank WHERE Computer_Name = @computerName", connection)
                    cmdEI.Parameters.AddWithValue("@computerName", computerName)
                    Dim readerEI As SqlDataReader = cmdEI.ExecuteReader()

                    Dim previousUsername As String = String.Empty

                    If readerEI.Read() Then
                        previousUsername = readerEI("Previous_Username").ToString()
                    End If
                    readerEI.Close()

                    ' Fetch data from ServiceNowCIs
                    Dim cmdCIs As New SqlCommand("SELECT TOP 1 * FROM blank WHERE Device_ID = @computerName", connection)
                    cmdCIs.Parameters.AddWithValue("@computerName", computerName)
                    Dim readerCIs As SqlDataReader = cmdCIs.ExecuteReader()

                    Dim deviceID As String = String.Empty
                    Dim status As String = String.Empty
                    Dim model As String = String.Empty
                    Dim serialNumber As String = String.Empty
                    Dim locationDescription As String = String.Empty
                    Dim locationNotes As String = String.Empty
                    Dim epicDepartment As String = String.Empty
                    Dim epicUserType As String = String.Empty
                    Dim createdOn As String = String.Empty

                    If readerCIs.Read() Then
                        deviceID = readerCIs("Device_ID").ToString()
                        status = readerCIs("Status").ToString()
                        model = readerCIs("Model").ToString()
                        serialNumber = readerCIs("Serial_Number").ToString()
                        locationDescription = readerCIs("Location_Description").ToString()
                        locationNotes = readerCIs("Location_Notes").ToString()
                        epicDepartment = readerCIs("Epic_Department").ToString()
                        epicUserType = readerCIs("Epic_User_Type").ToString()
                        createdOn = readerCIs("Created_On").ToString()
                    End If
                    readerCIs.Close()

                    ' Fetch data from ServiceNowPerson
                    Dim firstName As String = String.Empty
                    Dim lastName As String = String.Empty
                    If Not String.IsNullOrEmpty(previousUsername) Then
                        Dim cmdSP As New SqlCommand("SELECT TOP 1 * FROM blank WHERE LANID = @previousUsername", connection)
                        cmdSP.Parameters.AddWithValue("@previousUsername", previousUsername)
                        Dim readerSP As SqlDataReader = cmdSP.ExecuteReader()

                        If readerSP.Read() Then
                            firstName = readerSP("FirstName").ToString()
                            lastName = readerSP("LastName").ToString()
                        End If
                        readerSP.Close()
                    End If

                    Dim cmdApps As New SqlCommand("SELECT Application_Name FROM blank ea WHERE Computer_Name = @computerName", connection)
                    cmdApps.Parameters.AddWithValue("@computerName", computerName)
                    Dim readerApps As SqlDataReader = cmdApps.ExecuteReader()

                    While readerApps.Read()
                        CIAppsTxt.AppendText(readerApps("Application_Name").ToString() + vbCrLf)
                    End While
                    readerApps.Close()

                    Dim cmdDevices As New SqlCommand("SELECT USB_Devices FROM blank ea WHERE Computer_Name = @computerName", connection)
                    cmdDevices.Parameters.AddWithValue("@computerName", computerName)
                    Dim readerDevices As SqlDataReader = cmdDevices.ExecuteReader()

                    While readerDevices.Read()
                        Dim stringWithSpaces As String = readerDevices("USB_Devices").ToString().Replace(vbCrLf, " ").Replace(vbLf, " ").Replace(vbCr, " ").Replace(",", vbCrLf)
                        DevicesTxt.AppendText(stringWithSpaces)
                    End While
                    readerDevices.Close()

                    ' Append data to resultsSNSearchTxt
                    resultsSNSearchTxt.AppendText($"Computer Name: {deviceID}{vbCrLf}Last User: ")
                    If String.IsNullOrEmpty(previousUsername) Then
                        resultsSNSearchTxt.AppendText("N/A")
                    ElseIf String.IsNullOrEmpty(firstName) And String.IsNullOrEmpty(lastName) Then
                        resultsSNSearchTxt.AppendText(previousUsername)
                    Else
                        resultsSNSearchTxt.AppendText($"{previousUsername} - {firstName} {lastName}")
                    End If
                    resultsSNSearchTxt.AppendText($"{vbCrLf}Status: {status}{vbCrLf}Model: {model}{vbCrLf}Serial Number: {serialNumber}{vbCrLf}Location: {locationDescription}{vbCrLf}Location Notes: {locationNotes}{vbCrLf}Epic Department: {epicDepartment}{vbCrLf}Epic User Type: {epicUserType}{vbCrLf}Created On: {createdOn}")

                Catch ex As Exception
                    MessageBox.Show("Error reading from database" + vbCrLf + "Please contact:" + vbCrLf + "Jacob Baker @ baker.jacob@mayo.edu" + vbCrLf + "Craig Zank @ zank.craig@mayo.edu")
                End Try
            End Using
        End Using

    End Sub

    Private Sub peopleSearchBtn_Click(sender As Object, e As EventArgs) Handles peopleSearchBtn.Click

        Dim lanID As String = peopleSearchTxt.Text

        peopleResultsTxt.Clear()

        Using New ImpersonationHelper(domain, userName, password)
            Using connection As New SqlConnection(connectionString)
                Try
                    connection.Open()

                    ' Fetch data from EndpointInventory
                    Dim cmdEI As New SqlCommand("SELECT TOP 1 * FROM blank WHERE LANID = @lanID", connection)
                    cmdEI.Parameters.AddWithValue("@lanID", lanID)
                    Dim readerEI As SqlDataReader = cmdEI.ExecuteReader()

                    Dim previousUsername As String = String.Empty

                    If readerEI.Read() Then

                        peopleResultsTxt.AppendText("Name: " + readerEI("FirstName") + " " + readerEI("LastName") + vbCrLf + "LAN ID: " + readerEI("LANID") + vbCrLf + "Email: " + readerEI("Email") + vbCrLf + "Job Title: " + readerEI("Job_title") + vbCrLf + "Manager Name: " + readerEI("ManagerName") + vbCrLf + "Reporting Unit: " + readerEI("ReportingUnit"))

                    End If
                    readerEI.Close()

                Catch ex As Exception
                    MessageBox.Show("Error reading from database" + vbCrLf + "Please contact:" + vbCrLf + "Jacob Baker @ baker.jacob@mayo.edu" + vbCrLf + "Craig Zank @ zank.craig@mayo.edu")
                End Try
            End Using
        End Using

    End Sub

    Private Sub SetDefaultText(tb As TextBox, defaultText As String)
        tb.Text = defaultText
        tb.ForeColor = Color.Gray
        AddHandler tb.GotFocus, Sub(sender As Object, e As EventArgs)
                                    If tb.Text = defaultText Then
                                        tb.Text = ""
                                        tb.ForeColor = Color.Black
                                    End If
                                End Sub

        AddHandler tb.LostFocus, Sub(sender As Object, e As EventArgs)
                                     If String.IsNullOrWhiteSpace(tb.Text) Then
                                         tb.Text = defaultText
                                         tb.ForeColor = Color.Gray
                                     End If
                                 End Sub
    End Sub

    Private Function DataGridViewContains(dgv As DataGridView, app As String) As Boolean
        For Each row As DataGridViewRow In dgv.Rows
            If row.Cells(0).Value IsNot Nothing AndAlso row.Cells(0).Value.ToString() = app Then
                Return True
            End If
        Next
        Return False
    End Function

    Private Sub SortDataList(dgv As DataGridView)
        If dgv.InvokeRequired Then
            dgv.Invoke(New Action(Of DataGridView)(AddressOf SortDataList), dgv)
        Else
            Dim sortedList = dgv.Rows.Cast(Of DataGridViewRow)().
            Where(Function(row) Not row.IsNewRow).
            OrderBy(Function(row) row.Cells(0).Value.ToString(), StringComparer.OrdinalIgnoreCase).ToList()

            dgv.Rows.Clear()

            For Each row As DataGridViewRow In sortedList
                dgv.Rows.Add(row.Cells(0).Value.ToString())
            Next
        End If
    End Sub

    Private Function GetApplicationNamesByComputerName(computerName As String) As String

        'Dim appDataPath As String = System.IO.Path.Combine(My.Computer.FileSystem.SpecialDirectories.MyDocuments, "MayoITRefreshTool")
        'Dim fileName As String = "NWWI_Action_Softwares.txt"
        'Dim filePath As String = Path.Combine(appDataPath, fileName)
        'Dim fileName2 As String = "NWWI_NonAction_Softwares.txt"
        'Dim filePath2 As String = Path.Combine(appDataPath, fileName2)

        Dim appDataPath As String = System.IO.Path.Combine(My.Computer.FileSystem.SpecialDirectories.MyDocuments, "MayoITRefreshTool")
        Dim fileName As String = "NWWI_Action_Softwares.txt"
        Dim filePath As String = Path.Combine("blank", fileName)
        Dim fileName2 As String = "NWWI_NonAction_Softwares.txt"
        Dim filePath2 As String = Path.Combine("blank", fileName2)

        Dim presetActionApplicationNames As New List(Of String)()
        Dim presetNonActionApplicationNames As New List(Of String)()

        Try
            presetActionApplicationNames = File.ReadAllLines(filePath).ToList()
            presetNonActionApplicationNames = File.ReadAllLines(filePath2).ToList()
        Catch ex As Exception
            MessageBox.Show("Error reading the file: " & ex.Message)
            Return String.Empty 'If we can't read the files, it is better to exit the function.
        End Try

        Dim addActionApplicationNames As New List(Of String)()
        Dim addActionApplicationNamesFile As New List(Of String)()
        Dim addNonActionApplicationNames As New List(Of String)()
        Dim addNonActionApplicationNamesFile As New List(Of String)()
        Dim applicationNames As New List(Of String)()

        Using Impersonator As New ImpersonationHelper(domain, userName, password)
            Using connection As New SqlConnection(connectionString)
                Try
                    connection.Open()
                    Dim query As String = "SELECT Application_Name FROM blank WHERE Computer_Name = @ComputerName"
                    Using command As New SqlCommand(query, connection)
                        command.Parameters.AddWithValue("@ComputerName", computerName) 'Avoid SQL Injection

                        Using reader As SqlDataReader = command.ExecuteReader()
                            While reader.Read()
                                Dim appName As String = reader("Application_Name").ToString()

                                If presetActionApplicationNames.Contains(appName) Then

                                    ProcessActionApplicationNames(appName, addActionApplicationNames, applicationNames)

                                ElseIf presetNonActionApplicationNames.Contains(appName) Then

                                    addNonActionApplicationNames.Add(appName)

                                Else

                                    ProcessActionApplicationNames(appName, addActionApplicationNames, applicationNames)

                                End If
                            End While
                        End Using
                    End Using
                Catch ex As Exception
                    MessageBox.Show("Error connecting to SQL Server: " & ex.Message)
                Finally
                    If connection.State = ConnectionState.Open Then
                        connection.Close() 'Ensure the connection is closed even if an error occurred
                    End If
                End Try
            End Using
        End Using


        For Each app As String In addActionApplicationNames
            AddToActionDataList(app)
        Next
        SortDataList(actionDataList)

        For Each app As String In addNonActionApplicationNames
            AddToNoActionDataList(app)
        Next
        SortDataList(noActionDataList)


        HandleNewApplicationName()

        Return String.Join(", ", applicationNames)

    End Function

    Private Function GetApplicationNamesByComputerNameCSV(computerName As String) As String

        Dim appDataPath As String = System.IO.Path.Combine(My.Computer.FileSystem.SpecialDirectories.MyDocuments, "MayoITRefreshTool")
        Dim fileName As String = "NWWI_Action_Softwares.txt"
        Dim filePath As String = Path.Combine("blank", fileName)
        Dim fileName2 As String = "NWWI_NonAction_Softwares.txt"
        Dim filePath2 As String = Path.Combine("blank", fileName2)

        Dim presetActionApplicationNames As New List(Of String)()
        Dim presetNonActionApplicationNames As New List(Of String)()

        Try
            presetActionApplicationNames = File.ReadAllLines(filePath).ToList()
            presetNonActionApplicationNames = File.ReadAllLines(filePath2).ToList()
        Catch ex As Exception
            MessageBox.Show("Error reading the file: " & ex.Message)
            Return String.Empty 'If we can't read the files, it is better to exit the function.
        End Try

        Dim addActionApplicationNames As New List(Of String)()
        Dim addActionApplicationNamesFile As New List(Of String)()
        Dim addNonActionApplicationNames As New List(Of String)()
        Dim addNonActionApplicationNamesFile As New List(Of String)()
        Dim applicationNames As New List(Of String)()

        Using Impersonator As New ImpersonationHelper(domain, userName, password)
            Using connection As New SqlConnection(connectionString)
                Try
                    connection.Open()
                    Dim query As String = "SELECT Application_Name FROM blank WHERE Computer_Name = @ComputerName"
                    Using command As New SqlCommand(query, connection)
                        command.Parameters.AddWithValue("@ComputerName", computerName) 'Avoid SQL Injection

                        Using reader As SqlDataReader = command.ExecuteReader()
                            While reader.Read()
                                Dim appName As String = reader("Application_Name").ToString()

                                If presetActionApplicationNames.Contains(appName) Then

                                    ProcessActionApplicationNames(appName, addActionApplicationNames, applicationNames)

                                ElseIf presetNonActionApplicationNames.Contains(appName) Then

                                    addNonActionApplicationNames.Add(appName)

                                Else

                                    ProcessActionApplicationNames(appName, addActionApplicationNames, applicationNames)

                                End If
                            End While
                        End Using
                    End Using
                Catch ex As Exception
                    MessageBox.Show("Error connecting to SQL Server: " & ex.Message)
                Finally
                    If connection.State = ConnectionState.Open Then
                        connection.Close() 'Ensure the connection is closed even if an error occurred
                    End If
                End Try
            End Using
        End Using

        If noActionDataList.RowCount < 3 Then
            For Each app As String In addActionApplicationNames
                AddToActionDataList(app)
            Next
            SortDataList(actionDataList)

            For Each app As String In addNonActionApplicationNames
                AddToNoActionDataList(app)
            Next
            SortDataList(noActionDataList)


            HandleNewApplicationName()
        End If

        Return String.Join(", ", applicationNames)

    End Function

    Private Sub HandleNewApplicationName()

        'Dim appDataPath As String = System.IO.Path.Combine(My.Computer.FileSystem.SpecialDirectories.MyDocuments, "MayoITRefreshTool")
        'Dim fileName As String = "NWWI_Action_Softwares.txt"
        'Dim filePath As String = Path.Combine(appDataPath, fileName)

        Dim appDataPath As String = System.IO.Path.Combine(My.Computer.FileSystem.SpecialDirectories.MyDocuments, "MayoITRefreshTool")
        Dim fileName As String = "NWWI_Action_Softwares.txt"
        Dim filePath As String = Path.Combine("blank", fileName)

        Dim list As New List(Of String)()
        Dim presetActionApplicationNames As List(Of String) = File.ReadAllLines(filePath).ToList()

        ' Iterate through the rows of the DataGridView.
        For Each row As DataGridViewRow In actionDataList.Rows
            ' Iterate through the cells of the DataGridViewRow.
            For Each cell As DataGridViewCell In row.Cells
                ' Add the cell's value to the list.
                If Not cell.Value Is Nothing Then ' To avoid null value error
                    list.Add(Convert.ToString(cell.Value))
                End If
            Next
        Next

        For Each app In list

            If Not presetActionApplicationNames.Contains(app) Then

                File.AppendAllText(filePath, Environment.NewLine & app)

            End If

        Next

    End Sub

    Private Sub ProcessActionApplicationNames(appName As String, addActionApplicationNames As List(Of String), applicationNames As List(Of String))

        addActionApplicationNames.Add(appName)

        Dim path As String = "blankShort_Names.txt"
        Dim applicationsToCheck As New List(Of String)

        Try
            applicationsToCheck = File.ReadAllLines(path).ToList()
        Catch ex As Exception
            MessageBox.Show("An error occurred: " & ex.Message)
            Return
        End Try

        'Dim applicationsToCheck As New List(Of String)() From {"SoftID", "TEG", "Avaya", "BCA", "Genetec", "PaymentMate", "CPR+", "Heidelberg", "Varian", "CaptureOnTouch", "CapturePerfect", "Calabrio ONE", "Aperio"}

        Dim appFound As Boolean = False

        ' Special case for Milestone
        If appName.Contains("Milestone") Then
            If Not applicationNames.Contains("Genetec") Then
                applicationNames.Add("Genetec")
            End If
            appFound = True
        End If

        If Not appFound Then
            For Each application In applicationsToCheck
                If appName.Contains(application) Then
                    If Not applicationNames.Contains(application) Then
                        applicationNames.Add(application)
                    End If
                    appFound = True
                    Exit For
                End If
            Next
        End If

        If Not appFound Then
            applicationNames.Add(appName)
        End If

    End Sub

    Private Sub backgroundWorker_DoWork(sender As Object, e As DoWorkEventArgs) Handles backgroundWorker.DoWork
        ' Simulate a long-running operation (e.g., processing large data)
        Dim computerNames As String() = Nothing

        ' Execute the UI-related code on the UI thread using Invoke
        Me.Invoke(Sub()
                      computerNames = AutoAppRs.Text.Replace(Chr(13), "").Split(Chr(10))
                  End Sub)

        If autoAppBtn.Enabled = False Then

            For Each computerName As String In computerNames
                Dim applicationNames As String = GetApplicationNamesByComputerName(computerName)
            Next

        ElseIf appCSVGenerateBtn.Enabled = False Then

            Dim csvFilePath As String = String.Empty

            ' Allow user to choose the save location using SaveFileDialog
            Dim saveFileDialogResult As DialogResult = Me.Invoke(Function()
                                                                     Dim saveFileDialog As New SaveFileDialog()
                                                                     saveFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
                                                                     saveFileDialog.Title = "Save as CSV"
                                                                     If saveFileDialog.ShowDialog() = DialogResult.OK Then
                                                                         csvFilePath = saveFileDialog.FileName
                                                                         Return DialogResult.OK
                                                                     Else
                                                                         Return DialogResult.Cancel
                                                                     End If
                                                                 End Function)

            If saveFileDialogResult = DialogResult.Cancel Then
                Return
            End If

            Try
                Using writer As New StreamWriter(csvFilePath)
                    For Each computerName As String In computerNames
                        ' Get the applications for the current computer name
                        Dim applicationNames As String = GetApplicationNamesByComputerName(computerName)

                        ' Write the results to a CSV file with each computer's applications on a new line
                        writer.WriteLine(computerName & ",""" & applicationNames & """")
                    Next
                End Using

                ' Notify the user on the UI thread that the CSV file is created successfully
                Me.Invoke(Sub()
                              MessageBox.Show("CSV file created successfully.")
                          End Sub)
            Catch ex As IOException
                ' File is in use, prompt the user to close it or choose an action
                Dim result As DialogResult = Me.Invoke(Function()
                                                           Return MessageBox.Show("The file is currently in use by another process. Do you want to try again?", "File In Use", MessageBoxButtons.RetryCancel, MessageBoxIcon.Warning)
                                                       End Function)
                If result = DialogResult.Retry Then
                    backgroundWorker_DoWork(sender, e) ' Recursive call to retry the saving
                End If

            Catch ex As Exception
                ' Handle other exceptions (e.g., file access permission issues, etc.)
                Me.Invoke(Sub()
                              MessageBox.Show("An error occurred while processing the data: " & ex.Message)
                          End Sub)
            End Try
        End If

        'Sleep for a short duration to simulate work
        Thread.Sleep(100)
    End Sub

    Private Sub backgroundWorker_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles backgroundWorker.RunWorkerCompleted
        ' Operation is complete, update the UI on the main thread
        Me.Invoke(Sub()
                      ' Enable the button again
                      autoAppBtn.Enabled = True
                      appCSVGenerateBtn.Enabled = True
                  End Sub)
    End Sub

    Private Class ImpersonationHelper
        Implements IDisposable

        Private impersonationContext As WindowsImpersonationContext

        Public Sub New(domain As String, userName As String, password As String)
            Dim token As IntPtr = IntPtr.Zero
            Dim tokenDuplicate As IntPtr = IntPtr.Zero

            If LogonUser(userName, domain, password, 2, 0, token) Then
                If DuplicateToken(token, 2, tokenDuplicate) Then
                    Dim tempWindowsIdentity As New WindowsIdentity(tokenDuplicate)
                    impersonationContext = tempWindowsIdentity.Impersonate()
                Else
                    Throw New Exception("DuplicateToken failed.")
                End If
            Else
                Throw New Exception("LogonUser failed.")
            End If

            If Not tokenDuplicate.Equals(IntPtr.Zero) Then
                CloseHandle(tokenDuplicate)
            End If
            If Not token.Equals(IntPtr.Zero) Then
                CloseHandle(token)
            End If
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            If impersonationContext IsNot Nothing Then
                impersonationContext.Undo()
            End If
        End Sub

        Private Declare Auto Function LogonUser Lib "advapi32.dll" (lpszUsername As String, lpszDomain As String, lpszPassword As String, dwLogonType As Integer, dwLogonProvider As Integer, ByRef phToken As IntPtr) As Boolean
        Private Declare Auto Function DuplicateToken Lib "advapi32.dll" (ExistingTokenHandle As IntPtr, ImpersonationLevel As Integer, ByRef DuplicateTokenHandle As IntPtr) As Boolean
        Private Declare Auto Function CloseHandle Lib "kernel32.dll" (handle As IntPtr) As Boolean

    End Class

    Function GetOrdinalSuffix(day As Integer) As String
        If day Mod 100 >= 11 AndAlso day Mod 100 <= 13 Then
            Return "th"
        End If
        Select Case day Mod 10
            Case 1
                Return "st"
            Case 2
                Return "nd"
            Case 3
                Return "rd"
            Case Else
                Return "th"
        End Select
    End Function

    Private Function GetLANByComputerName(computerName As String, type As String) As String
        Dim currentUser As String = System.Environment.UserName

        Using Impersonator As New ImpersonationHelper(domain, userName, password)
            ' Your SQL Server connection code goes here
            ' For example:
            Dim query As String = "SELECT Previous_Username FROM blank WHERE Computer_Name = '" + computerName + "'"

            Dim resultBuilder As New StringBuilder()
            Using connection As New SqlConnection(connectionString)
                Try
                    connection.Open()
                    Using command As New SqlCommand(query, connection)
                        Using reader As SqlDataReader = command.ExecuteReader()
                            Try
                                While reader.Read()
                                    Dim fileName As String = "UserIDs.txt"
                                    Dim filePath As String = Path.Combine("blank", fileName)
                                    Dim lineCount = File.ReadAllLines(filePath).Length
                                    Dim fileName2 As String = "EmailIDs.txt"
                                    Dim filePath2 As String = Path.Combine("blank", fileName2)
                                    Dim lineCount2 = File.ReadAllLines(filePath2).Length

                                    Dim UserIDs As String
                                    Dim UserIDsList As String() = File.ReadAllLines(filePath)

                                    Dim EmailIDs As String
                                    Dim EmailIDsList As String() = File.ReadAllLines(filePath2)

                                    Dim foundId As String = String.Empty

                                    File.AppendAllText(filePath2, Environment.NewLine + computerName + ", EID:" + lineCount2.ToString() + ", White")

                                    For Each item As String In UserIDsList
                                        If item.StartsWith(currentUser) Then
                                            Dim parts As String() = item.Split(" "c)
                                            foundId = parts(parts.Length - 1)
                                        End If
                                    Next

                                    Try
                                        UserIDs = File.ReadAllText(filePath)
                                        If Not UserIDs.Contains(currentUser) Then
                                            File.AppendAllText(filePath, Environment.NewLine + currentUser + " - ID:" + lineCount.ToString())
                                        End If
                                    Catch ex As Exception
                                        MessageBox.Show("Error reading the file: " & ex.Message)
                                        Return String.Empty
                                    End Try

                                    Dim username As String = reader("Previous_Username").ToString()
                                    Dim compliance As String = ""

                                    'Create an instance of Outlook
                                    Dim outlookObj As Outlook.Application = New Outlook.Application()

                                    ' Create an instance of MailItem
                                    Dim mailItem As Outlook.MailItem = DirectCast(outlookObj.CreateItem(Outlook.OlItemType.olMailItem), Outlook.MailItem)

                                    ' Set the recipient
                                    mailItem.To = username

                                    Dim name As String = GetNameFromEmail(username)
                                    Dim separator As String = " - "

                                    For Each email As String In CheckedListBox1.CheckedItems
                                        Dim index As Integer = email.IndexOf(separator)
                                        If index <> -1 Then
                                            ' Extract the substring after the separator
                                            Dim extractedText As String = email.Substring(index + separator.Length)
                                            mailItem.ReplyRecipients.Add(extractedText)
                                            Dim ccRecipient As Outlook.Recipient = mailItem.Recipients.Add(extractedText)
                                            ccRecipient.Type = Outlook.OlMailRecipientType.olCC
                                        Else
                                            Console.WriteLine("Separator not found in the string.")
                                        End If

                                    Next

                                    mailItem.Recipients.ResolveAll()

                                    Dim selectedDate As Date = MonthCalendar1.SelectionRange.Start
                                    Dim daysUntilNextMonday As Integer = ((DayOfWeek.Monday - selectedDate.DayOfWeek) + 7) Mod 7
                                    If daysUntilNextMonday = 0 Then daysUntilNextMonday = 7
                                    Dim nextMonday As Date = selectedDate.AddDays(daysUntilNextMonday)

                                    Dim dayWithSuffix1 As String = nextMonday.Day.ToString() & GetOrdinalSuffix(nextMonday.Day)
                                    Dim formattedDate1 As String = nextMonday.ToString("MMMM ") & dayWithSuffix1

                                    Dim dayWithSuffix2 As String = selectedDate.Day.ToString() & GetOrdinalSuffix(selectedDate.Day)
                                    Dim formattedDate2 As String = selectedDate.ToString("MMMM ") & dayWithSuffix2

                                    ' Set the subject and load HTML body from file
                                    If type = "Avaya" Then
                                        mailItem.Subject = "PC Refresh for " + computerName + " - Avaya"
                                        mailItem.HTMLBody = File.ReadAllText("blankAvayaEmailTemplate.txt")
                                    ElseIf type = "Laptop" Then
                                        mailItem.Subject = "Laptop Refresh for " + computerName
                                        If CheckBox4.Checked = True Then
                                            mailItem.Subject += "                                                                                                                                                                                                        Comp:(P" + foundId + ")(EID:" + lineCount2.ToString() + ")"
                                        End If
                                        mailItem.HTMLBody = File.ReadAllText("blankLaptopEmailTemplate.txt")
                                    End If

                                    If CheckBox4.Checked = True Then
                                        compliance = "<p><i><b><span class=""redc"">Please contact us as soon as possible. In order for your laptop to be replaced in a timely manner, your response is requested by " + formattedDate2 + ". Failure to respond by this deadline will result in your laptop being bumped to a later refresh cycle. We appreciate your understanding and cooperation in this matter.<br><br>Please answer the following questions in your response, and we will then coordinate with you to arrange a refresh date during the week of " + formattedDate1 + ". If you are unavailable that week, there is no need to respond as you will receive another notice at a later date.</span></b></i></p>"
                                    End If

                                    ' Replace placeholders in HTML body
                                    mailItem.HTMLBody = mailItem.HTMLBody.Replace("{name}", name).Replace("{computerName}", computerName).Replace("{formattedDate2}", formattedDate2).Replace("{formattedDate1}", formattedDate1).Replace("{compliance}", compliance)

                                    mailItem.Display(mailItem)

                                End While
                            Catch ex As Exception
                                Console.WriteLine("Error during email setup: " & ex.Message)
                            End Try
                        End Using
                    End Using
                Catch ex As Exception
                    MessageBox.Show("Error connecting to SQL Server: " & ex.Message)
                End Try
            End Using
        End Using

        Return "Email prepared and displayed successfully."
    End Function


    Public Function GetNameFromEmail(emailAddress As String) As String
        Dim outlookApp As New Outlook.Application()
        Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
        Dim recipient As Outlook.Recipient = Nothing
        Dim addressEntry As Outlook.AddressEntry = Nothing
        Dim firstName As String = ""

        Try
            ' Try to resolve the email address
            recipient = ns.CreateRecipient(emailAddress)
            If recipient IsNot Nothing Then
                recipient.Resolve()

                If recipient.Resolved Then
                    addressEntry = recipient.AddressEntry
                    If addressEntry IsNot Nothing Then
                        ' Get the name from the address entry
                        Dim fullName As String = addressEntry.Name
                        If Not String.IsNullOrEmpty(fullName) Then
                            ' Extract the first name from the full name
                            Dim commaIndex As Integer = fullName.IndexOf(","c)
                            Dim spaceAfterComma As Integer = fullName.IndexOf(" "c, commaIndex + 2)

                            If commaIndex >= 0 AndAlso spaceAfterComma > commaIndex + 2 AndAlso spaceAfterComma < fullName.Length - 1 Then
                                firstName = fullName.Substring(commaIndex + 2, spaceAfterComma - commaIndex - 2).Trim()
                                Return firstName
                            End If
                        End If
                    Else
                        MessageBox.Show("No address entry found for the email address.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                Else
                    MessageBox.Show("Unable to resolve the email address.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            Else
                MessageBox.Show("Recipient is null.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            MessageBox.Show("An error occurred: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' Release COM objects
            If addressEntry IsNot Nothing Then Marshal.ReleaseComObject(addressEntry)
            If recipient IsNot Nothing Then Marshal.ReleaseComObject(recipient)
            If ns IsNot Nothing Then Marshal.ReleaseComObject(ns)
            If outlookApp IsNot Nothing Then Marshal.ReleaseComObject(outlookApp)
        End Try
    End Function

    Private Sub AddToNoActionDataList(app As String)
        If noActionDataList.InvokeRequired Then
            noActionDataList.Invoke(New Action(Of String)(AddressOf AddToNoActionDataList), app)
        Else
            If Not DataGridViewContains(noActionDataList, app) Then
                noActionDataList.Rows.Add(app)
            End If
        End If
    End Sub

    Private Sub AddToActionDataList(app As String)
        If actionDataList.InvokeRequired Then
            actionDataList.Invoke(New Action(Of String)(AddressOf AddToActionDataList), app)
        Else
            If Not DataGridViewContains(actionDataList, app) Then
                actionDataList.Rows.Add(app)
            End If
        End If
    End Sub

    Public Class OutlookHelper
        Public Shared Function GetNameFromEmail(email As String) As String
            Dim application As New Outlook.Application()
            Dim namespaceObj As Outlook.NameSpace = application.GetNamespace("MAPI")

            ' Get the AddressEntry object for the email address
            Dim addressEntry As Outlook.AddressEntry = namespaceObj.CreateRecipient(email).AddressEntry

            If addressEntry IsNot Nothing Then
                If addressEntry.DisplayType = Outlook.OlDisplayType.olUser Then
                    ' If the address entry represents a user, get the user's name
                    Dim user As Outlook.ExchangeUser = addressEntry.GetExchangeUser()
                    If user IsNot Nothing Then
                        Return user.FullName
                    End If
                ElseIf addressEntry.DisplayType = Outlook.OlDisplayType.olDistList Then
                    ' If the address entry represents a distribution list, get the list's name
                    Return addressEntry.Name
                End If
            End If

            ' Return "Name not found" if the name could not be retrieved
            Return "Name not found"
        End Function
    End Class

    Private Sub pd_PrintPage(sender As Object, e As Printing.PrintPageEventArgs)
        'Calculate size and position of each cell
        Dim cellWidth As Integer = 200
        Dim cellHeight As Integer = 140
        Dim xMargin As Integer = (e.PageBounds.Width - 5 * cellWidth) / 2
        Dim yMargin As Integer = (e.PageBounds.Height - 5 * cellHeight) / 2

        'Draw grid lines
        Dim pen As New Pen(Color.Black, 1)
        For row As Integer = 0 To 4
            For col As Integer = 0 To 4
                Dim x As Integer = xMargin + col * cellWidth
                Dim y As Integer = yMargin + row * cellHeight
                e.Graphics.DrawRectangle(pen, x, y, cellWidth, cellHeight)
            Next
        Next

        'Add text to each cell
        Dim fontBold18 As New Font("Arial", 18, FontStyle.Bold)
        Dim font12 As New Font("Arial", 12, FontStyle.Bold)
        Dim font10 As New Font("Arial", 10)
        Dim fontBold13 As New Font("Arial", 13, FontStyle.Bold)
        Dim textFormat As New StringFormat()
        textFormat.Alignment = StringAlignment.Center
        textFormat.LineAlignment = StringAlignment.Center


        Dim inttest As Integer = 0

        For row As Integer = 0 To 4
            For col As Integer = 0 To 4
                inttest += 1
                Dim x As Integer = xMargin + col * cellWidth
                Dim y As Integer = yMargin + row * cellHeight
                Dim cellNumber As String = (row + 1).ToString() + "," + (col + 1).ToString()
                If control > 1 Then
                    e.Graphics.DrawString(oldRsStr((control - 1) * 25 + (inttest - 1)), fontBold18, Brushes.Black, New RectangleF(x, y - 45, cellWidth, cellHeight), textFormat)
                    e.Graphics.DrawString(newRsStr((control - 1) * 25 + (inttest - 1)), font12, Brushes.Black, New RectangleF(x, y - 10 + fontBold18.Height - 20, cellWidth, cellHeight - fontBold18.Height), textFormat)
                    e.Graphics.DrawString(buildingsStr((control - 1) * 25 + (inttest - 1)), font10, Brushes.Black, New RectangleF(x, y + fontBold18.Height + font12.Height, cellWidth, cellHeight - fontBold18.Height - font12.Height - font10.Height * 2), textFormat)
                    e.Graphics.DrawString(floorsStr((control - 1) * 25 + (inttest - 1)), font10, Brushes.Black, New RectangleF(x, y + 20 + fontBold18.Height + font12.Height, cellWidth, cellHeight - fontBold18.Height - font12.Height - font10.Height * 2), textFormat)
                    e.Graphics.DrawString(roomsStr((control - 1) * 25 + (inttest - 1)), fontBold13, Brushes.Black, New RectangleF(x, y - 5 + cellHeight - fontBold13.Height, cellWidth, fontBold13.Height), textFormat)
                Else
                    e.Graphics.DrawString(oldRsStr(inttest - 1), fontBold18, Brushes.Black, New RectangleF(x, y - 45, cellWidth, cellHeight), textFormat)
                    e.Graphics.DrawString(newRsStr(inttest - 1), font12, Brushes.Black, New RectangleF(x, y - 10 + fontBold18.Height - 20, cellWidth, cellHeight - fontBold18.Height), textFormat)
                    e.Graphics.DrawString(buildingsStr(inttest - 1), font10, Brushes.Black, New RectangleF(x, y + fontBold18.Height + font12.Height, cellWidth, cellHeight - fontBold18.Height - font12.Height - font10.Height * 2), textFormat)
                    e.Graphics.DrawString(floorsStr(inttest - 1), font10, Brushes.Black, New RectangleF(x, y + 20 + fontBold18.Height + font12.Height, cellWidth, cellHeight - fontBold18.Height - font12.Height - font10.Height * 2), textFormat)
                    e.Graphics.DrawString(roomsStr(inttest - 1), fontBold13, Brushes.Black, New RectangleF(x, y - 5 + cellHeight - fontBold13.Height, cellWidth, fontBold13.Height), textFormat)
                End If

            Next
        Next


    End Sub

    Private Sub appCSVGenerateBtn_Click(sender As Object, e As EventArgs) Handles appCSVGenerateBtn.Click

        actionDataList.Rows.Clear()
        noActionDataList.Rows.Clear()

        appCSVGenerateBtn.Enabled = False

        ' Start the backgroundWorker to perform the long-running operation
        backgroundWorker.RunWorkerAsync()
        'backgroundWorker.RunWorkerAsync()

        'appCSVGenerateBtn.Enabled = False

        '' Simulate a long-running operation (e.g., processing large data)
        'Dim computerNames As String() = Nothing

        '' Execute the UI-related code on the UI thread using Invoke
        'Me.Invoke(Sub()
        '              computerNames = AutoAppRs.Text.Replace(Chr(13), "").Split(Chr(10))
        '          End Sub)

        'Dim csvFilePath As String = String.Empty

        '' Allow user to choose the save location using SaveFileDialog
        'Dim saveFileDialogResult As DialogResult = Me.Invoke(Function()
        '                                                         Dim saveFileDialog As New SaveFileDialog()
        '                                                         saveFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
        '                                                         saveFileDialog.Title = "Save as CSV"
        '                                                         If saveFileDialog.ShowDialog() = DialogResult.OK Then
        '                                                             csvFilePath = saveFileDialog.FileName
        '                                                             Return DialogResult.OK
        '                                                         Else
        '                                                             Return DialogResult.Cancel
        '                                                         End If
        '                                                     End Function)

        'If saveFileDialogResult = DialogResult.Cancel Then
        '    Return
        'End If

        'Try
        '    Using writer As New StreamWriter(csvFilePath)
        '        For Each computerName As String In computerNames
        '            ' Get the applications for the current computer name
        '            Dim applicationNames As String = GetApplicationNamesByComputerNameCSV(computerName)

        '            ' Write the results to a CSV file with each computer's applications on a new line
        '            writer.WriteLine(computerName & ",""" & applicationNames & """")
        '        Next
        '    End Using

        '    ' Notify the user on the UI thread that the CSV file is created successfully
        '    Me.Invoke(Sub()
        '                  MessageBox.Show("CSV file created successfully.")
        '              End Sub)
        'Catch ex As IOException
        '    ' File is in use, prompt the user to close it or choose an action
        '    Dim result As DialogResult = Me.Invoke(Function()
        '                                               Return MessageBox.Show("The file is currently in use by another process. Do you want to try again?", "File In Use", MessageBoxButtons.RetryCancel, MessageBoxIcon.Warning)
        '                                           End Function)
        '    If result = DialogResult.Retry Then
        '        backgroundWorker_DoWork(sender, e) ' Recursive call to retry the saving
        '    End If

        'Catch ex As Exception
        '    ' Handle other exceptions (e.g., file access permission issues, etc.)
        '    Me.Invoke(Sub()
        '                  MessageBox.Show("An error occurred while processing the data: " & ex.Message)
        '              End Sub)
        'End Try

        'appCSVGenerateBtn.Enabled = True
    End Sub

    'Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    '    Dim computerNames As String() = RichTextBox1.Text.Replace(Chr(13), "").Split(Chr(10))

    '    Using New ImpersonationHelper(domain, userName, password)
    '        Using connection As New SqlConnection(connectionString)
    '            Try
    '                connection.Open()

    '                For Each computerName As String In computerNames
    '                    Dim previousUsername As String = String.Empty

    '                    ' Fetch data from EndpointInventory
    '                    Dim cmdEI As New SqlCommand("SELECT TOP 1 * FROM blank WHERE Computer_Name = @computerName", connection)
    '                    cmdEI.Parameters.AddWithValue("@computerName", computerName)
    '                    Dim readerEI As SqlDataReader = cmdEI.ExecuteReader()

    '                    If readerEI.Read() Then
    '                        'previousUsername = readerEI("Previous_Username").ToString()
    '                        previousUsername = readerEI("Room").ToString()
    '                    End If
    '                    readerEI.Close()

    '                    If Not RichTextBox2.Text.Contains(computerName) Then
    '                        RichTextBox2.AppendText(computerName + " - " + previousUsername + vbCrLf)
    '                    End If


    '                    'If Not RichTextBox2.Text.Contains(computerName) Then
    '                    '    If previousUsername.Contains("EPR") Then
    '                    '        RichTextBox2.AppendText("ENT_PUBLIC" + vbCrLf)
    '                    '    ElseIf previousUsername.Contains("SPD") Or previousUsername.Contains("TU") Or previousUsername.Contains("WA") Then
    '                    '        RichTextBox2.AppendText("SSO Not Installed" + vbCrLf)
    '                    '    Else
    '                    '        RichTextBox2.AppendText("ENT_PRIVATE" + vbCrLf)
    '                    '    End If
    '                    'End If

    '                Next

    '            Catch ex As Exception
    '                MessageBox.Show("Error reading from database" + vbCrLf + "Please contact:" + vbCrLf + "Jacob Baker @ baker.jacob@mayo.edu" + vbCrLf + "Craig Zank @ zank.craig@mayo.edu")
    '            End Try
    '        End Using
    '    End Using


    'End Sub

    Private Sub SNSearchMenuItem_Click(sender As Object, e As EventArgs) Handles SNSearchMenuItem.Click

        Dim selectedItems = PingList.SelectedItems(0)

        TabControl1.SelectedTab = TabPage9

        searchSNTxt.Text = selectedItems.ToString().Substring(0, Math.Min(selectedItems.ToString().Length, 8))

        searchSNBtn.PerformClick()

    End Sub
    Dim previousUsername As String = String.Empty
    Private Sub PeopleSearchMenuItem_Click(sender As Object, e As EventArgs) Handles PeopleSearchMenuItem.Click

        Dim selectedItems = PingList.SelectedItems(0)

        TabControl1.SelectedTab = TabPage12

        Dim computerName As String = selectedItems.ToString().Substring(0, Math.Min(selectedItems.ToString().Length, 8))

        resultsSNSearchTxt.Clear()
        CIAppsTxt.Clear()
        DevicesTxt.Clear()

        Using New ImpersonationHelper(domain, userName, password)
            Using connection As New SqlConnection(connectionString)
                Try
                    connection.Open()

                    ' Fetch data from EndpointInventory
                    Dim cmdEI As New SqlCommand("SELECT TOP 1 * FROM blank WHERE Computer_Name = @computerName", connection)
                    cmdEI.Parameters.AddWithValue("@computerName", computerName)
                    Dim readerEI As SqlDataReader = cmdEI.ExecuteReader()



                    If readerEI.Read() Then
                        previousUsername = readerEI("Previous_Username").ToString()
                    End If
                    readerEI.Close()
                    Console.WriteLine(previousUsername)
                    'peopleSearchTxt.Text = previousUsername

                Catch ex As Exception
                    MessageBox.Show("Error reading from database" + vbCrLf + "Please contact:" + vbCrLf + "Jacob Baker @ baker.jacob@mayo.edu" + vbCrLf + "Craig Zank @ zank.craig@mayo.edu")
                End Try
            End Using
        End Using
        Console.WriteLine(previousUsername)
        peopleSearchTxt.Text = Me.previousUsername

        peopleSearchBtn.PerformClick()

    End Sub

    Private Sub CopyRoomsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyRoomsToolStripMenuItem.Click


        Dim selectedItems As New StringBuilder()

        ' Check if there are selected items
        If PingList.SelectedItems.Count = 0 Then
            MessageBox.Show("No items selected.")
            Return
        End If

        Using impersonation As New ImpersonationHelper(domain, userName, password)
            Using connection As New SqlConnection(connectionString)
                Try
                    connection.Open()
                    For Each item In PingList.SelectedItems
                        ' Fetch data from EndpointInventory
                        Using cmdEI As New SqlCommand("SELECT TOP 1 Room FROM blank WHERE Computer_Name = @item", connection)
                            ' Use only the first 8 characters of the item for matching
                            Dim itemName As String = item.ToString().Substring(0, Math.Min(item.ToString().Length, 8))
                            cmdEI.Parameters.AddWithValue("@item", itemName)

                            Using readerEI As SqlDataReader = cmdEI.ExecuteReader()
                                If readerEI.Read() Then
                                    selectedItems.AppendLine(readerEI("Room").ToString().Trim()) ' Add the username to StringBuilder
                                Else
                                    selectedItems.AppendLine("N/A")
                                End If
                            End Using
                        End Using
                    Next

                Catch ex As Exception
                    MessageBox.Show("Error reading from database" & vbCrLf & "Exception message: " & ex.Message)
                Finally
                    If connection.State = ConnectionState.Open Then
                        connection.Close()
                    End If
                End Try
            End Using
        End Using

        ' Use Invoke to perform the clipboard operation on the UI thread
        If Me.InvokeRequired Then
            Me.Invoke(New MethodInvoker(Sub() CopyToClipboard(selectedItems.ToString())))
        Else
            CopyToClipboard(selectedItems.ToString())
        End If

    End Sub

    Private Sub CopyLANIDsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyLANIDsToolStripMenuItem.Click

        Dim selectedItems As New StringBuilder()

        ' Check if there are selected items
        If PingList.SelectedItems.Count = 0 Then
            MessageBox.Show("No items selected.")
            Return
        End If

        Using impersonation As New ImpersonationHelper(domain, userName, password)
            Using connection As New SqlConnection(connectionString)
                Try
                    connection.Open()
                    For Each item In PingList.SelectedItems
                        ' Fetch data from EndpointInventory
                        Using cmdEI As New SqlCommand("SELECT TOP 1 Previous_Username FROM blank WHERE Computer_Name = @item", connection)
                            ' Use only the first 8 characters of the item for matching
                            Dim itemName As String = item.ToString().Substring(0, Math.Min(item.ToString().Length, 8))
                            cmdEI.Parameters.AddWithValue("@item", itemName)

                            Using readerEI As SqlDataReader = cmdEI.ExecuteReader()
                                If readerEI.Read() Then
                                    selectedItems.AppendLine(readerEI("Previous_Username").ToString().Trim()) ' Add the username to StringBuilder
                                Else
                                    selectedItems.AppendLine("N/A")
                                End If
                            End Using
                        End Using
                    Next

                Catch ex As Exception
                    MessageBox.Show("Error reading from database" & vbCrLf & "Exception message: " & ex.Message)
                Finally
                    If connection.State = ConnectionState.Open Then
                        connection.Close()
                    End If
                End Try
            End Using
        End Using

        ' Use Invoke to perform the clipboard operation on the UI thread
        If Me.InvokeRequired Then
            Me.Invoke(New MethodInvoker(Sub() CopyToClipboard(selectedItems.ToString())))
        Else
            CopyToClipboard(selectedItems.ToString())
        End If
    End Sub

    Private Sub CopyToClipboard(text As String)
        If String.IsNullOrEmpty(text) Then
            MessageBox.Show("No data to copy.")
        Else
            My.Computer.Clipboard.SetText(text)
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            MonthCalendar1.Enabled = True
        ElseIf CheckBox4.Checked = False Then
            MonthCalendar1.Enabled = False
        End If
    End Sub

    Private Sub CheckedListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBox1.SelectedIndexChanged

    End Sub

    Private Sub IPToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles IPToolStripMenuItem.Click
        ' Initializes a new list to store the selected IP addresses.
        Dim selectedItems As New List(Of String)

        ' Iterates through each selected item in the PingList.
        For Each item In PingList.SelectedItems
            ' Finds the position of the colon (:) in the item's string, which is used to locate the IP address.
            Dim index As Integer = item.ToString().IndexOf(":")
            ' Checks if the colon was found in the string.
            If index <> -1 Then
                ' Adds the substring after the colon (the IP address) to the selectedItems list.
                selectedItems.Add(item.ToString().Substring(index + 1))
            End If
        Next

        ' Joins the list of selected IP addresses into a single string, separated by new lines.
        Dim result As String = String.Join(Environment.NewLine, selectedItems)

        ' Copies the resulting string of IP addresses to the clipboard.
        My.Computer.Clipboard.SetText(result)
    End Sub

    Private Sub CopyStatusesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyStatusesToolStripMenuItem.Click
        ' Initializes a new list to store the selected IP addresses.
        Dim selectedItems As New List(Of String)

        ' Iterates through each selected item in the PingList.
        For Each item In PingList.SelectedItems
            If item.ToString().Contains("Fail") Then
                selectedItems.Add("Fail")
            ElseIf item.ToString().Contains("Timed Out") Then
                selectedItems.Add("Timed Out/Unreachable")
            ElseIf item.ToString().Contains("Success") Then
                selectedItems.Add("Success")
            End If
        Next
        Dim result As String = String.Join(Environment.NewLine, selectedItems)
        ' Copies the resulting string of IP addresses to the clipboard.
        My.Computer.Clipboard.SetText(result)
    End Sub

    Private Sub CopyStatusesIPsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyStatusesIPsToolStripMenuItem.Click
        ' Initializes a new list to store the selected IP addresses.
        Dim selectedItems As New List(Of String)
        Dim status As String = ""

        ' Iterates through each selected item in the PingList.
        For Each item In PingList.SelectedItems
            ' Finds the position of the colon (:) in the item's string, which is used to locate the IP address.
            Dim index As Integer = item.ToString().IndexOf(":")
            ' Checks if the colon was found in the string.
            If index <> -1 Then
                If item.ToString().Contains("Fail") Then
                    status = "Fail"
                ElseIf item.ToString().Contains("Timed Out") Then
                    status = "Timed Out/Unreachable"
                ElseIf item.ToString().Contains("Success") Then
                    status = "Success"
                End If
                ' Adds the substring after the colon (the IP address) to the selectedItems list.
                selectedItems.Add(status + " - " + item.ToString().Substring(index + 1))
            End If
        Next

        ' Joins the list of selected IP addresses into a single string, separated by new lines.
        Dim result As String = String.Join(Environment.NewLine, selectedItems)

        ' Copies the resulting string of IP addresses to the clipboard.
        My.Computer.Clipboard.SetText(result)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dialog1.ShowDialog()
    End Sub




    'Private Sub Button1_Click(sender As Object, e As EventArgs)

    '    Dim ips As String() = RichTextBox1.Text.Replace(Chr(13), "").Split(Chr(10))
    '    Dim ip As String() = RichTextBox3.Text.Replace(Chr(13), "").Split(Chr(10))

    '    For Each item In ips
    '        If Not ip.Contains(item) Then
    '            RichTextBox4.AppendText(item + vbCrLf)
    '        End If
    '    Next

    'End Sub

End Class

