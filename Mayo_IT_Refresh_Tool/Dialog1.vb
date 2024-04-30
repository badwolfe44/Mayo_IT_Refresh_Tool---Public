Imports System.IO
Imports System.Windows.Forms

Public Class Dialog1

    Private fn As String = ""
    Private ln As String = ""
    Private lan As String = ""

    Private Sub Dialog1_Load(sender As Object, e As EventArgs) Handles Me.Load
        SetDefaultText(TextBox1, "First Name")
        SetDefaultText(TextBox2, "Last Name")
        SetDefaultText(TextBox3, "LAN ID")
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

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Dim path As String = "\\R0327234\mrit\EmailRecipients.txt"
        fn = TextBox1.Text
        ln = TextBox2.Text
        lan = TextBox3.Text
        File.AppendAllText(path, Environment.NewLine + fn + " " + ln + " - " + lan)
        Form1.firstName = fn
        Form1.lastName = ln
        Form1.lanid = lan
        Form1.RefreshCheckbox()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

End Class
