Public Class Information


    Dim note_type As String



    Function information_run(ByVal title As String, ByVal type As String)
        txt_monitor.Text = ""
        Me.Text = title

        Me.Show()

        Timer1.Enabled = True
        note_type = type
        note_display = True
    End Function
   
    Private Sub Information_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick


        System.Windows.Forms.Application.DoEvents()

        If (note_display = False) Or (run = False) Then
            txt_monitor.BackColor = Color.DarkGray
            Timer1.Enabled = False
            Me.Hide()
        End If

        lbl_title.Text = note_type

        Select Case note_type
            Case note_delay

                txt_monitor.Text = note_value & "s"
                txt_monitor.BackColor = Color.LightYellow
            Case note_run
                System.Windows.Forms.Application.DoEvents()
                txt_monitor.Text = note_string
                If txt_monitor.BackColor = Color.LightYellow Then

                    txt_monitor.BackColor = Color.LightCyan
                Else
                    txt_monitor.BackColor = Color.LightYellow
                End If
            Case note_count
                txt_monitor.Text = note_value
                txt_monitor.BackColor = Color.LightCyan
        End Select

      
    End Sub
End Class