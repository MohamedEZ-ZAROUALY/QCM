Imports System.Data.OleDb
Public Class QCM
    Dim strConnection As String = "Provider =Microsoft.Jet.OLEDB.4.0; Data Source=..\..\AppData\QCM.mdb"
    Dim oConnection As New OleDbConnection(strConnection)
    Dim strRequest As String
    Dim compteur As Integer = 0
    Dim reponsesvides As Integer = 0

    Function correction()
        Dim a As Integer
        strRequest = "SELECT count (Select NumeroQuestion from Notes where MatriculeStudent= " & CInt(TextBox1.Text) & ")  from Notes ;"
        Try
            Dim oCommand As New OleDbCommand(strRequest, oConnection)
            oConnection.Open()
            Dim oreader As OleDbDataReader = oCommand.ExecuteReader()
            a = oreader.GetValue(0)
            oConnection.Close()
        Catch ex As Exception
            MsgBox("Erreur :" + ex.Message)
        End Try
        Return a
    End Function

    Private Sub QCM_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Button4.Enabled = True
        Button1.Enabled = True
        Button2.Enabled = False
        Button3.Enabled = False
        Button5.Enabled = False
        Button6.Enabled = False
    End Sub
    Sub vides(y As Integer)
        If RadioButton1.Checked = False And RadioButton2.Checked = False And RadioButton3.Checked = False And RadioButton4.Checked = False Then
            y = y + 1
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        strRequest = "select Matricule from Students;"
        Dim a As Integer = 0
        Try
            Dim oCommand As New OleDbCommand(strRequest, oConnection)
            oConnection.Open()
            Dim oreader As OleDbDataReader = oCommand.ExecuteReader()
            While oreader.Read
                If oreader.GetValue(0) = Val(TextBox1.Text) Then
                    a = 1
                End If
            End While
            oConnection.Close()
        Catch ex As Exception
            MsgBox("Erreur :" + ex.Message)
        End Try
        strRequest = "select MotPasse from Students;"
        Dim b As Integer = 0
        Try
            Dim oCommand As New OleDbCommand(strRequest, oConnection)
            oConnection.Open()
            Dim oreader As OleDbDataReader = oCommand.ExecuteReader()
            While oreader.Read
                If oreader.GetValue(0) = CStr(TextBox2.Text) Then
                    b = 1
                End If
            End While
            oConnection.Close()
        Catch ex As Exception
            MsgBox("Erreur :" + ex.Message)
        End Try
        If a = 1 And b = 1 Then

            strRequest = "select Nom,Prenom,DateNaissance from Students where Matricule=" & Val(TextBox1.Text) & " ;"
            Try
                Dim oCommand As New OleDbCommand(strRequest, oConnection)
                oConnection.Open()
                Dim oreader As OleDbDataReader = oCommand.ExecuteReader()
                While oreader.Read
                    Label5.Text = oreader.GetString(0)
                    Label7.Text = oreader.GetString(1)
                    Label9.Text = oreader.GetValue(2)

                End While
                oConnection.Close()
            Catch ex As Exception
                MsgBox("Erreur :" + ex.Message)
            End Try
        Else
            MsgBox("Les informations que vous avez rentré sont erronés ")
            Me.Close()

        End If
        strRequest = "select distinct NumeroQCM from Questions;"
        Try
            Dim oCommand As New OleDbCommand(strRequest, oConnection)
            oConnection.Open()
            Dim oreader As OleDbDataReader = oCommand.ExecuteReader()
            While oreader.Read
                ComboBox1.Items.Add(oreader.GetValue(0))
            End While
            oConnection.Close()
        Catch ex As Exception
            MsgBox("Erreur :" + ex.Message)
        End Try
        Button1.Enabled = False
        Button2.Enabled = True
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Button6.Enabled = True
        Button2.Enabled = False
        strRequest = "select Question,Reponse1,Reponse2,Reponse3,Reponse4 from Questions where NumeroQCM=" & CInt(ComboBox1.SelectedItem) & " ;"
        Dim i As Integer
        Try
            Dim oCommand As New OleDbCommand(strRequest, oConnection)
            oConnection.Open()
            Dim oreader As OleDbDataReader = oCommand.ExecuteReader()
            While oreader.Read
                If compteur = i Then
                    Label13.Text = oreader.GetString(0)
                    Label12.Text = CStr(compteur + 1)
                    RadioButton1.Text = oreader.GetString(1)
                    RadioButton2.Text = oreader.GetString(2)
                    RadioButton3.Text = oreader.GetString(3)
                    RadioButton4.Text = oreader.GetString(4)
                End If
                i = i + 1
            End While
            oConnection.Close()
        Catch ex As Exception
            MsgBox("Erreur :" + ex.Message)
        End Try
        Button3.Enabled = True
        Button4.Enabled = False

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Button5.Enabled = False
        vides(reponsesvides)
        If MsgBox("Vous voulez vraiment terminer le QCM ? ", MsgBoxStyle.YesNo, vbQuestion) = vbYes Then
            MsgBox("Les réponses vides sont = " & reponsesvides)
            Button3.Enabled = False
            Button4.Enabled = True


        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        RadioButton1.Checked = False
        RadioButton2.Checked = False
        RadioButton3.Checked = False
        RadioButton4.Checked = False
        compteur = compteur + 1
        Button5.Enabled = True
        strRequest = "select Question,Reponse1,Reponse2,Reponse3,Reponse4 from Questions where NumeroQCM=" & Val(ComboBox1.SelectedItem) & " ;"
        Dim i As Integer = 0
        Try
            Dim oCommand As New OleDbCommand(strRequest, oConnection)
            oConnection.Open()
            Dim oreader As OleDbDataReader = oCommand.ExecuteReader()
            While oreader.Read
                If compteur = i Then
                    Label13.Text = oreader.GetString(0)
                    Label12.Text = CStr(compteur + 1)
                    RadioButton1.Text = oreader.GetString(1)
                    RadioButton2.Text = oreader.GetString(2)
                    RadioButton3.Text = oreader.GetString(3)
                    RadioButton4.Text = oreader.GetString(4)
                End If
                i = i + 1
            End While

            oConnection.Close()
        Catch ex As Exception
            MsgBox("Erreur :" + ex.Message)
        End Try
        If compteur = 3 Then
            Button6.Enabled = False
        End If
        vides(reponsesvides)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        RadioButton1.Checked = False
        RadioButton2.Checked = False
        RadioButton3.Checked = False
        RadioButton4.Checked = False
        compteur = compteur - 1
        Button6.Enabled = True
        strRequest = "select Question,Reponse1,Reponse2,Reponse3,Reponse4 from Questions where NumeroQCM=" & Val(ComboBox1.SelectedItem) & " ;"
        Dim i As Integer = 0
        Try
            Dim oCommand As New OleDbCommand(strRequest, oConnection)
            oConnection.Open()
            Dim oreader As OleDbDataReader = oCommand.ExecuteReader()
            While oreader.Read
                If compteur = i Then
                    Label13.Text = oreader.GetString(0)
                    Label12.Text = CStr(compteur + 1)
                    RadioButton1.Text = oreader.GetString(1)
                    RadioButton2.Text = oreader.GetString(2)
                    RadioButton3.Text = oreader.GetString(3)
                    RadioButton4.Text = oreader.GetString(4)
                End If
                i = i + 1
            End While

            oConnection.Close()
        Catch ex As Exception
            MsgBox("Erreur :" + ex.Message)
        End Try
        If compteur = 0 Then
            Button5.Enabled = False
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If MsgBox("Vous êtes sûr de vouloir fermer l'application ?", vbOKCancel + vbQuestion) = vbOK Then
            Me.Close()
        End If
    End Sub
End Class
