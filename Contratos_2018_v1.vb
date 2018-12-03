Imports Microsoft.Office.Interop.Word 'control de office
Imports System.IO 'sistema de archivos
Imports Microsoft.Office.Interop
Public Class frmBusqueda

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs)
        DatosBindingSource.Filter = "A1 LIKE '%" & TextBox1.Text & "%'"

    End Sub
    Private Sub AsignaturaBindingNavigatorSaveItem_Click(sender As Object, e As EventArgs)
        Me.Validate()
        Me.AsignaturaBindingSource.EndEdit()
        Me.TableAdapterManager.UpdateAll(Me.DOCENTESDataSet)

    End Sub

    Private Sub frmBusqueda_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Creditos.ShowDialog()
    End Sub

    Private Sub frmBusqueda_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        'TODO: esta línea de código carga datos en la tabla 'DOCENTESDataSet1.Asignatura' Puede moverla o quitarla según sea necesario.
        Me.AsignaturaTableAdapter1.Fill(Me.DocentesDataSet1.Asignatura)
        'TODO: esta línea de código carga datos en la tabla 'DOCENTESDataSet1.Datos' Puede moverla o quitarla según sea necesario.
        Me.DatosTableAdapter1.Fill(Me.DocentesDataSet1.Datos)

    End Sub

    Private Sub A1ListBox_SelectedIndexChanged(sender As Object, e As EventArgs)
        Me.AsignaturaTableAdapter.Fill(Me.DOCENTESDataSet.Asignatura)
    End Sub

    Private Sub A1TextBox_TextChanged(sender As Object, e As EventArgs) Handles A1TextBox.TextChanged

    End Sub

    Private Sub A1Label_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub A7Label_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub A8Label_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedIndexChanged

    End Sub

    Private Sub TextBox1_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        DatosBindingSource.Filter = "A1 LIKE '%" & TextBox1.Text & "%'"

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox2_TextChanged_2(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        AsignaturaBindingSource.Filter = "A1 LIKE '%" & TextBox2.Text & "%'"

    End Sub

    Private Sub Label9_Click(sender As Object, e As EventArgs) Handles Label9.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ListBox3.Items.Add(ListBox1.Text)
        lsbModalidad.Items.Add(TextBox3.Text)
        ListBox4.Items.Add(A2TextBox.Text)

    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        TextBox3.Text = "Regular"

    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        TextBox3.Text = "Encuentro"

    End Sub

    Private Sub ListBox3_SelectedValueChanged(sender As Object, e As EventArgs) Handles ListBox3.SelectedValueChanged

        If (ListBox3.SelectedIndex > -1) Then
            ListBox3.Items.RemoveAt(ListBox3.SelectedIndex)

        End If
    End Sub

    Private Sub ListBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox4.SelectedIndexChanged
        If (ListBox4.SelectedIndex > -1) Then
            ListBox4.Items.RemoveAt(ListBox4.SelectedIndex)
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim suma As Integer
        Dim total As Double

        For Each elemento In ListBox4.Items
            suma += elemento.ToString
        Next

        If (A7TextBox.Text = "Tec.") Then
            total = suma * 112.55
            TextBox4.Text = Format((suma * 112.55), "##,##00.00")
            TextBox5.Text = Numalet.ToCardinal(total)

        End If
        If (A7TextBox.Text = "Lic.") Then
            total = suma * 159.3
            TextBox4.Text = Format((suma * 159.3), "##,##00.00")
            TextBox5.Text = Numalet.ToCardinal(total)
        End If
        If (A7TextBox.Text = "Esp.") Then
            total = suma * 178.45
            TextBox4.Text = Format((suma * 178.45), "##,##00.00")
            TextBox5.Text = Numalet.ToCardinal(total)
        End If

        If (A7TextBox.Text = "MSc.") Then
            total = suma * 195.44
            TextBox4.Text = Format((suma * 195.44), "##,##00.00")
            TextBox5.Text = Numalet.ToCardinal(total)
        End If
        If (A7TextBox.Text = "Dr.") Then
            total = suma * 212.45
            TextBox4.Text = Format((suma * 212.45), "##,##00.00")
            TextBox5.Text = Numalet.ToCardinal(total)
        End If

    End Sub

    Private Sub DatosBindingNavigatorSaveItem_Click(sender As Object, e As EventArgs)
        Me.Validate()
        Me.DatosBindingSource.EndEdit()
        Me.TableAdapterManager1.UpdateAll(Me.DocentesDataSet1)

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'Generar Archivo de Word
        Dim MSWord As New Word.Application
        Dim Documento As Word.Document
        'Dim carpeta As String = My.Computer.FileSystem.CurrentDirectory
        Dim carpeta As String = Environment.CurrentDirectory
        'MsgBox("El Contrato se guardará en : C:\Contratos\GENERADOS\" & A1TextBox.Text & ".docx")
        FileCopy(carpeta + "\plantilla\plantilla.docx", carpeta + "\Generados\" & A1TextBox.Text & ".docx")
        Documento = MSWord.Documents.Open(carpeta + "\Generados\" & A1TextBox.Text & ".docx")

        Documento.Bookmarks.Item("var1").Range.Text = A1TextBox.Text
        Documento.Bookmarks.Item("x8").Range.Text = A8TextBox.Text
        Documento.Bookmarks.Item("x9").Range.Text = A9TextBox.Text
        Documento.Bookmarks.Item("var2").Range.Text = A2TextBox1.Text
        Documento.Bookmarks.Item("x10").Range.Text = A12TextBox.Text
        Documento.Bookmarks.Item("Var3").Range.Text = A7TextBox.Text
        Documento.Bookmarks.Item("Var8").Range.Text = DateTimePicker1.Text
        Documento.Bookmarks.Item("Var9").Range.Text = DateTimePicker2.Text
        Documento.Bookmarks.Item("Var4").Range.Text = A6TextBox.Text
        Documento.Bookmarks.Item("x11").Range.Text = A11TextBox.Text
        Documento.Bookmarks.Item("x13").Range.Text = A13TextBox.Text
        Documento.Bookmarks.Item("x12").Range.Text = A10TextBox.Text
        Documento.Bookmarks.Item("Var10").Range.Text = NumericUpDown1.Text
        Documento.Bookmarks.Item("Var11").Range.Text = ComboBox1.Text
        Documento.Bookmarks.Item("Var7").Range.Text = A3TextBox.Text
        Documento.Bookmarks.Item("Var6").Range.Text = A4TextBox.Text
        Documento.Bookmarks.Item("Var5").Range.Text = A5TextBox.Text
        Documento.Bookmarks.Item("pago").Range.Text = TextBox4.Text
        Documento.Bookmarks.Item("pago2").Range.Text = TextBox5.Text

        Dim regularTotal = 0.0, sabatinoTotal = 0.0, total = 0.0

        For index = 0 To lsbModalidad.Items.Count - 1
            Documento.Bookmarks.Item("Asig" & (index + 1)).Range.Text = ListBox3.Items(index).ToString()
            If (lsbModalidad.Items(index).ToString() = "Regular") Then
                Documento.Bookmarks.Item("R" & (index + 1)).Range.Text = ListBox4.Items(index).ToString()
                regularTotal += Decimal.Parse(ListBox4.Items(index))
            Else
                Documento.Bookmarks.Item("S" & (index + 1)).Range.Text = ListBox4.Items(index).ToString()
                sabatinoTotal += Decimal.Parse(ListBox4.Items(index))
            End If
        Next

        total = regularTotal + sabatinoTotal

        Documento.Bookmarks.Item("t1").Range.Text = regularTotal.ToString("##0.00")
        Documento.Bookmarks.Item("t2").Range.Text = sabatinoTotal.ToString("##0.00")
        Documento.Bookmarks.Item("t3").Range.Text = total.ToString("##0.00")

        Documento.Save()
        MSWord.Visible = True
    End Sub

    Private Sub Timer1_Tick_1(sender As Object, e As EventArgs) Handles Timer1.Tick
        TextBox6.Text = String.Format("{0:G}", DateTime.Now)
    End Sub

    Private Sub Label10_Click(sender As Object, e As EventArgs) Handles Label10.Click

    End Sub

    Private Sub TextBox6_MaskInputRejected(sender As Object, e As MaskInputRejectedEventArgs) Handles TextBox6.MaskInputRejected

    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub Panel1_DoubleClick(sender As Object, e As EventArgs) Handles Panel1.DoubleClick
        Creditos.ShowDialog()
    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint

    End Sub
End Class

