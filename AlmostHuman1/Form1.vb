Public Class Form1

    Dim Worksheetcell As Integer


    Private Zustimmen As String

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            CheckBox2.Checked = False
            Zustimmen = "True"
        ElseIf CheckBox2.Checked = True Then
            CheckBox1.Checked = False

        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            CheckBox1.Checked = False
            Zustimmen = "False"

        ElseIf CheckBox1.Checked = True Then
            CheckBox2.Checked = False

        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        Dim xlsWorkBook As Microsoft.Office.Interop.Excel.Workbook
        Dim xlsWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim xls As New Microsoft.Office.Interop.Excel.Application

        Dim resourcesFolder = IO.Path.GetFullPath("C:\Users\User\Desktop\")
        Dim fileName = "AlmostHuman.xlsx"

        xlsWorkBook = xls.Workbooks.Open(resourcesFolder & fileName)
        xlsWorkSheet = xlsWorkBook.Sheets("Sheet1")



        xlsWorkSheet.Cells(1, 1) = Zustimmen
        xlsWorkSheet.Cells(1, 2) = RichTextBox1.Text
        xlsWorkSheet.Cells(1, 3) = RichTextBox2.Text

        xlsWorkBook.Close()
        xls.Quit()



    End Sub
End Class
