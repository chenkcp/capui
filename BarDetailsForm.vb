Imports NextCapServer = ncapsrv
Imports ncapsrv

Public Class BarDetailsForm
    Public Sub LoadBarDetails(ByVal lot As String)
        ' Update the label text with lot and current time
        LabelLotAndTime.Text = $"{lot} {DateTime.Now:yyyy-MM-dd HH:mm tt}"

        ' Assume clsLot has properties corresponding to the columns
        DataGridViewBarDetails.Rows.Clear()
        DataGridViewBarDetails.Columns.Clear()

        DataGridViewBarDetails.Columns.Add("LotId", "Lot ID")
        DataGridViewBarDetails.Columns.Add("Inspection", "Inspection")
        DataGridViewBarDetails.Columns.Add("NumberOf", "Number of")
        DataGridViewBarDetails.Columns.Add("Disposition", "Disposition")
        DataGridViewBarDetails.Columns.Add("Product", "Product")
        DataGridViewBarDetails.Columns.Add("Class", "Class")
        DataGridViewBarDetails.Columns.Add("CodeL1", "Code L1")

        ' generate 5 sample rows
        For i As Integer = 0 To 4
            Dim row As DataGridViewRow = New DataGridViewRow()
            row.CreateCells(DataGridViewBarDetails)

            row.Cells(0).Value = lot
            ' this is fake data, change to real data if you needed
            row.Cells(1).Value = $"Inspection Value {i + 1}"
            row.Cells(2).Value = $"Number of Value {i + 1}"
            row.Cells(3).Value = $"Disposition Value {i + 1}"
            row.Cells(4).Value = $"Product Value {i + 1}"
            row.Cells(5).Value = $"Class Value {i + 1}"
            row.Cells(6).Value = $"Code L1 Value {i + 1}"

            ' add to DataGridView
            DataGridViewBarDetails.Rows.Add(row)
        Next
    End Sub

    Private Sub okButton_Click(sender As Object, e As EventArgs) Handles okButton.Click
        MessageBox.Show("OK Button Clicked!")
    End Sub

    Private Sub editButton_Click(sender As Object, e As EventArgs) Handles editButton.Click
        MessageBox.Show("Edit Button Clicked!")
    End Sub

    Private Sub deleteButton_Click(sender As Object, e As EventArgs) Handles deleteButton.Click
        MessageBox.Show("Delete Button Clicked!")
    End Sub
End Class