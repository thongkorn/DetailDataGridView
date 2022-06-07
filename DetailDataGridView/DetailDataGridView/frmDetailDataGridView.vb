#Region "ABOUT"
' / --------------------------------------------------------------------------------
' / Developer : Mr.Surapon Yodsanga (Thongkorn Tubtimkrob)
' / eMail : thongkorn@hotmail.com
' / URL: http://www.g2gnet.com (Khon Kaen - Thailand)
' / Facebook: http://www.facebook.com/g2gnet (for Thailand)
' / Facebook: http://www.facebook.com/commonindy (Worldwide)
' / More Info: http://www.g2gsoft.com
' /
' / Purpose: Demonstrate Add/Remove rows and calculate product sales results.
' / Microsoft Visual Basic .NET (2010)
' /
' / This is open source code under @CopyLeft by Thongkorn Tubtimkrob.
' / You can modify and/or distribute without to inform the developer.
' / --------------------------------------------------------------------------------
#End Region

Public Class frmDetailDataGridView

    ' / --------------------------------------------------------------------------------
    '/ Don't forget to set Form has KeyPreview = True
    Private Sub frmDetailDataGridView_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.F7
                '/ Add Row
                Call btnAddRow_Click(sender, e)
            Case Keys.F8
                '/ Remove Row
                Call btnRemoveRow_Click(sender, e)
        End Select
    End Sub

    ' / --------------------------------------------------------------------------------
    Private Sub frmDetailDataGridView_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.KeyPreview = True  '/ สามารถกดปุ่มฟังค์ชั่นคีย์ลงในฟอร์มได้
        Call InitializeGrid()
        '/
        txtSumTotal.ReadOnly = True
        txtSumTotal.Text = "0.00"
    End Sub

    ' / --------------------------------------------------------------------------------
    Private Sub InitializeGrid()
        With dgvData
            .RowHeadersVisible = False
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
            .AllowUserToResizeRows = False
            .MultiSelect = False
            .ReadOnly = False
            .RowTemplate.MinimumHeight = 27
            .RowTemplate.Height = 27
            '/ Columns Specified
            '/ Index = 0
            .Columns.Add("PK", "Primary Key")
            With .Columns("PK")
                .ReadOnly = True
                .DefaultCellStyle.BackColor = Color.LightGoldenrodYellow
                .Visible = True 'False '/ ปกติหลัก Primary Key จะต้องถูกซ่อนไว้
            End With
            '/ Index = 1
            .Columns.Add("ProductID", "Product ID")
            '/ Index = 2
            .Columns.Add("ProductName", "Product Name")
            '/ Index = 3
            .Columns.Add("Quantity", "Quantity")
            .Columns("Quantity").ValueType = GetType(Integer)
            '/ Index = 4
            .Columns.Add("UnitPrice", "Unit Price")
            .Columns("UnitPrice").ValueType = GetType(Double)
            '/ Index = 5
            .Columns.Add("Total", "Total")
            .Columns("Total").ValueType = GetType(Double)
            .Font = New Font("Tahoma", 11)
            '/ Total Column
            With .Columns("Total")
                .ReadOnly = True
                .DefaultCellStyle.BackColor = Color.LightGoldenrodYellow
                .DefaultCellStyle.ForeColor = Color.Blue
                .DefaultCellStyle.Font = New Font(dgvData.Font, FontStyle.Bold)
            End With
            '/ Alignment MiddleRight only columns 3 to 5
            For i As Byte = 3 To 5
                '/ Header Alignment
                .Columns(i).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
                '/ Cell Alignment
                .Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Next
            '/ Auto size column width of each main by sorting the field.
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            '/ Adjust Header Styles
            With .ColumnHeadersDefaultCellStyle
                .BackColor = Color.RoyalBlue
                .ForeColor = Color.White
                .Font = New Font("Tahoma", 11, FontStyle.Bold)
            End With
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .ColumnHeadersHeight = 36
            '/ กำหนดให้ EnableHeadersVisualStyles = False เพื่อให้ยอมรับการเปลี่ยนแปลงสีพื้นหลังของ Header
            .EnableHeadersVisualStyles = False
        End With

    End Sub

    ' / --------------------------------------------------------------------------------
    ' / Add new row
    Private Sub btnAddRow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        Dim Position As Integer = dgvData.Rows.Count - 1
        Dim PK As Integer = 1   '/ Initialize value if without rows.
        '/ Get value at the last row
        Dim LastRow As New DataGridViewRow
        '/ ตรวจสอบค่าแถวสุดท้ายว่ามีค่า Primary Key เท่าไหร่ก็ให้บวก 1 (เป็นการจำลองการทำงาน โดยไม่ติดต่อกับ DataBase)
        '/ กรณีใช้ฐานข้อมูลจริงๆ ให้ตัดส่วนนี้ทิ้งแล้วใช้ Primary Key ของสินค้าจากฐานข้อมูลแทน
        If Position >= 0 Then
            '/ ไปแถวสุดท้าย
            LastRow = dgvData.Rows.OfType(Of DataGridViewRow).Last()
            '/ จากนั้นให้เพิ่มค่าขึ้น +1 (Column Index = 0)
            PK = LastRow.Cells(0).Value + 1
        End If
        Dim RandomClass As New Random()
        '/ Sample data
        '/ Primary Key, Product ID, Product Name, Quantity, UnitPrice, Total
        Dim row As String() = New String() {PK, "PRO000" & PK, "Product " & PK, 1, Format(RandomClass.Next(1, 1000), "0.00"), "0.00"}
        dgvData.Rows.Add(row)
        '/ โฟกัสไปที่ Column(3) หรือ Quantity (จำนวน)
        dgvData.CurrentCell = dgvData.Rows(dgvData.RowCount - 1).Cells(3)
        dgvData.Focus()

        '/ ไปคำนวณหาค่าผลรวม
        Call CalSumTotal()
    End Sub

    ' / --------------------------------------------------------------------------------
    ' / Remove select row
    Private Sub btnRemoveRow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemove.Click
        If dgvData.RowCount = 0 Then Exit Sub
        '/
        dgvData.Rows.Remove(dgvData.CurrentRow)
        dgvData.Refresh()
        '/ เมื่อแถวรายการถูกลบออกไป และยังคงมีแถวรายการอยู่ ต้องไปคำนวณหาค่าผลรวมใหม่
        If dgvData.RowCount > 0 Then Call CalSumTotal()
    End Sub

    ' / --------------------------------------------------------------------------------
    ' / Calcualte sum of Total (Column Index = 5)
    ' / ทำทุกครั้งที่มีการเพิ่มหรือลบแถวรายการ และมีการเปลี่ยนแปลงค่าในเซลล์ Quantity, UnitPrice
    Private Sub CalSumTotal()
        txtSumTotal.Text = "0.00"
        '/ วนรอบตามจำนวนแถวที่มีอยู่ปัจจุบัน
        For i As Integer = 0 To dgvData.RowCount - 1
            '/ หลักสุดท้ายของตารางกริด = [จำนวน x ราคา]
            dgvData.Rows(i).Cells(5).Value = Format(dgvData.Rows(i).Cells(3).Value * dgvData.Rows(i).Cells(4).Value, "#,##0.00")
            '/ นำค่าจาก Total มารวมกันเพื่อแสดงผลในสรุปผลรวม (x = x + y)
            txtSumTotal.Text = Format(CDbl(txtSumTotal.Text) + CDbl(dgvData.Rows(i).Cells(5).Value), "#,##0.00")
        Next
    End Sub

    ' / --------------------------------------------------------------------------------
    ' / After you press Enter
    Private Sub dgvData_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvData.CellEndEdit
        '/ เกิดการเปลี่ยนแปลงค่าในหลัก Index ที่ 3 หรือ 4
        Select Case e.ColumnIndex
            Case 3, 4 '/ Column Index = 3 (Quantity), Column Index = 4 (UnitPrice)
                '/ Quantity
                '/ การดัก Error กรณีมีค่า Null Value ให้ใส่ค่า 0 ลงไปแทน
                If IsDBNull(dgvData.Rows(e.RowIndex).Cells(3).Value) Then dgvData.Rows(e.RowIndex).Cells(3).Value = "0"
                Dim Quantity As Integer = dgvData.Rows(e.RowIndex).Cells(3).Value
                '/ UnitPrice
                '/ If Null Value
                If IsDBNull(dgvData.Rows(e.RowIndex).Cells(4).Value) Then dgvData.Rows(e.RowIndex).Cells(4).Value = "0.00"
                Dim UnitPrice As Double = dgvData.Rows(e.RowIndex).Cells(4).Value

                '/ Quantity x UnitPrice
                dgvData.Rows(e.RowIndex).Cells(5).Value = (Quantity * UnitPrice).ToString("#,##0.00")

                '/ Calculate Summary
                Call CalSumTotal()
        End Select
    End Sub

    ' / --------------------------------------------------------------------------------
    Private Sub dgvData_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles dgvData.EditingControlShowing
        Select Case dgvData.Columns(dgvData.CurrentCell.ColumnIndex).Name
            ' / Can use both Colume Index or Field Name
            Case "Quantity", "UnitPrice"
                '/ Stop and Start event handler
                RemoveHandler e.Control.KeyPress, AddressOf ValidKeyPress
                AddHandler e.Control.KeyPress, AddressOf ValidKeyPress
        End Select
    End Sub

    ' / --------------------------------------------------------------------------------
    Private Sub ValidKeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Dim tb As TextBox = sender
        Select Case dgvData.CurrentCell.ColumnIndex
            Case 3  ' Quantity is Integer
                Select Case e.KeyChar
                    Case "0" To "9"   ' digits 0 - 9 allowed
                    Case ChrW(Keys.Back)    ' backspace allowed for deleting (Delete key automatically overrides)

                    Case Else ' everything else ....
                        ' True = CPU cancel the KeyPress event
                        e.Handled = True ' and it's just like you never pressed a key at all
                End Select

            Case 4  ' UnitPrice is Double
                Select Case e.KeyChar
                    Case "0" To "9"
                        ' Allowed "."
                    Case "."
                        ' can present "." only one
                        If InStr(tb.Text, ".") Then e.Handled = True

                    Case ChrW(Keys.Back)
                        '/ Return False is Default value

                    Case Else
                        e.Handled = True

                End Select
        End Select
    End Sub

    Private Sub frmDetailDataGridView_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
        GC.SuppressFinalize(Me)
        Application.Exit()
    End Sub

End Class
