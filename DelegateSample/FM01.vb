Public Class FM01

    Private cls_watch As CL_FOLDER_WATCH '監視


    Private Sub FM01_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '監視スタート
        cls_watch = New CL_FOLDER_WATCH(Me)
        cls_watch.Watch_Start()

    End Sub


    ''' <summary>
    ''' DGVをリフレッシュする。
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ListRefresh()
        Dim mydoc As String = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)

        Dim di As New System.IO.DirectoryInfo(mydoc)
        Dim files As System.IO.FileInfo() = di.GetFiles("*.*", System.IO.SearchOption.TopDirectoryOnly)

        Dim dt As New DataTable

        dt.Columns.Add("Filename")
        dt.Columns.Add("FilePath")
        dt.Columns.Add("FileCreateTime")
        dt.Columns.Add("FileLastWriteTime")

        Dim row As DataRow

        For Each a As System.IO.FileInfo In files
            row = dt.NewRow
            row.Item("Filename") = a.Name
            row.Item("FilePath") = a.FullName
            row.Item("FileCreateTime") = a.CreationTime
            row.Item("FileLastWriteTime") = a.LastWriteTime
            dt.Rows.Add(row)

        Next

        DataGridView1.DataSource = dt


    End Sub

End Class
