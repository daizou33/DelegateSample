Public Class CL_FOLDER_WATCH

    'コントロールを扱うためのデリゲート宣言
    Delegate Sub ListRefreshDelegate()

    'デリゲート宣言をデータ型とした変数を作成
    Public RefreshDelegate As ListRefreshDelegate

    'フォームの参照を保持する
    Private _FM01 As FM01

    Private watcher As System.IO.FileSystemWatcher = Nothing

    Dim strFilter As String = "*.*"
    Dim sFileSrcPath As String = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="frm"></param>
    ''' <remarks></remarks>
    Public Sub New(frm As Form)
        '呼び出し元のフォームインスタンス
        _FM01 = CType(frm, FM01)

        '呼び出し元フォームにあるListRefresh関数をDelegateにセット
        RefreshDelegate = New ListRefreshDelegate(AddressOf _FM01.ListRefresh)

    End Sub

    ''' <summary>
    ''' 監視を開始する。
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Watch_Start()

        If Not (watcher Is Nothing) Then
            Return
        End If

        watcher = New System.IO.FileSystemWatcher

        '監視するディレクトリを指定
        watcher.Path = sFileSrcPath

        '最終アクセス日時、最終更新日時、ファイル、フォルダ名の変更を監視する
        watcher.NotifyFilter = System.IO.NotifyFilters.LastAccess Or _
            System.IO.NotifyFilters.LastWrite Or _
            System.IO.NotifyFilters.FileName Or _
            System.IO.NotifyFilters.DirectoryName

        'すべてのファイルを監視
        watcher.Filter = strFilter

        'イベントハンドラの追加
        AddHandler watcher.Changed, AddressOf watcher_Changed
        AddHandler watcher.Created, AddressOf watcher_Changed
        AddHandler watcher.Deleted, AddressOf watcher_Changed
        AddHandler watcher.Renamed, AddressOf watcher_Changed

        '監視を開始する
        watcher.EnableRaisingEvents = True
        Console.WriteLine("監視を開始しました。")

    End Sub

    ''' <summary>
    ''' 監視を停止する。
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Watch_STOP()
        If Not (watcher Is Nothing) Then
            watcher.EnableRaisingEvents = False
            watcher.Dispose()
            watcher = Nothing
            Console.WriteLine("監視を終了しました。")
        End If
    End Sub

    ''' <summary>
    ''' 監視対象に変化があった場合の処理
    ''' </summary>
    ''' <param name="source"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub watcher_Changed(ByVal source As System.Object, ByVal e As System.IO.FileSystemEventArgs)
        Select Case e.ChangeType
            Case System.IO.WatcherChangeTypes.Changed
                Console.WriteLine(("ファイル 「" + e.FullPath + "」が変更されました。"))

            Case System.IO.WatcherChangeTypes.Created
                Console.WriteLine(("ファイル 「" + e.FullPath + "」が作成されました。"))

            Case System.IO.WatcherChangeTypes.Deleted
                Console.WriteLine(("ファイル 「" + e.FullPath + "」が削除されました。"))

            Case System.IO.WatcherChangeTypes.Renamed
                Console.WriteLine(("ファイル 「" + e.FullPath + "」の名前が変更されました。"))
        End Select

        'FM01のListRefreshメソッドを呼ぶ
        _FM01.Invoke(RefreshDelegate, New Object() {})

    End Sub


End Class
