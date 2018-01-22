Public Class Form1
    '複数タスクでファイルをコピーするプログラム
    Private Async Function CopyFiles(ByVal files As String()) As Task(Of Boolean)
        'タスクをString()の長さ分だけ作成（ラムダ式）
        Dim tasks As IEnumerable(Of Task(Of Boolean)) =
            files.Select(Function(file) (
            Task.Run(Function()
                         'Exceptionを拾ったら失敗としてFalseを格納する
                         Try
                             'テスト用なので強制上書きを有効に、copy\にコピーしている
                             IO.File.Copy(file, "copy\" + IO.Path.GetFileName(file), True)
                         Catch ex As Exception
                             MsgBox(ex.Message)
                             Return False
                         End Try
                         Return True
                     End Function)
                     ))
        Dim result As IEnumerable(Of Boolean) = Await Task.WhenAll(tasks) '結果待ち

        '失敗チェック、失敗が有ればFalseを返す
        For Each flag As Boolean In result
            If flag = False Then
                Return False
            End If
        Next
        Return True
    End Function

    'おまけ、ファイル削除
    Private Async Function DeleteFiles(ByVal files As String()) As Task(Of Boolean)
        'タスクをString()の長さ分だけ作成（ラムダ式）
        Dim tasks As IEnumerable(Of Task(Of Boolean)) =
            files.Select(Function(file) (
            Task.Run(Function()
                         'Exceptionを拾ったら失敗としてFalseを格納する
                         Try
                             IO.File.Delete("copy\" + IO.Path.GetFileName(file))
                         Catch ex As Exception
                             MsgBox(ex.Message)
                             Return False
                         End Try
                         Return True
                     End Function)
                     ))
        Dim result As IEnumerable(Of Boolean) = Await Task.WhenAll(tasks) '結果待ち

        '失敗チェック、失敗が有ればFalseを返す
        For Each flag As Boolean In result
            If flag = False Then
                Return False
            End If
        Next
        Return True
    End Function

    Private Async Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '実行ファイル以下のoriginalという名前のディレクトリに含まれるファイル一覧を取得
        Dim files As String() = IO.Directory.GetFiles("original\", "*", IO.SearchOption.TopDirectoryOnly)
        '取得したファイル一覧に対してコピー開始
        Dim task As Task(Of Boolean) = CopyFiles(files)
        'Dim task As Task(Of Boolean) = DeleteFiles(files)

        '何回も押されたくないので無効化
        Button1.Enabled = False
        MsgBox("処理開始")
        '処理待ち
        Dim result As Boolean = Await task
        Button1.Enabled = True
        MsgBox(result)
    End Sub
End Class
