Imports MySql.Data.MySqlClient
Module Module1

    'Dim csvPath As String = "D:\Users\z112\source\repos\ConsoleApp2\stock_data\"
    Dim csvPath As String = "C:\Users\r2d2\source\repos\chart_gallery\stock_data\"
    'ReadOnly csvPath As String = "C:\Users\r2d2\source\repos\chart_gallery\stock_data\"

    Sub Main()
        'ConnectMySql()

        'MakeZip()
        'Exit Sub

        'GetCmdParam()

        Dim securitiesCodes() As Integer = {
            9101, 6326
        }

        Dim Prices As New ActiveMarket.Prices
        Dim Calendar As New ActiveMarket.Calendar
        Dim hash As New Hashtable
        Dim date_position As Integer
        Dim date_range As Integer
        Dim output As String
        Dim csvFile As String

        For securitiesCode = 0 To UBound(securitiesCodes)
            csvFile = csvPath + CType(securitiesCodes(securitiesCode), String) + ".csv"
            'ファイル削除
            'System.IO.File.Delete(csvFile)

            Prices.Read(securitiesCodes(securitiesCode))

            'csvファイル存在確認
            If System.IO.File.Exists(csvFile) Then
                '存在したら最終行(改行を含まず)のdate_positionの数値を取得
                Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift_JIS") 'CSVファイルのエンコードを指定（今回はShift_JIS）
                Dim lineCount As Integer = UBound(My.Computer.FileSystem.OpenTextFileReader(csvFile).ReadToEnd.Split(Chr(13))) - 2 '先頭2行(銘柄名、ヘッダ)を除外
                Dim fs As New System.IO.FileStream(csvFile, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite)
                Dim srr As New System.IO.StreamReader(fs, enc) '読み込むファイルを開く
                Dim line, posi As String
                line = srr.ReadLine()
                For i As Integer = 1 To lineCount - 1
                    line = srr.ReadLine()
                Next i
                srr.Close()
                srr = Nothing

                posi = CType(Right(line, 5).Replace("""", ""), Integer)

                date_position = posi
            Else
                date_position = Prices.Begin() '5632?

                'csvヘッダー表示
                OutputCsv(CType(securitiesCodes(securitiesCode), String) + " " + Prices.Name + ",,,,," + vbCrLf _
                    + """date"",""open"",""hight"",""low"",""close"",""power"",""End"",""date_position""", csvFile)
            End If

            date_range = Prices.End - date_position
            Dim stock_array(date_range, 7) As String

            For i = 0 To date_range - 1
                date_position = date_position + 1
                If Prices.IsClosed(date_position) Then
                    Continue For
                End If

                stock_array(i, 0) = Format(Calendar.Date(date_position), "yyyy-MM-dd")
                stock_array(i, 1) = Math.Floor(Prices.Open(date_position))
                stock_array(i, 2) = Math.Floor(Prices.High(date_position))
                stock_array(i, 3) = Math.Floor(Prices.Low(date_position))
                stock_array(i, 4) = Math.Floor(Prices.Close(date_position))
                stock_array(i, 5) = Math.Floor(Prices.Volume(date_position) * 1000)
                stock_array(i, 6) = Math.Floor(Prices.Close(date_position))
                stock_array(i, 7) = date_position
            Next

            Try
            Catch ex As Exception
            End Try

            hash.Add(securitiesCodes(securitiesCode), Prices.Name)


            For i = 0 To date_range
                If IsNothing(stock_array(i, 0)) Then
                    Continue For
                End If

                output = """" + stock_array(i, 0) + """,""" + stock_array(i, 1) + """,""" + stock_array(i, 2) + """,""" + stock_array(i, 3) _
                     + """,""" + stock_array(i, 4) + """,""" + stock_array(i, 5) + """,""" + stock_array(i, 6) + """,""" + stock_array(i, 7) + """"
                'System.Diagnostics.Debug.WriteLine(output)
                OutputCsv(output, csvFile)
            Next
        Next

    End Sub

    Sub OutputCsv(output, csvFile)
        Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift_JIS") 'CSVファイルのエンコードを指定（今回はShift_JIS）
        Dim sr As New System.IO.StreamWriter(csvFile + "", True, enc) '書き込むファイルを開く
        sr.Write(output)
        sr.Write(vbCrLf) '改行
        sr.Close()
    End Sub

    Sub GetCmdParam()
        Console.WriteLine(System.Environment.CommandLine) 'コマンドライン引数を表示する
        Dim cmds As String() = System.Environment.GetCommandLineArgs() 'コマンドライン引数を配列で取得する
        Dim cmd As String 'コマンドライン引数を列挙する
        For Each cmd In cmds
            Console.WriteLine(cmd)
        Next
    End Sub

    Sub MakeZip()
        'ZIP書庫を作成
        'System.IO.Compression.ZipFile.CreateFromDirectory("C:\temp\test\dir", "C:\temp\test\1.zip", System.IO.Compression.CompressionLevel.Optimal, False, System.Text.Encoding.GetEncoding("shift_jis"))
    End Sub

    Sub ConnectMySql()
        Const DB_Source As String = "localhost"
        Const DB_Port As String = "3306"
        Const DB_Name As String = "stocks"
        Const DB_Id As String = "root"
        Const DB_Pw As String = "rage5557"

        Using Conn As New MySqlConnection("Database=" + DB_Name _
                                        + ";Data Source=" + DB_Source _
                                        + ";Port=" + DB_Port _
                                        + ";User Id=" + DB_Id _
                                        + ";Password=" + DB_Pw _
                                        + ";sqlservermode=True;")
            Conn.Open()
            'data追加
            'Dim query = "INSERT INTO test VALUES (2, 'test2')"
            'Dim cmd As MySqlCommand = New MySqlCommand(query, Conn)
            'cmd.ExecuteNonQuery()

            'select
            Dim query = "select * from stocks.s6326 limit 1"
            Dim cmd As MySqlCommand = New MySqlCommand(query, Conn)
            Dim data As MySqlDataReader = cmd.ExecuteReader
            '結果を表示
            While data.Read()
                Console.WriteLine(data("close"))
            End While

            Conn.Close()

            Conn.Open()
            Dim q = "insert into s6326 values ('1989-02-21','985','990','970','979','1957000','979','1');"
            Dim c As MySqlCommand = New MySqlCommand(q, Conn)
            c.ExecuteNonQuery()

            Conn.Close()
        End Using
    End Sub
End Module