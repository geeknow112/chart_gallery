﻿Module Module1

    Dim csvPath As String = "C:\Users\testUser\source\repos\chart_gallery\stock_data\"

    Sub Main()
        'makeZip()
        'Exit Sub

        'getCmdParam()

        Dim securitiesCodes() As Integer = {
            9101
        }
        'Dim securitiesCodes() As Integer = {
        '    9101, 9104, 9107, '海運業
        '    4021, '日産化
        '    4183, '三井化学
        '    4005, '住友化
        '    4188, '三菱ケミＨＤ
        '    4911, '資生堂
        '    3407, '旭化成
        '    4042, '東ソー
        '    6988, '日東電
        '    3405, 'クラレ
        '    4061, 'デンカ
        '    4208, '宇部興
        '    4272, '日化薬
        '    4004, '昭電工
        '    4631, 'ＤＩＣ
        '    4043, 'トクヤマ
        '    4901, '富士フイルム
        '    4452, '花王
        '    4063, '信越化
        '    9984  'ソフトバンクグループ
        '}

        Dim Prices As New ActiveMarket.Prices
        Dim Calendar As New ActiveMarket.Calendar
        Dim hash As New Hashtable
        Dim date_position As Integer
        Dim date_range, test As Integer
        Dim output As String
        Dim csvFile As String

        For securitiesCode = 0 To UBound(securitiesCodes)
            Prices.Read(securitiesCodes(securitiesCode))
            date_position = Prices.Begin() '5632?
            date_range = Prices.End - date_position
            Dim stock_array(date_range, 7) As String

            For i = 0 To date_range - 1
                date_position = date_position + 1
                If Prices.IsClosed(date_position) Then
                    Continue For
                End If

                stock_array(i, 0) = Format(Calendar.Date(date_position), "yyyy-MM-dd")
                stock_array(i, 1) = Prices.Open(date_position)
                stock_array(i, 2) = Prices.High(date_position)
                stock_array(i, 3) = Prices.Low(date_position)
                stock_array(i, 4) = Prices.Close(date_position)
                stock_array(i, 5) = Math.Floor(Prices.Volume(date_position) * 1000)
                stock_array(i, 6) = Prices.Close(date_position)
                stock_array(i, 7) = date_position
            Next

            Try
            Catch ex As Exception
            End Try

            hash.Add(securitiesCodes(securitiesCode), Prices.Name)

            csvFile = csvPath + CType(securitiesCodes(securitiesCode), String) + ".csv"

            'ファイル削除
            System.IO.File.Delete(csvFile)

            'csvヘッダー表示
            outputCsv(CType(securitiesCodes(securitiesCode), String) + " " + Prices.Name + ",,,,," + vbCrLf _
                    + """date"",""open"",""hight"",""low"",""close"",""power"",""End"",""date_position""", csvFile)

            For i = 0 To date_range
                If IsNothing(stock_array(i, 0)) Then
                    Continue For
                End If

                output = """" + stock_array(i, 0) + """,""" + stock_array(i, 1) + """,""" + stock_array(i, 2) + """,""" + stock_array(i, 3) _
                     + """,""" + stock_array(i, 4) + """,""" + stock_array(i, 5) + """,""" + stock_array(i, 6) + """,""" + stock_array(i, 7) + """"
                'System.Diagnostics.Debug.WriteLine(output)
                outputCsv(output, csvFile)
            Next
        Next

    End Sub

    Sub outputCsv(output, csvFile)
        Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift_JIS") 'CSVファイルのエンコードを指定（今回はShift_JIS）
        Dim sr As New System.IO.StreamWriter(csvFile, True, enc) '書き込むファイルを開く
        sr.Write(output)
        sr.Write(vbCrLf) '改行
        sr.Close()
    End Sub

    Sub getCmdParam()
        Console.WriteLine(System.Environment.CommandLine) 'コマンドライン引数を表示する
        Dim cmds As String() = System.Environment.GetCommandLineArgs() 'コマンドライン引数を配列で取得する
        Dim cmd As String 'コマンドライン引数を列挙する
        For Each cmd In cmds
            Console.WriteLine(cmd)
        Next
    End Sub

    Sub makeZip()
        'ZIP書庫を作成
        System.IO.Compression.ZipFile.CreateFromDirectory("C:\temp\test\dir", "C:\temp\test\1.zip", System.IO.Compression.CompressionLevel.Optimal, False, System.Text.Encoding.GetEncoding("shift_jis"))
    End Sub

End Module