Module Module1

    'Dim csv As String = "D:\Users\z112\source\repos\ConsoleApp2\stock_info.csv"
    Dim csv As String = "C:\Users\r2d2\Downloads\git_test\chart_analytics\9101_2019.csv"

    Sub Main()
        'getCmdParam()

        Dim securitiesCodes() As Integer = {
            9101, 9104
        }

        Dim Prices As New ActiveMarket.Prices
        Prices.Read(securitiesCodes(0))

        Dim Calendar As New ActiveMarket.Calendar

        '9104 東証1部 (株)商船三井（海運業）,,,,,
        '"date","open","hight","low","close","power","end"
        '"2019-01-04","2300","2334","2251","2318","1113300","2318"

        Dim hash As New Hashtable
        Dim date_position As Integer
        Dim date_range As Integer

        date_position = 7900
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
        Next

        Try
        Catch ex As Exception

        End Try

        hash.Add("9101", Prices.Name)


        'System.Diagnostics.Trace.WriteLine("test trace messate !!!")
        'System.Diagnostics.Debug.WriteLine(Prices.Name)

        'For Each a As String In arr
        '    System.Diagnostics.Debug.WriteLine(a)
        'Next

        Dim output As String
        output = CType(securitiesCodes(0), String) + " " + Prices.Name + ",,,,," + vbCrLf _
                + """date"",""open"",""hight"",""low"",""close"",""power"",""End"""

        'ファイル削除
        System.IO.File.Delete(csv)
        outputCsv(output)

        For i = 0 To date_range
            If IsNothing(stock_array(i, 0)) Then
                Continue For
            End If
            '9101,,,,,
            '"date","open","hight","low","close","power","end"
            output = """" + stock_array(i, 0) + """,""" + stock_array(i, 1) + """,""" + stock_array(i, 2) + """,""" + stock_array(i, 3) _
                 + """,""" + stock_array(i, 4) + """,""" + stock_array(i, 5) + """,""" + stock_array(i, 6) + """"
            System.Diagnostics.Debug.WriteLine(output)
            outputCsv(output)
        Next

    End Sub

    Sub outputCsv(output)
        Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift_JIS") 'CSVファイルのエンコードを指定（今回はShift_JIS）
        Dim sr As New System.IO.StreamWriter(csv, True, enc) '書き込むファイルを開く
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

End Module