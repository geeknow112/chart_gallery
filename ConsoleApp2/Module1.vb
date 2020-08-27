Imports MySql.Data.MySqlClient
Module Module1

    'Dim csvPath As String = "D:\Users\z112\source\repos\ConsoleApp2\stock_data\"
    Dim csvPath As String = "C:\Users\r2d2\source\repos\chart_gallery\stock_data\"
    'ReadOnly csvPath As String = "C:\Users\r2d2\source\repos\chart_gallery\stock_data\"

    Sub Main()
        ConnectMySql()
    End Sub

    Sub MainBk()

        'MakeZip()
        'Exit Sub

        'GetCmdParam()

        Dim securitiesCodes() As Integer = {
            9101, 9104, 9107, 6326
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
            'Conn.Open()
            'data追加
            'Dim query = "INSERT INTO test VALUES (2, 'test2')"
            'Dim cmd As MySqlCommand = New MySqlCommand(query, Conn)
            'cmd.ExecuteNonQuery()
            'Conn.Close()

            'Conn.Open()
            'Dim q = "insert into s6326 values ('1989-02-21','985','990','970','979','1957000','979','1');"
            'Dim c As MySqlCommand = New MySqlCommand(q, Conn)
            'c.ExecuteNonQuery()

            'Dim q = "create table s9984 (id int not null auto_increment, date varchar(20) unique, open int, hight int, low int, close int, power int, End int, date_position int unique, rgdt timestamp, primary key(id)) comment='9984 SOFTBANK';"
            'Dim c As MySqlCommand = New MySqlCommand(q, Conn)
            'c.ExecuteNonQuery()
            'Conn.Close()

            Dim securitiesCodes() As Integer = {
                9101, 9104, 9107, '海運業
                9202, 'ＡＮＡＨＤ
                9301, '三菱倉
                9412, 'スカパーＪ
                9433, 'ＫＤＤＩ
                9437, 'ＮＴＴドコモ
                9984, 'ＳＢＧ
                9613, 'ＮＴＴデータ
                9432, 'ＮＴＴ
                9503, '関西電
                9502, '中部電
                9501, '東電ＨＤ
                9531, '東ガス
                9532, '大ガス
                2413, 'エムスリー
                9602, '東宝
                4704, 'トレンド
                4751, 'サイバー
                4689, 'ＺＨＤ
                4755, '楽天
                6178, '日本郵政
                9735, 'セコム
                2432, 'ディーエヌエ
                4324, '電通グループ
                9766, 'コナミＨＤ
                6098, 'リクルート
                8630, 'ＳＯＭＰＯ
                8750, '第一生命ＨＤ
                8795, 'Ｔ＆Ｄ
                8725, 'ＭＳ＆ＡＤ
                8766, '東京海上
                8697, '日本取引所
                8253, 'クレセゾン
                8830, '住友不
                8804, '東建物
                8801, '三井不
                3289, '東急不ＨＤ
                8802, '菱地所
                9022, 'ＪＲ東海
                9021, 'ＪＲ西日本
                9020, 'ＪＲ東日本
                9009, '京成
                9005, '東急
                9007, '小田急
                9008, '京王
                9001, '東武
                9062, '日通
                9064, 'ヤマトＨＤ
                1332, '日水
                1333, 'マルハニチロ
                1605, '国際石開帝石
                1925, 'ハウス
                1928, '積ハウス
                1808, '長谷工
                1803, '清水建
                1802, '大林組
                1801, '大成建
                1812, '鹿島
                1721, 'コムシスＨＤ
                1963, '日揮ＨＤ
                2502, 'アサヒ
                2501, 'サッポロＨＤ
                2002, '日清粉Ｇ
                2531, '宝ＨＬＤ
                2282, '日ハム
                2914, 'ＪＴ
                2802, '味の素
                2503, 'キリンＨＤ
                2871, 'ニチレイ
                2269, '明治ＨＤ
                2801, 'キッコマン
                3103, 'ユニチカ
                3401, '帝人
                3402, '東レ
                3101, '東洋紡
                3861, '王子ＨＤ
                3863, '日本紙
                4021, '日産化
                4183, '三井化学
                4005, '住友化
                4188, '三菱ケミＨＤ
                4911, '資生堂
                3407, '旭化成
                4042, '東ソー
                6988, '日東電
                3405, 'クラレ
                4061, 'デンカ
                4208, '宇部興
                4272, '日化薬
                4004, '昭電工
                4631, 'ＤＩＣ
                4043, 'トクヤマ
                4901, '富士フイルム
                4452, '花王
                4063, '信越化
                4578, '大塚ＨＤ
                4519, '中外薬
                4502, '武田
                4503, 'アステラス
                4506, '大日本住友
                4151, '協和キリン
                4568, '第一三共
                4507, '塩野義
                4523, 'エーザイ
                5020, 'ＥＮＥＯＳ
                5019, '出光興産
                5108, 'ブリヂストン
                5101, '浜ゴム
                5233, '太平洋セメ
                5202, '板硝子
                5301, '東海カ
                5201, 'ＡＧＣ
                5214, '日電硝
                5333, 'ガイシ
                5232, '住友大阪
                5332, 'ＴＯＴＯ
                5411, 'ＪＦＥ
                5401, '日本製鉄
                5406, '神戸鋼
                5541, '大平金
                5714, 'ＤＯＷＡ
                5713, '住友鉱
                5803, 'フジクラ
                5706, '三井金
                5901, '洋缶ＨＤ
                5703, '日軽金ＨＤ
                5801, '古河電
                5802, '住友電
                5707, '東邦鉛
                5711, '三菱マ
                3436, 'ＳＵＭＣＯ
                6367, 'ダイキン
                6305, '日立建機
                6326, 'クボタ
                6301, 'コマツ
                5631, '日製鋼
                6113, 'アマダ
                7013, 'ＩＨＩ
                7011, '三菱重
                6472, 'ＮＴＮ
                6473, 'ジェイテクト
                7004, '日立造
                6471, '日精工
                6302, '住友重
                6361, '荏原
                6103, 'オークマ
                8035, '東エレク
                6702, '富士通
                6762, 'ＴＤＫ
                6701, 'ＮＥＣ
                6902, 'デンソー
                6479, 'ミネベア
                6841, '横河電
                6506, '安川電
                6770, 'アルプスアル
                6503, '三菱電
                6976, '太陽誘電
                3105, '日清紡ＨＤ
                7735, 'スクリン
                6752, 'パナソニック
                6674, 'ＧＳユアサ
                7752, 'リコー
                7751, 'キヤノン
                6703, 'ＯＫＩ
                6724, 'エプソン
                6952, 'カシオ
                6501, '日立
                6504, '富士電機
                6971, '京セラ
                6857, 'アドテスト
                6758, 'ソニー
                6645, 'オムロン
                6954, 'ファナック
                7003, '三井Ｅ＆Ｓ
                7012, '川重
                7269, 'スズキ
                7261, 'マツダ
                7211, '三菱自
                7270, 'ＳＵＢＡＲＵ
                7201, '日産自
                7202, 'いすゞ
                7272, 'ヤマハ発
                7205, '日野自
                7203, 'トヨタ
                7267, 'ホンダ
                7733, 'オリンパス
                4902, 'コニカミノル
                7762, 'シチズン
                7731, 'ニコン
                4543, 'テルモ
                7832, 'バンナムＨＤ
                7911, '凸版
                7912, '大日印
                7951, 'ヤマハ
                8053, '住友商
                8058, '三菱商
                8031, '三井物
                8002, '丸紅
                8015, '豊田通商
                2768, '双日
                8001, '伊藤忠
                9983, 'ファストリ
                3086, 'Ｊフロント
                3099, '三越伊勢丹
                8233, '高島屋
                3382, 'セブン＆アイ
                8252, '丸井Ｇ
                8267, 'イオン
                8028, 'ファミマ
                7186, 'コンコルディ
                8306, '三菱ＵＦＪ
                8411, 'みずほＦＧ
                8331, '千葉銀
                8308, 'りそなＨＤ
                8354, 'ふくおかＦＧ
                8304, 'あおぞら銀
                8355, '静岡銀
                8316, '三井住友ＦＧ
                8309, '三井住友トラ
                8303, '新生銀
                8628, '松井
                8604, '野村
                8601 '大和
            }

            Dim Prices As New ActiveMarket.Prices
            Dim Calendar As New ActiveMarket.Calendar
            Dim hash As New Hashtable
            Dim date_position As Integer
            Dim date_range As Integer
            Dim code As String

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
                    stock_array(i, 1) = Math.Floor(Prices.Open(date_position))
                    stock_array(i, 2) = Math.Floor(Prices.High(date_position))
                    stock_array(i, 3) = Math.Floor(Prices.Low(date_position))
                    stock_array(i, 4) = Math.Floor(Prices.Close(date_position))
                    stock_array(i, 5) = Math.Floor(Prices.Volume(date_position) * 1000)
                    stock_array(i, 6) = Math.Floor(Prices.Close(date_position))
                    stock_array(i, 7) = date_position
                Next

                hash.Add(securitiesCodes(securitiesCode), Prices.Name)
                code = securitiesCodes(securitiesCode).ToString
                Conn.Open()
                Dim q = "create table IF NOT EXISTS s" + code + " (id int not null auto_increment, date varchar(20) unique, open int, hight int, low int, close int, power int, End int, rgdt timestamp, primary key(id)) comment='" + code + " " + Prices.Name + "';"
                Dim c As MySqlCommand = New MySqlCommand(q, Conn)
                c.ExecuteNonQuery()

                'Dim t = "truncate s" + code + ";"
                'Dim tr As MySqlCommand = New MySqlCommand(t, Conn)
                'tr.ExecuteNonQuery()

                'max_date取得
                Dim query = "select max(date) as max_date from stocks.s" + code + " limit 1"
                Dim cmd As MySqlCommand = New MySqlCommand(query, Conn)
                Dim data As MySqlDataReader = cmd.ExecuteReader
                Dim max_date As String
                '結果を表示
                While data.Read()
                    'Console.WriteLine(data("close"))
                    max_date = data("max_date")
                End While
                Conn.Close()
                Conn.Open()
                For i = 0 To date_range
                    If IsNothing(stock_array(i, 0)) Then
                        Continue For
                    End If

                    If stock_array(i, 0) = max_date Then
                        i = i + 1
                        If IsNothing(stock_array(i, 0)) Then
                            Continue For
                        End If
                        Dim ins = "insert into s" + code + " values (null, '" + stock_array(i, 0) + "','" + stock_array(i, 1) + "','" + stock_array(i, 2) + "','" + stock_array(i, 3) _
                        + "','" + stock_array(i, 4) + "','" + stock_array(i, 5) + "','" + stock_array(i, 6) + "', now());"

                        Dim insert As MySqlCommand = New MySqlCommand(ins, Conn)
                        insert.ExecuteNonQuery()
                        Exit For
                    End If
                Next
                Conn.Close()
            Next
        End Using
    End Sub
End Module