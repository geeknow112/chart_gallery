Module StockCodes
    Public Function SetNikkei225() As Integer()
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
                8601, '大和
                1001, 1007
            }
        SetNikkei225 = securitiesCodes
    End Function

    Public Function SetJpx400() As Integer()
        Dim securitiesCodes() As Integer = {
            1332, '日水
            1333, 'マルハニチロ
            1605, 'ＩＮＰＥＸ
            1719, '安藤ハザマ
            1720, '東急建
            1721, 'コムシスＨＤ
            1766, '東建コーポ
            1801, '大成建
            1802, '大林組
            1803, '清水建
            1808, '長谷工
            1812, '鹿島
            1820, '西松建
            1821, '三井住友建
            1824, '前田建
            1860, '戸田建
            1861, '熊谷組
            1878, '大東建
            1881, 'ＮＩＰＰＯ
            1893, '五洋建
            1911, '住友林
            1925, 'ハウス
            1928, '積ハウス
            1951, '協エクシオ
            1959, '九電工
            2201, '森永
            2222, '寿スピリッツ
            2229, 'カルビー
            2264, '森永乳
            2267, 'ヤクルト
            2269, '明治ＨＤ
            2282, '日ハム
            2502, 'アサヒ
            2503, 'キリンＨＤ
            2587, 'サントリＢＦ
            2593, '伊藤園
            2801, 'キッコマン
            2802, '味の素
            2809, 'キユーピー
            2811, 'カゴメ
            2815, 'アリアケ
            2871, 'ニチレイ
            2875, '東洋水
            2897, '日清食ＨＤ
            2914, 'ＪＴ
            3401, '帝人
            3402, '東レ
            3861, '王子ＨＤ
            3405, 'クラレ
            3407, '旭化成
            4004, '昭電工
            4005, '住友化
            4021, '日産化
            4042, '東ソー
            4043, 'トクヤマ
            4061, 'デンカ
            4063, '信越化
            4088, 'エアウォータ
            4091, '日本酸素ＨＤ
            4182, '菱ガス化
            4183, '三井化学
            4185, 'ＪＳＲ
            4188, '三菱ケミＨＤ
            4189, 'ＫＨネオケム
            4202, 'ダイセル
            4204, '積水化
            4206, 'アイカ
            4208, '宇部興
            4403, '日油
            4452, '花王
            4612, '日本ペＨＤ
            4613, '関西ペ
            4631, 'ＤＩＣ
            4911, '資生堂
            4912, 'ライオン
            4921, 'ファンケル
            4922, 'コーセー
            4927, 'ポーラＨＤ
            6988, '日東電
            4151, '協和キリン
            4502, '武田
            4503, 'アステラス
            4506, '大日本住友
            4507, '塩野義
            4516, '日本新薬
            4519, '中外薬
            4521, '科研薬
            4523, 'エーザイ
            4527, 'ロート
            4528, '小野薬
            4530, '久光薬
            4536, '参天薬
            4568, '第一三共
            4578, '大塚ＨＤ
            4887, 'サワイＧＨＤ
            4967, '小林製薬
            5019, '出光興産
            5020, 'ＥＮＥＯＳ
            5021, 'コスモＨＤ
            5101, '浜ゴム
            5105, 'ＴＯＹＯ
            5108, 'ブリヂストン
            5110, '住友ゴ
            5201, 'ＡＧＣ
            5233, '太平洋セメ
            5301, '東海カ
            5332, 'ＴＯＴＯ
            5333, 'ガイシ
            5334, '特殊陶
            5393, 'ニチアス
            5401, '日本製鉄
            5411, 'ＪＦＥ
            3436, 'ＳＵＭＣＯ
            5713, '住友鉱
            5801, '古河電
            5802, '住友電
            5857, 'アサヒＨＤ
            5929, '三和ＨＤ
            5947, 'リンナイ
            5631, '日製鋼
            6005, '三浦工
            6113, 'アマダ
            6134, 'ＦＵＪＩ
            6136, 'ＯＳＧ
            6141, 'ＤＭＧ森精機
            6201, '豊田織
            6268, 'ナブテスコ
            6273, 'ＳＭＣ
            6301, 'コマツ
            6302, '住友重
            6305, '日立建機
            6326, 'クボタ
            6367, 'ダイキン
            6383, 'ダイフク
            6432, '竹内製作所
            6465, 'ホシザキ
            6471, '日精工
            6481, 'ＴＨＫ
            7011, '三菱重
            7013, 'ＩＨＩ
            6448, 'ブラザー
            6479, 'ミネベア
            6501, '日立
            6503, '三菱電
            6504, '富士電機
            6506, '安川電
            6586, 'マキタ
            6588, '東芝テック
            6594, '日電産
            6645, 'オムロン
            6670, 'ＭＣＪ
            6701, 'ＮＥＣ
            6702, '富士通
            6723, 'ルネサス
            6724, 'エプソン
            6728, 'アルバック
            6750, 'エレコム
            6752, 'パナソニック
            6753, 'シャープ
            6754, 'アンリツ
            6758, 'ソニーＧ
            6762, 'ＴＤＫ
            6770, 'アルプスアル
            6841, '横河電
            6845, 'アズビル
            6849, '日本光電
            6856, '堀場製
            6857, 'アドテスト
            6861, 'キーエンス
            6869, 'シスメックス
            6877, 'ＯＢＡＲＡＧ
            6902, 'デンソー
            6920, 'レーザーテク
            6923, 'スタンレー
            6952, 'カシオ
            6954, 'ファナック
            6965, 'ホトニクス
            6971, '京セラ
            6976, '太陽誘電
            6981, '村田製
            7735, 'スクリン
            7751, 'キヤノン
            8035, '東エレク
            3116, 'トヨタ紡織
            7202, 'いすゞ
            7203, 'トヨタ
            7205, '日野自
            7259, 'アイシン
            7261, 'マツダ
            7267, 'ホンダ
            7269, 'スズキ
            7270, 'ＳＵＢＡＲＵ
            7272, 'ヤマハ発
            7276, '小糸製
            7282, '豊田合
            7313, 'ＴＳテック
            7309, 'シマノ
            4543, 'テルモ
            6146, 'ディスコ
            7701, '島津
            7717, 'Ｖテク
            7729, '東京精
            7731, 'ニコン
            7733, 'オリンパス
            7741, 'ＨＯＹＡ
            7747, '朝日インテク
            7832, 'バンナムＨＤ
            7846, 'パイロット
            7951, 'ヤマハ
            7956, 'ピジョン
            7988, 'ニフコ
            2768, '双日
            2784, 'アルフレッサ
            3038, '神戸物産
            3107, 'ダイワボＨＤ
            3167, 'ＴＯＫＡＩ
            3360, 'シップＨＤ
            7458, '第一興商
            7459, 'メディパル
            7575, '日本ライフＬ
            8001, '伊藤忠
            8002, '丸紅
            8015, '豊田通商
            8020, '兼松
            8031, '三井物
            8053, '住友商
            8058, '三菱商
            8088, '岩谷産
            8111, 'ゴルドウイン
            8113, 'ユニチャーム
            8283, 'ＰＡＬＴＡＣ
            9810, '日鉄物産
            9962, 'ミスミＧ
            2651, 'ローソン
            2670, 'ＡＢＣマート
            2782, 'セリア
            3048, 'ビックカメラ
            3064, 'モノタロウ
            3086, 'Ｊフロント
            3088, 'マツキヨＨＤ
            3092, 'ＺＯＺＯ
            3141, 'ウエルシア
            3148, 'クリエイトＳ
            3349, 'コスモス薬品
            3382, 'セブン＆アイ
            3391, 'ツルハＨＤ
            3549, 'クスリアオキ
            7419, 'ノジマ
            7453, '良品計画
            7532, 'パンパシＨＤ
            7564, 'ワークマン
            7649, 'スギＨＤ
            8252, '丸井Ｇ
            8267, 'イオン
            8273, 'イズミ
            8279, 'ヤオコー
            8282, 'ケーズＨＤ
            9627, 'アインＨＤ
            9843, 'ニトリＨＤ
            9983, 'ファストリ
            9989, 'サンドラッグ
            7167, 'めぶきＦＧ
            7186, 'コンコルディ
            8303, '新生銀
            8304, 'あおぞら銀
            8306, '三菱ＵＦＪ
            8308, 'りそなＨＤ
            8309, '三井住友トラ
            8316, '三井住友ＦＧ
            8331, '千葉銀
            8354, 'ふくおかＦＧ
            8410, 'セブン銀
            8411, 'みずほＦＧ
            7148, 'ＦＰＧ
            7164, '全国保証
            8424, '芙蓉リース
            8439, '東京センチュ
            8473, 'ＳＢＩ
            8570, 'イオンＦＳ
            8572, 'アコム
            8585, 'オリコ
            8591, 'オリックス
            8593, '三菱ＨＣキャ
            8697, '日本取引所
            8601, '大和
            8604, '野村
            8630, 'ＳＯＭＰＯ
            8725, 'ＭＳ＆ＡＤ
            8750, '第一生命ＨＤ
            8766, '東京海上
            8795, 'Ｔ＆Ｄ
            2337, 'いちご
            3003, 'ヒューリック
            3231, '野村不ＨＤ
            3288, 'オープンＨ
            3289, '東急不ＨＤ
            3291, '飯田ＧＨＤ
            4666, 'パーク２４
            8801, '三井不
            8802, '菱地所
            8804, '東建物
            8830, '住友不
            8850, 'スターツ
            8905, 'イオンモール
            9706, '日本空港ビル
            9001, '東武
            9003, '相鉄ＨＤ
            9005, '東急
            9007, '小田急
            9008, '京王
            9009, '京成
            9020, 'ＪＲ東日本
            9021, 'ＪＲ西日本
            9022, 'ＪＲ東海
            9024, '西武ＨＤ
            9041, '近鉄ＧＨＤ
            9042, '阪急阪神
            9044, '南海電
            9045, '京阪ＨＤ
            9048, '名鉄
            9142, 'ＪＲ九州
            9062, '日通
            9064, 'ヤマトＨＤ
            9065, '山九
            9086, '日立物流
            9201, 'ＪＡＬ
            9202, 'ＡＮＡＨＤ
            3738, 'ティーガイア
            9432, 'ＮＴＴ
            9433, 'ＫＤＤＩ
            9435, '光通信
            9613, 'ＮＴＴデータ
            9984, 'ＳＢＧ
            9502, '中部電
            9503, '関西電
            9504, '中国電
            9506, '東北電
            9508, '九州電
            9509, '北海電
            9513, 'Ｊパワー
            9531, '東ガス
            9532, '大ガス
            2121, 'ミクシィ
            2127, '日本Ｍ＆Ａ
            2146, 'ＵＴ
            2175, 'エスエムエス
            2181, 'パーソルＨＤ
            2317, 'システナ
            2327, 'ＮＳＳＯＬ
            2331, 'ＡＬＳＯＫ
            2371, 'カカクコム
            2379, 'ディップ
            2412, 'ベネ・ワン
            2413, 'エムスリー
            2427, 'アウトソシン
            2433, '博報堂ＤＹ
            2702, 'マクドナルド
            3197, 'すかいらーく
            3543, 'コメダ
            3563, 'Ｆ＆ＬＣ
            3626, 'ＴＩＳ
            3635, 'コーテクＨＤ
            3659, 'ネクソン
            3765, 'ガンホー
            3769, 'ＧＭＯ−ＰＧ
            3932, 'アカツキ
            4307, '野村総研
            4324, '電通グループ
            4348, 'インフォコム
            4661, 'ＯＬＣ
            4684, 'オービック
            4686, 'ジャスト
            4689, 'ＺＨＤ
            4704, 'トレンド
            4716, '日本オラクル
            4732, 'ＵＳＳ
            4739, 'ＣＴＣ
            4755, '楽天グループ
            4768, '大塚商会
            4816, '東映アニメ
            4819, 'Ｄガレージ
            4849, 'エンジャパン
            6028, 'テクノプロＨ
            6035, 'ＩＲジャパン
            6098, 'リクルート
            6532, 'ベイカレント
            7550, 'ゼンショＨＤ
            7974, '任天堂
            8056, 'ユニシス
            8876, 'リログループ
            9602, '東宝
            9603, 'ＨＩＳ
            9678, 'カナモト
            9684, 'スクエニＨＤ
            9697, 'カプコン
            9719, 'ＳＣＳＫ
            9735, 'セコム
            9744, 'メイテック
            9766 'コナミＨＤ
        }
        SetJpx400 = securitiesCodes
    End Function

End Module
