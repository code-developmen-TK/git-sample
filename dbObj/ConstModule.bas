Option Compare Database
Option Explicit
'------------------------------------------------------------------------------
'システム使用する定数の定義モジュール
'-------------------------------------------------------------------------------

'ADO　カーソルタイプ定数
Public Const adOpenForwardOnly As Integer = 0 '前方スクロールカーソル　レコードセットの先頭から末尾に向かって移動することができる
Public Const adOpenKeyset As Integer = 1 'キーセットカーソル　レコードセットの全ての方向に移動することができる。他のユーザーが更新したレコードは参照することができますが、追加、削除したレコードは参照できない。
Public Const adOpenDynamic As Integer = 2 '動的カーソル　レコードセットの全ての方向に移動することができる。他のユーザーが追加、更新、削除したレコードも参照することができる。
Public Const adOpenStatic As Integer = 3 '静的カーソル　レコードセットの全ての方向に移動することができる。他のユーザーによる追加、更新、削除は参照することができない。

'ADO　ロックタイプ定数
Public Const adLockReadOnly As Integer = 0 '読み取り専用 データの更新・追加・削除はできない
Public Const adLockPessimistic As Integer = 1 '排他的ロック　編集直後にレコードをロック
Public Const adLockOptimistic As Integer = 2 '共有的ロック　Updateメソッドを呼び出した場合にのみ、共有的ロック
Public Const adLockBatchOptimistic As Integer = 3 '複数のレコードをバッチ更新

'ADO オプション定数
Public Const adCmdText As Integer = 1 'コマンドまたはストアド プロシージャのテキスト定義として評価します。

'ADO カーソルロケーション定数
Public Const adUseClient As Integer = 3 'ローカル カーソル
Public Const adUseServer As Integer = 2 'データプロバイダーカーソル

'ADO スキーマ情報定数
Public Const adSchemaColumns As Integer = 4      'カラムの定義
Public Const adSchemaTables As Long = 20         'テーブルの定義
Public Const adSchemaPrimaryKeys As Integer = 28 '主キーの定義
 
'ADO データ型定数
Public Const adBoolean As Integer = 11         '真偽型
Public Const adUnsignedTinyInt As Integer = 17 'バイト型（符号なし）
Public Const adSmallInt As Integer = 2         '整数型（符号付き）
Public Const adInteger As Integer = 3          '長整数型（符号付き）
Public Const adCurrency As Integer = 6         '通貨型（符号付き）
Public Const adSingle As Integer = 4           '単精度浮動小数点型
Public Const adDouble As Integer = 5           '倍精度浮動小数点型
Public Const adDate As Integer = 7             '日付/時刻型
Public Const adWChar As Integer = 130          '文字列型
Public Const adLongVarBinary As Integer = 205  'ロングバイナリ型

'ADOレコードセットの状態を表す定数
Public Const adStateClosed As Integer = 0     'オブジェクトは閉じていることを示す。
Public Const adStateOpen As Integer = 1       'オブジェクトは開いていることを示す。
Public Const adStateConnecting As Integer = 2 'オブジェクトは接続していることを示す。
Public Const adStateExecuting As Integer = 4  'オブジェクトはコマンドを実行していることを示す。
Public Const adStateFetching As Integer = 8   'オブジェクトの行が取得されていることを示す。

'Officeオブジェクト ファイル・フォルダ選択ダイアログで使う定数
Public Const msoFileDialogFilePicker As Integer = 3 'ファイルを選択する場合
Public Const msoFileDialogFolderPicker As Integer = 4 'フォルダを選択する場合

'テストデータシステム全体で使う設定値を定数にセット
Public Const SYSTEM_FILE_NAME As String = "テストデータシステム.accdb"
Public Const CONNECT_DATABASE_NAME As String = "テストデータベース.accdb"
Public Const DEFAULT_TEXT_FOLDER_PATH As String = "D:\VBA開発\access\テキストデータ"
'Public Const DATABASE_PROVIDOR As String = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=D:\VBA開発\access\テストデータベース.accdb"

'システムで使用するテーブル名の定数
'Public Const TEMPORARY_TABLE_NAME As String = "T_受入物情報データ読み込み用テーブル"
'Public Const LOADING_TABLE_NAME As String = "T_受入物情報データテーブル"
Public Const ACCEPTANCE_DATA_DABLE As String = "T_受入物データテーブル"
Public Const ACCEPTANCE_INSPECT_DATA_DABLE As String = "T_受入物検査データテーブル"
'Public Const ACCEPTANCE_INSPECT_DATA_DABLE As String = "T_受入物外内容器対応データテーブル"
Public Const ACCEPTANCE_INNER_OUTER_CORRESPONDENCE_TABLE As String = "T_受入物外内容器対応データテーブル"
Public Const ACCEPTANCE_COMPOSITION_DATA_TABLE As String = "T_受入物組成データテーブル"
Public Const ACCEPTANCE_SLIP_MANAGEMENT_TABLE As String = "T_受入物伝票管理"
Public Const ACCEPTANCE_SLIP_DETALS_TABLE As String = "T_受入物伝票詳細情報データテーブル"
Public Const TREATMENT_TABLE As String = "T_処理テーブル"
Public Const TREATMENT_DATE_CORRESPOMDENCE As String = "T_処理日対応テーブル"

'Excel操作関係の定数
Public Const xlDown As Integer = -4121                         '下へ
Public Const xlToLeft As Integer = 4159                        '左へ
Public Const xlToRight As Integer = 4161                       '右へ
Public Const xLUp As Integer = -4162                           '上へ

Public Const HISTORY_SHEET_FIRST_ROW As Long = 8               '読み込むデータの入っているシートの最初の行。
Public Const HISTORY_SHEET_FIRST_CLUMN As Long = 1             '読み込むデータの入っているシートの最初の列。
Public Const HISTORY_SHEET_CLUMNS As Long = 34                 '読み込むデータの入っているシートの列数。
Public Const HISTORY_SHEET_NUMBER1 As String = "カテゴリ１"    '読み込むデータの入っているシート１の名前
Public Const HISTORY_SHEET_NUMBER2 As String = "カテゴリ2"     '読み込むデータの入っているシート２の名前
Public Const SPRIT_ROW As Long = 21                            '読み込むデータの入っているシート（配列）の左右分割位置

'Public Const SOTOYOUKI_NUMBAER As Long = 4                     '読み込むデータの入っているシートの外容器番号の入っている列

'Public Const UCHIYOUKI_NUMBER1 As Long = 9                     '読み込むデータの入っているシートの内容器番号の入っている列
'Public Const CONTENT1 As Long = 10                             '読み込むデータの入っているシートの内容物の入っている列
'Public Const WEIGHT1 As Long = 12                              '読み込むデータの入っているシートの重量の入っている列
'Public Const DOSE1 As Long = 13                                '読み込むデータの入っているシートの量の入っている列
'Public Const ORENGE1 As Long = 14                                '読み込むデータの入っているシートの量の入っている列
'Public Const GREEN1 As Long = 15                                '読み込むデータの入っているシートの量の入っている列
'Public Const BLACK1 As Long = 16                                '読み込むデータの入っているシートの量の入っている列

'Public Const UCHIYOUKI_NUMBER2 As Long = 21                    '読み込むデータの入っているシートの内容器番号の入っている列
'Public Const CONTENT2 As Long = 24                             '読み込むデータの入っているシートの内容物の入っている列
'Public Const WEIGHT2 As Long = 23                              '読み込むデータの入っているシートの重量の入っている列
'Public Const DOSE2 As Long = 25                                '読み込むデータの入っているシートの量の入っている列
'Public Const ORENGE2 As Long = 26                                '読み込むデータの入っているシートの量の入っている列
'Public Const GREEN2 As Long = 27                                '読み込むデータの入っているシートの量の入っている列
'Public Const BLACK2 As Long = 28                                '読み込むデータの入っているシートの量の入っている列


Public Const DEFAULT_FOLDER As String = "D:\プログラム等開発\excel\履歴管理データ" '最初に開くフォルダを指定
'Public Const PROCESSING_DATE As Long = 32 '処理日の入った列
'Public Const TREATMENT_DATE_COLUMN As Long = 32 '処理日の入った列
 
'Excel履歴管理データの列番号の定数
Public Const cst缶数 As String = 1
Public Const cst記号 As String = 2
Public Const cst番号 As String = 3
Public Const cst外容器番号 As String = 4
Public Const cst封入日 As String = 5
Public Const cstW量 As String = 6
Public Const cst収納数 As String = 7
Public Const cst部屋 As String = 8
Public Const cst内容器番号1 As String = 9
Public Const cst内容物1 As String = 10
Public Const cst種別 As String = 11
Public Const cst重量1 As String = 12
Public Const cst染料1 As String = 13
Public Const cstオレンジ1 As String = 14
Public Const cstミドリ1 As String = 15
Public Const cstクロ1 As String = 16
Public Const cst前処理 As String = 17
Public Const cst判定 As String = 18
Public Const cst戻し As String = 19
Public Const cst高染料 As String = 20
Public Const cst内容器番号2 As String = 21
Public Const cst分割 As String = 22
Public Const cst重量2 As String = 23
Public Const cst内容物2 As String = 24
Public Const cst染料2 As String = 25
Public Const cstオレンジ2 As String = 26
Public Const cstミドリ2 As String = 27
Public Const cstクロ2 As String = 28
Public Const cst処理可 As String = 29
Public Const cstブランク As String = 30
Public Const cst保留 As String = 31
Public Const cst処理日 As String = 32
Public Const cst処理物バッチ番号 As String = 33
Public Const cst備考 As String = 34
Public Const cst分割後の内容器番号位置 = 1

'使用テーブル名の一覧
'MT_種類
'MT_場所
'MT_内容器種別
'T_受入物伝票管理
'T_受入物組成
'T_受入物伝票詳細情報
'T_受入物検査
'T_受入物容器対応
'T_受入物情報
'T_処理
'T_処理日対応
'T_処理記録
'T_処理物収納情報
'T_処理物外容器封入
'T_処理物容器情報
'T_払出容器情報