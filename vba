''----------------------------------------------------------------------------------------
''--
''-- AM_USERテーブルの更新
''--
''-- 2016.09.09 作成
''--
''----------------------------------------------------------------------------------------

Option Explicit
On Error Resume Next

Dim oParam      ''-- パラメータ
Dim objConn     ''-- ADO Connector
Dim objRS       ''-- ADO RecordSet

Dim ConnStr     ''-- 接続文字列
Dim SERVER      ''-- 接続サーバ
Dim DBS         ''-- 接続DB
Dim UID         ''-- 接続User Id
Dim PWD         ''-- Password

Dim sqlStr      ''-- SQL構文
Dim DtCnt
Dim setValue
Dim CsvPath     ''-- CSVファイルパス
Dim CsvPath2    ''-- CSVファイルパス(BULK INSERT)

Dim myArrayList ''-- 更新SQLリスト

Dim objFSO      ''-- FileSystemObject

Set oParam = WScript.Arguments

''-- 引数確認 --
If oParam.Count < 5 then
   WScript.Echo "パラメータエラー"
   WScript.Echo "プログラム名 :AM_USER.vbs"
   WScript.Echo "  第1パラメータ:接続サーバ"
   WScript.Echo "  第2パラメータ:接続DB"
   WScript.Echo "  第3パラメータ:UserID"
   WScript.Echo "  第4パラメータ:Password"
   WScript.Echo "  第5パラメータ:CSVファイルパス"
   WScript.Quit
End If

''-- 引数セット
SERVER  = oParam(0)
DBS     = oParam(1)
UID     = oParam(2)
PWD     = oParam(3)
CsvPath = oParam(4)

''-- CSVファイルの存在確認
Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FileExists(CsvPath) = False Then
   WScript.Echo "指定CSVファイル [" & CsvPath & "] が見つかりません。" & vbcrlf
   WScript.Quit
Else
   ''-- 参照ファイルの絶対パスを取得
   CsvPath2 = objFSO.GetAbsolutePathName(CsvPath)
End If
Set objFSO = Nothing

''-- 接続文字列準備
ConnStr = "Provider=sqloledb;Data Source=" & SERVER & ";" & _
          "Initial Catalog=" & DBS & ";" & _
          "User ID=" & UID & ";" & _
          "Password=" & PWD & ";"

''-- データベース接続
Set objConn = CreateObject("ADODB.Connection")

objConn.ConnectionString = ConnStr
objConn.Open
objConn.CursorLocation = 3 ' クライアントサイドカーソルに変更

''-- エラー確認
If Err.Number <> 0 then
   WScript.Echo "エラー：[" & Err.Number & "] " & Err.Description
   WScript.Echo "接続[NG]" & vbcrlf
   WScript.Quit
Else
   WScript.Echo "接続[OK]" & vbcrlf
End If

''-- 接続先DBを指定
sqlStr = "USE " & DBS
objConn.Execute sqlStr
If Err.Number <> 0 Then
   WScript.Echo "エラー：[" & Err.Number & "] " & Err.Description
   WScript.Quit
End If

''-- 中間テーブルの処理(存在した場合は消す) --
sqlStr = "IF OBJECT_ID(N'tempdb..##AM_USER', N'U') IS NOT NULL" & vbcrlf & _
         "DROP TABLE ##AM_USER"
objConn.Execute sqlStr
If Err.Number <> 0 Then
   WScript.Echo "エラー：[" & Err.Number & "] " & Err.Description
   WScript.Quit
End If

''-- 中間テーブル作成 --
sqlStr = "CREATE TABLE ##AM_USER (" & _
          "UserID varchar(10)," & _
          "PassCD varchar(15)," & _
          "PassHenkoDT datetime," & _
          "PassHenkoUserID varchar(10)," & _
          "RkWebKadoKBN char(1)," & _
          "WebKadoStopDT datetime," & _
          "WebKadoStopUserID varchar(10)," & _
          "KyoikuEndFLG char(1)," & _      
         ")"
objConn.Execute sqlStr
If Err.Number <> 0 Then
   WScript.Echo "エラー：[" & Err.Number & "] " & Err.Description
   WScript.Quit
End If

''-- 中間テーブルへのデータ投入 --
sqlStr = "BULK INSERT ##AM_USER from '" & CsvPath2 & "' WITH( fieldterminator = '~~~', rowterminator = '@@\r' )"
objConn.Execute sqlStr
If Err.Number <> 0 Then
   WScript.Echo "エラー：[" & Err.Number & "] " & Err.Description
   WScript.Quit
End If

''-- BULK INSERTしたテーブルから本テーブルを更新する --
sqlStr = "UPDATE rkwebopen1.AM_USER " & vbcrlf & _
         "SET UserID=B.SUID" & vbcrlf & _
         "   ,PassCD=B.SPASS " & vbcrlf & _
         "   ,PassHenkoDT=B.SPASSDATE " & vbcrlf & _
         "   ,PassHenkoUserID=B.SPUSER " & vbcrlf & _
         "   ,RkWebKadoKBN=B.KK " & vbcrlf & _      
         "   ,WebKadoStopDT=B.KKSTOPDATE " & vbcrlf & _        
         "   ,WebKadoStopUserID=B.KS " & vbcrlf & _        
         "   ,KyoikuEndFLG=B.EF " & vbcrlf & _
         "FROM rkwebopen1.AM_USER A" & vbcrlf & _
         "     INNER JOIN ##AM_USER B" & vbcrlf & _
         "     ON A.UserID=B.UserID " & vbcrlf & _ 
         "WHERE A.KK=B.KK"
objConn.Execute sqlStr
If Err.Number <> 0 Then
   WScript.Echo "エラー：[" & Err.Number & "] " & Err.Description
   WScript.Echo sqlStr
   WScript.Quit
End If


Set objRS = CreateObject("ADODB.Recordset")
''-- スイッチング完了時の更新内容 --
sqlStr = "SELECT @@ROWCOUNT"
''-- 検索実行
objRS.Open sqlStr, objConn
''-- エラー確認
If Err.Number <> 0 then
   WScript.Echo "エラー：[" & Err.Number & "] " & Err.Description
   WScript.Echo "検索[NG]" & vbcrlf
   WScript.Quit
End If

Do Until objRS.EOF
     DtCnt = objRS(0).Value
     objRS.MoveNext
Loop

objRS.Close
Set objRS = Nothing
WScript.Echo vbcrlf & "処理件数:" & DtCnt & vbcrlf

objConn.Close
WScript.Echo "切断[OK]"
Set objConn = Nothing

