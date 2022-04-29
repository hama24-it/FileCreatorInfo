#-----------------------------------------------------------
# ■目的
# ・ファイルの作成者情報を確認し、管理外で作成されたファイルがないか確認

# ■対象の場所
# ・指定したフォルダ（サブフォルダを含む）の配下にあるファイル

# ■対象のファイル拡張子
# ・"*.xls*"
# ・"*.doc*"

# ■言語設計
# 01.対象のフォルダの情報を取得
# 02.foreach ループにより取得したファイル情報を順次処理
# 03.ファイルの場合の処理を実行
#	・if 条件分岐により、指定したファイル拡張子の場合に処理を実行
#	・ファイル作成者情報を取得し変数に格納
# 04.変数に格納した情報を整形し順次に表示
#-----------------------------------------------------------
# [Code]
# ■シェルオブジェクトを作成 
$shell = New-Object -ComObject Shell.Application

# ■処理対象のフォルダ
$targetFolder = "C:\Users\hama24\Desktop\test"

# ■ファイル拡張子の指定
[string[]]$include = @("*.xls*", "*.doc*")

# ■変数にフォルダの情報を格納（-Recurse：サブフォルダを含む）
$itemList = Get-ChildItem $targetFolder -Recurse

foreach($item in $itemList)
{
    # ■PSIsContainer プロパティによりフォルダ・ファイルの判定
    if($item.PSIsContainer)
    {
        # ■フォルダの場合の処理
	# 
    }
    else
    {
	# ■フォルダの指定
	$shellFolder = $shell.NameSpace($item.Directory.FullName)

	# ■ファイルの種類を指定
	$File_01 = (Get-ChildItem $shellFolder.parseName($item.Name).Path -include $include).Name

	# ■文字列がNULL・空白か判定
	if([string]::IsNullOrEmpty($File_01))
	{
	# 空の場合の処理
	# 
	}
	else
	{
	# ■指定したファイルが存在する場合の処理
	$shellFile = $shellFolder.parseName($File_01)

        # ■ファイルの場合の処理
	# ■文字列を変数に格納
	$Create = "20(作成者)_"
	# ■ファイルの作成者情報を変数に格納
	$CreateData = $shellFolder.GetDetailsOf($shellFile, 20)
	Write-output "$Create$CreateData,$item"
	}
    }
} 
#-----------------------------------------------------------
