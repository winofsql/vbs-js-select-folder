Set Shell = CreateObject( "Shell.Application" )
Set WSHShell = CreateObject("WScript.Shell")
' ダイアログ表示
Set objFolder = Shell.BrowseForFolder( 0, "フォルダ選択", 11, 0 )
' キャンセル
if objFolder is nothing then
    WScript.Quit
end if
if not objFolder.Self.IsFileSystem then
    WSHShell.Popup "ファイルシステムではありません"
    WScript.Quit
end if

WScript.Echo objFolder.Self.Path

strItems = ""

' フォルダ内のコレクションを取得
Set objFolderItems = objFolder.Items()

' コレクションを列挙
nFiles = objFolderItems.Count
For I = 0 to nFiles - 1
    Set objItem = objFolderItems.Item(I)
    if objItem.isFolder then
        strItems = strItems & "[Folder] " & objItem.Name & vbCrLf
    else
        strItems = strItems & "[      ] " & objItem.Name & vbCrLf
    end if
Next

' 全て表示
WScript.Echo strItems

