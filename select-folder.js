var Shell = new ActiveXObject( "Shell.Application" );
var WSHShell = new ActiveXObject("WScript.Shell");
// ダイアログ表示
var objFolder = Shell.BrowseForFolder( 0, "フォルダ選択", 11, 0 );
// キャンセル
if ( objFolder == null ) {
    WScript.Quit();
}
if ( !objFolder.Self.IsFileSystem ) {
    WSHShell.Popup( "ファイルシステムではありません" );
    WScript.Quit();
}

WScript.Echo( objFolder.Self.Path );

var strItems = ""

// フォルダ内のコレクションを取得
var objFolderItems = objFolder.Items();

// コレクションを列挙
var nFiles = objFolderItems.Count;
for( i = 0; i < nFiles; i++ ) {
    var objItem = objFolderItems.Item(i)
    if ( objItem.isFolder ) {
        strItems += "[Folder] " + objItem.Name + "\n";
    }
    else {
        strItems += "[      ] " + objItem.Name + "\n";
    }
}

// 全て表示
WScript.Echo( strItems );

