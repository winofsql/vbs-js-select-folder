var Shell = new ActiveXObject( "Shell.Application" );
var WSHShell = new ActiveXObject("WScript.Shell");
// �_�C�A���O�\��
var objFolder = Shell.BrowseForFolder( 0, "�t�H���_�I��", 11, 0 );
// �L�����Z��
if ( objFolder == null ) {
    WScript.Quit();
}
if ( !objFolder.Self.IsFileSystem ) {
    WSHShell.Popup( "�t�@�C���V�X�e���ł͂���܂���" );
    WScript.Quit();
}

WScript.Echo( objFolder.Self.Path );

var strItems = ""

// �t�H���_���̃R���N�V�������擾
var objFolderItems = objFolder.Items();

// �R���N�V�������
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

// �S�ĕ\��
WScript.Echo( strItems );

