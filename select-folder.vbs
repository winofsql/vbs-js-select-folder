Set Shell = CreateObject( "Shell.Application" )
Set WSHShell = CreateObject("WScript.Shell")
' �_�C�A���O�\��
Set objFolder = Shell.BrowseForFolder( 0, "�t�H���_�I��", 11, 0 )
' �L�����Z��
if objFolder is nothing then
    WScript.Quit
end if
if not objFolder.Self.IsFileSystem then
    WSHShell.Popup "�t�@�C���V�X�e���ł͂���܂���"
    WScript.Quit
end if

WScript.Echo objFolder.Self.Path

strItems = ""

' �t�H���_���̃R���N�V�������擾
Set objFolderItems = objFolder.Items()

' �R���N�V�������
nFiles = objFolderItems.Count
For I = 0 to nFiles - 1
    Set objItem = objFolderItems.Item(I)
    if objItem.isFolder then
        strItems = strItems & "[Folder] " & objItem.Name & vbCrLf
    else
        strItems = strItems & "[      ] " & objItem.Name & vbCrLf
    end if
Next

' �S�ĕ\��
WScript.Echo strItems

