# vbs-js-select-folder( SHIFT_JIS )
WSH : VBScript でフォルダを選択して、フォルダ内のファイル(含フォルダ)を列挙する

## add settings.json ( Code Runner )
```javascript
    "code-runner.showRunIconInEditorTitleMenu": false,
    "code-runner.executorMap": {
        "javascript": "cscript //Nologo"
    }
```
## 実行画面
![image](https://user-images.githubusercontent.com/1501327/131621428-5438ea0f-f814-42a6-b5bb-d1f51d75a509.png)

[BrowseForFolder メソッド](https://docs.microsoft.com/ja-jp/windows/win32/shell/ishelldispatch-browseforfolder)

## 引数の 11 の 意味
BIF_RETURNONLYFSDIRS (0x00000001)
--
BIF_DONTGOBELOWDOMAIN (0x00000002)
--
BIF_RETURNFSANCESTORS (0x00000008)
--
## 実行結果
```
C:\java
[      ] ADDITIONAL_LICENSE_INFO
[      ] ASSEMBLY_EXCEPTION     
[Folder] bin
[Folder] demo
[Folder] include
[Folder] jre
[Folder] lib
[      ] LICENSE
[      ] release
[Folder] sample
[Folder] src.zip
[      ] THIRD_PARTY_README     
```
