# vbs-js-select-folder
VBScript でフォルダを選択して、フォルダ内のファイル(含フォルダ)を列挙する
```javascript
    "code-runner.showRunIconInEditorTitleMenu": false,
    "code-runner.executorMap": {
        "javascript": "cscript //Nologo"
    }
```
![image](https://user-images.githubusercontent.com/1501327/131621428-5438ea0f-f814-42a6-b5bb-d1f51d75a509.png)

[BrowseForFolder メソッド](https://docs.microsoft.com/ja-jp/windows/win32/shell/ishelldispatch-browseforfolder)

## 11
BIF_RETURNONLYFSDIRS (0x00000001)
--
BIF_DONTGOBELOWDOMAIN (0x00000002)
--
BIF_RETURNFSANCESTORS (0x00000008)
--
