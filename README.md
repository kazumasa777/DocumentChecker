1. Word/Excel主要プロパティをDetailに正確に出力
Word（docx）・Excel（xlsx/xlsm/xls）は、
タイトル、件名、タグ、分類、作成者、前回保存者、改訂番号、バージョン番号
を実際にファイルから取得し、Detail列に
の形式で出力します（空欄は空文字）。
2. C3チェックを完全に除外
C3（余白チェック）はresultsシートに一切出力されないようにします。
3. other_filesシートに対象外ファイルが必ず出力されるよう修正
「pdf, docx, doc, xlsx, xls, vsd, vsdx, ppt, pptx」以外のファイルは必ずother_filesシートに出力されるようロジックを再修正します。

python -m pip download aspose-diagram-python -d aspose_download
3. 取得した wheel を展開し、aspose フォルダ一式を次の場所へ配置します。
   <exeと同じフォルダ>\plugins\aspose_diagram\aspose
4. 配置例:
   documentChecker_gui_new.exe
   plugins\aspose_diagram\aspose\diagram\...
