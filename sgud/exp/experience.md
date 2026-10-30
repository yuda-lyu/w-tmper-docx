# 累積經驗

本檔為執行工作時特有之累積經驗,留給後續之嘗試、執行與判斷,由agent自行編列、擴充與維護.過長時抽出舊者為`experience_<年月日>.md`,並於本檔註明何時抽出.

## 2026-10-30 納入圖片模組、移除 xmldom(run01)

- **`npm uninstall` 會把鎖檔弄髒**(npm 12.1.0,registry 設為私有 Verdaccio):它刪光全部 `resolved` 欄位(1117 處)並把鎖檔根 `version` 改成 package.json 的值,diff 達千行.處置:不用 npm 產出之鎖檔,改以 HEAD 之鎖檔為底只移除目標套件(用腳本做,並以 `assert.deepStrictEqual` 驗證與 npm 結果「忽略 resolved 與根 version」後完全一致),diff 只剩 20 行.同理 `package.json` 會被補結尾換行,HEAD 原無,手動還原.
- **babel 編譯產物納入 src 時**:`module.exports = function(){…}()` 轉 ESM 須寫 `export default (function(){…})()`,漏括號會被當成函式宣告而語法錯誤.保留上游樣式(tab、分號、`_createClass`)與檔案切分,才能與上游 `git diff --no-index` 逐檔比對出「只差預期修改」.
- **docxtemplater 的 `supportedFileTypes` 會靜默丟棄模組**(`dropUnsupportedFileTypesModules`),不能拿來擋 pptx(會悄悄不處理 `[[%key]]`);改在模組 `optionsTransformer` 拋錯.docxtemplater 於 `compile()` 先決定 `fileType` 再呼叫 `optionsTransformer`,故該處可用 `docxtemplater.fileType`.
- **`DocUtils.convertPixelsToEmus` 是上游掛在 docxtemplater 共用 `DocUtils` 上的副作用**(`docUtils.mjs`),保留之;`import Docxtemplater from 'docxtemplater'` 後取 `Docxtemplater.DocUtils`(CJS 無具名匯出).
- **image-size 2.0.2 讀 JPEG 的最小可解析結構**:`FFD8` + 一個帶長度的 `FFE0`(APP0)段 + `FFC0`(SOF0,含高、寬)+ `FFD9`;只有 `FFD8` 緊接 SOF 會回 `Invalid JPG, no size found`(它以第一段長度定位下一標記).
- **圖片模組一律把圖存成 `image_generated_N.png`**,連 JPEG 內容也是;`N` 為模組實例內遞增、跨 part 共用(頁首先於 document.xml 被處理時,頁首圖為 `_1`).新圖 rId＝該 rels 檔現有最大 `rIdN`+1,故範本含 `rId100` 時圖為 `rId101` 起.
- **頁首無獨立 rels 時**,模組以 `document.xml.rels` 複製後清空關聯另建 `word/_rels/header1.xml.rels`;此分支(`createEmptyRelsDoc`)須由含頁首圖片之範本才會走到,`test/tmp.docx` 不涵蓋,故 `test/tools/imageFixture.mjs` 自建含頁首之模板.
- **「輸出不變」的證明法**:同一模板、同一輸入,舊實作(複製舊 `WTmperDocx.mjs` 與舊模組到 `./tmp/…/old/node_modules/`,使其不受 uninstall 影響)與新實作各套印,逐一比對 zip 內全部項目 sha256;整檔 byte 不同僅因 zip 項目時間戳.固化進測試的是「語意斷言 + 修改前輸出之摘要 golden」,golden 內以來源圖鍵標籤代替壓縮後位元組,避免綁定 zlib 版本.
- **`test/replace.test.mjs` 會覆寫被追蹤之 `test/report.docx`**,跑完測試要還原該檔內容(不用 git 還原指令時,以 `git show HEAD:test/report.docx` 取出位元組寫回).
