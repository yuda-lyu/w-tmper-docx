# 規則帳本

本檔為套件設計與開發階段之規則帳本(藍圖),為要求程式碼須長成什麼樣,或由規格產出之測試(項目/說明/流程)有哪些,包含軸、單一擁有者、已知限制、破壞性變更紀錄,以及對應是哪些測試檔案或指令等,由agent自行編列、擴充與維護.當前已做到哪些記於`sgud/state/state.md`.過長時依類型拆為`rulebook_<類型>.md`,本檔只留索引.

> 狀態欄:已套＝程式與測試皆到位;未做/部分須寫明缺什麼.測試一律 `npm test`(= `mocha --parallel "test/*.test.mjs"`),單檔自行以 `npx mocha --no-parallel <檔> --timeout 180000` 執行;本套件無瀏覽器 e2e,故無「改動範圍 → e2e 檔」對照,改到 `src/` 即跑全部 unit(約 1 秒).

## 規則總表

| 編號 | 規則(程式碼須長成什麼樣) | 測試/驗證 | 狀態 |
|---|---|---|---|
| DEP-01 | 套件不依賴 `xmldom`(已棄用,多數漏洞無修補版)與 npm 套件 `docxtemplater-image-module-free`:`package.json`、`package-lock.json`、已安裝之 `node_modules` 與 `src/` 皆不得出現(`@xmldom/xmldom` 為 docxtemplater 自有相依,不在此限) | `test/unit-dep.test.mjs`(前 4 案);`npm ls xmldom` 須為 `(empty)` | 已套 |
| IMG-01 | 圖片模組以 MIT 授權之程式納入 `src/imageModule/`(4 檔:`index.mjs`、`imgManager.mjs`、`templates.mjs`、`docUtils.mjs`),**只支援 docx**,pptx 路徑已刪除;保留之程式與上游 `docxtemplater-image-module-free` 1.1.1 `js/` 之差異僅限:CommonJS→ES module、移除 pptx 與 xmldom、docx 守門(見 IMG-03) | `test/unit-dep.test.mjs`(pptx 路徑已刪除案);上游差異清單見 `盤點`欄 | 已套 |
| IMG-02 | 插圖輸出契約(升版前後不得改變):模板 `[[%key]]` 行內圖、`[[%%key]]` 置中圖(整段);`getImage(tagValue, tagName)` 以 `tagValue` 為檔案路徑讀入;`getSize` 依 `opt.widthMaxDef`/`heightMaxDef`(預設 1000)與 `opt.kpWidthMax[key]`/`kpHeightMax[key]`,寬或高超過上限時按比例縮放(`ratio=min(widthMax/w,heightMax/h)`),否則維持原尺寸;EMU＝`round(px*9525)`;空值(`''`/null/undefined)不產生圖;新圖關聯 Id＝該 rels 內現有最大 `rIdN` + 1;圖檔存為 `word/media/image_generated_N.png`(**副檔名一律 png,即使內容為 JPEG**,上游行為,維持);`[Content_Types].xml` 登錄 `png→image/png`;頁首/頁尾等無獨立 rels 之 part 另建 rels(僅含自己的圖關聯) | `test/unit-image.test.mjs`(語意斷言 14 案 + golden 1 案);golden＝`test/unit-image.golden.json`(修改前實作之輸出摘要) | 已套 |
| IMG-03 | 非 docx 模板(pptx 等)明確拒絕:`optionsTransformer` 於 `docxtemplater.fileType !== 'docx'` 時拋 `ImageModule only supports docx templates, got fileType: <型別>`,`wtd()` 隨之 reject 且不寫出檔案(不得悄悄改走 docx 路徑產出壞檔) | `test/unit-image.test.mjs`(pptx 被明確拒絕案) | 已套 |
| LIC-01 | 納入之第三方程式須保留授權聲明:`src/imageModule/*.mjs` 檔頭註明來源(`index.mjs` 載原作者 Edgar HIPP、Dk Saha 與 MIT 聲明及修改清單),根目錄 `LICENSE` 附第三方段落與 MIT 全文(上游為 MIT/GPLv3 雙授權,本套件採 MIT) | `test/unit-dep.test.mjs`(授權案) | 已套 |
| TST-01 | 圖片測試之夾具與摘要工具集中於 `test/tools/imageFixture.mjs`(產圖、產多欄位模板、摘要 docx);測試臨時檔放 `test/_tmp/image/` 並由測試 `after` 刪除;不使用 `./tmp/` | 目視/Grep:`test/` 內無 `./tmp` | 已套 |

## 盤點指令

- DEP-01:`npm ls xmldom`(期望 `(empty)`);Grep `xmldom|docxtemplater-image-module-free` 於 `src/`、`package.json`、`package-lock.json`(`src/imageModule/` 檔頭註解與 `LICENSE` 之文字說明可出現,模組指定符不可).
- IMG-01 上游差異:`npm pack docxtemplater-image-module-free@1.1.1 --pack-destination <暫存夾>` 解開後,以 `git diff --no-index <上游>/js/<檔>.js src/imageModule/<檔>.mjs` 逐檔比對;2026-10-30 之差異僅:檔頭、`require`→`import`/`module.exports`→`export default`、`getInnerPptx`/`getRenderedPartPptx`/`getPptxImageXml` 刪除、`ImgManager` 去 `fileType` 參數(`ppt`/`../media` 分支刪除)、`optionsTransformer` 加守門.
- IMG-02:輸出等價以「同一模板分別以修改前/後實作套印,逐一比對 zip 內全部項目之 sha256」證明(2026-10-30 run01:4 個情境,全部項目相同;整檔 byte 不同僅因 zip 項目時間戳).

## 已知限制

- 圖片副檔名固定 `.png`(內容為 JPEG 亦然),Word 仍能顯示,但屬上游行為,未修.
- `src/imageModule/` 為上游 babel 編譯產物之 ESM 化,保留其樣式(tab、分號、`_createClass`),以 `/* eslint-disable */` 排除專案 lint,目的是與上游可逐行比對;不為「好看」而改寫.
- `DocUtils.convertPixelsToEmus` 掛在 docxtemplater 共用之 `DocUtils` 物件上(上游副作用,維持).

## 破壞性變更紀錄

- 2026-10-30(待發布,`package.json` version 由主控決定):圖片模組由 npm 套件改為內建,pptx 路徑刪除.本套件原即只套印 docx(`WTmperDocx` 之 JSDoc、README 皆為 docx),此前傳入 pptx 模板雖可能跑出結果,但從未是支援用法;現改為明確 reject.使用端 API(`wtd(kpData, fpTmp, fpOut, opt)`、`[[%key]]` 語法、`getImage`/`getSize` 行為)不變.
