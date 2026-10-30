# 狀態(現況)

本檔記載當前已開發功能之最新狀態:各功能做到哪裡、對應之規則帳本條目(`sgud/rule/rulebook.md`)與測試檔、已知缺口與待辦、最近一次驗證或發布之版本.只記現況,過程寫入 `sgud/exp/experience.md`.

## 功能

| 功能 | 做到哪裡 | 規則帳本 | 測試檔 |
|---|---|---|---|
| 文字取代 `[[key]]`(含中文、換行 `\n`、迴圈 `[[#k]]…[[/k]]`) | 已完成;委由 docxtemplater(`paragraphLoop`、`linebreaks`、定界符 `[[ ]]`) | — | `test/replace.test.mjs`(僅比輸出檔大小)、`test/unit-image.test.mjs`(文字欄位取代案) |
| 圖片取代 `[[%key]]`、`[[%%key]]` | 已完成;圖片模組內建於 `src/imageModule/`(2026-10-30 起,原為 npm 套件 `docxtemplater-image-module-free`),只支援 docx | IMG-01、IMG-02、IMG-03 | `test/unit-image.test.mjs`(15 案)、`test/replace.test.mjs` |
| 圖寬高上限 `widthMaxDef`/`heightMaxDef`/`kpWidthMax`/`kpHeightMax` | 已完成(`src/WTmperDocx.mjs` 之 `getSize`) | IMG-02 | `test/unit-image.test.mjs` |
| 相依與授權守門 | 已完成 | DEP-01、LIC-01 | `test/unit-dep.test.mjs`(6 案) |

## 最近一次驗證

- 2026-10-30 run01(修正輪,對應 Gitea issue #28):`npm ls xmldom` 為 `(empty)`;`npm test` 22 passing(`replace` 1 + `unit-image` 15 + `unit-dep` 6);變異 11 個皆被測試攔下;修改前後套印同一模板(4 情境)zip 內全部項目 sha256 相同.**尚未發布**(發布由主控/使用者處理,`package.json` version 維持 1.0.28).

## 已知缺口與待辦

- `src/WTmperDocx.mjs` 之 `silenceConsoleWarn` 為既有未使用函式(死碼,依規定只提出不刪);`eslint --ext .mjs src` 因此回報 1 個 `no-unused-vars`.
- `test/replace.test.mjs` 只比對輸出檔大小(弱斷言,屬現狀指紋);語意斷言由 `unit-image.test.mjs` 補強,未改動該檔.
- `test/replace.test.mjs` 會覆寫 git 追蹤中之 `test/report.docx`(zip 時間戳不同),跑完測試後該檔會呈修改狀態,屬既有行為.
- `package-lock.json` 根項目 `version`(1.0.27)落後 `package.json`(1.0.28),屬既有狀況,未動.
- `dist/` 與 `docs/` 不手動重建(發布流程自動產製);run01 已另以暫存目錄驗證新程式可被 rollup 打包為 UMD,且打包結果套印與修改前一致.
