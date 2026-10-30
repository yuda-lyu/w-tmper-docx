import fs from 'fs'
import path from 'path'
import assert from 'assert'


//相依守門: 圖片模組已納入src/imageModule, 套件不再依賴npm套件docxtemplater-image-module-free, 相依樹亦不含已棄用且多數漏洞無修補版之xmldom(非@xmldom/xmldom).
//授權守門: 納入之第三方程式須保留原作者與MIT授權聲明.


let pkg = JSON.parse(fs.readFileSync('./package.json', 'utf8'))
let lock = JSON.parse(fs.readFileSync('./package-lock.json', 'utf8'))

function listSrc(fd) {
    return fs.readdirSync(fd, { recursive: true, withFileTypes: true })
        .filter((d) => d.isFile())
        .map((d) => path.join(d.parentPath || d.path, d.name))
}

//以模組指定符(from 'xmldom' / require('xmldom') / import('xmldom') / import 'xmldom')判斷引用, 不含註解內提到之字樣與@xmldom/xmldom
function importsModule(text, name) {
    let re = new RegExp(`(?:from\\s*|require\\s*\\(\\s*|import\\s*\\(\\s*|import\\s+)['"]${name}['"]`)
    return re.test(text)
}


describe('unit-dep', function() {

    it('package.json不宣告docxtemplater-image-module-free與xmldom', function() {
        for (let field of ['dependencies', 'devDependencies', 'peerDependencies', 'optionalDependencies']) {
            let o = pkg[field] || {}
            assert.ok(!('docxtemplater-image-module-free' in o), `${field}: docxtemplater-image-module-free`)
            assert.ok(!('xmldom' in o), `${field}: xmldom`)
        }
    })

    it('package-lock.json之相依樹不含docxtemplater-image-module-free與xmldom(任何套件之相依亦未列)', function() {
        for (let [k, v] of Object.entries(lock.packages)) {
            assert.ok(!/(^|\/)node_modules\/docxtemplater-image-module-free$/.test(k), k)
            assert.ok(!/(^|\/)node_modules\/xmldom$/.test(k), k)
            for (let field of ['dependencies', 'devDependencies', 'peerDependencies', 'optionalDependencies']) {
                let o = v[field] || {}
                assert.ok(!('docxtemplater-image-module-free' in o), `${k} ${field}: docxtemplater-image-module-free`)
                assert.ok(!('xmldom' in o), `${k} ${field}: xmldom`)
            }
        }
    })

    it('已安裝之node_modules不含docxtemplater-image-module-free與xmldom(須執行npm i使其與鎖檔一致)', function() {
        assert.ok(!fs.existsSync('./node_modules/docxtemplater-image-module-free'))
        assert.ok(!fs.existsSync('./node_modules/xmldom'))
    })

    it('src/不引用xmldom與docxtemplater-image-module-free', function() {
        let files = listSrc('./src')
        assert.ok(files.length >= 5, 'src files')
        for (let f of files) {
            let text = fs.readFileSync(f, 'utf8')
            assert.ok(!importsModule(text, 'xmldom'), `${f} imports xmldom`)
            assert.ok(!importsModule(text, 'docxtemplater-image-module-free'), `${f} imports docxtemplater-image-module-free`)
        }
    })

    it('圖片模組之pptx路徑已刪除(本套件只套印docx)', function() {
        for (let f of listSrc('./src/imageModule')) {
            let text = fs.readFileSync(f, 'utf8')
            assert.ok(!/getInnerPptx|getRenderedPartPptx|getPptxImageXml/.test(text), f)
        }
    })

    it('納入之第三方程式保留原作者與MIT授權聲明, 且LICENSE載明全文', function() {
        let head = fs.readFileSync('./src/imageModule/index.mjs', 'utf8').slice(0, 3000)
        assert.ok(head.includes('Edgar HIPP'))
        assert.ok(head.includes('Dk Saha'))
        assert.ok(head.includes('MIT'))
        let lic = fs.readFileSync('./LICENSE', 'utf8')
        assert.ok(lic.includes('Copyright (c) 2013 Edgar HIPP'))
        assert.ok(lic.includes('docxtemplater-image-module-free'))
        assert.ok((lic.match(/Permission is hereby granted, free of charge/g) || []).length >= 2, 'MIT全文須各載明一份(本套件與第三方)')
        for (let f of listSrc('./src/imageModule')) {
            let t = fs.readFileSync(f, 'utf8').slice(0, 600)
            assert.ok(t.includes('docxtemplater-image-module-free'), `${f} 檔頭須註明來源`)
        }
    })

})
