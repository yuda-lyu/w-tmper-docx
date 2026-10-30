import fs from 'fs'
import assert from 'assert'
import PizZip from 'pizzip'
import wtd from '../src/WTmperDocx.mjs'
import ImageModule from '../src/imageModule/index.mjs'
import { writeImages, buildData, buildTemplate, buildMiniPptx, summarizeDocx, sha256 } from './tools/imageFixture.mjs'


//圖片欄位[[%key]]之套印行為. 規格來源: src/WTmperDocx.mjs 之JSDoc(kpData圖鍵須給[[%keyImage]], opt.widthMaxDef/heightMaxDef/kpWidthMax/kpHeightMax單位px),
//以及圖片模組之約定([[%%key]]置中, px換EMU為px*9525, 空值不產生圖).
//golden(test/unit-image.golden.json)為修改前(圖片模組為npm套件docxtemplater-image-module-free 1.1.1)套印同一模板之輸出摘要, 用於證明輸出未變.


let fd = './test/_tmp/image'
let golden = JSON.parse(fs.readFileSync('./test/unit-image.golden.json', 'utf8')).scenarios

let optDefault = {}
let optCaps = { widthMaxDef: 800, heightMaxDef: 600, kpWidthMax: { imgBig: 600 }, kpHeightMax: { imgBig: 300, pic: 40 } }
let optWidthOnly = { kpWidthMax: { imgBig: 500 } }

//預期: [圖鍵, cx(EMU), cy(EMU), 是否置中], 順序為part內出現順序
let expDocDefault = [
    ['imgA', 1143000, 762000, false], //120x80, 未超過1000
    ['imgB', 3810000, 1905000, true], //400x200, [[%%imgB]]置中
    ['imgBig', 9525000, 7143750, false], //1200x900, 超過1000: ratio=min(1000/1200,1000/900), 縮為1000x750
    ['imgCell', 2857500, 1428750, false], //表格內, 300x150
    ['r1', 609600, 609600, false], //迴圈第1筆, 64x64
    ['r2', 857250, 285750, false], //迴圈第2筆, 90x30
]
let expHdr = [['hdr', 476250, 476250, false]] //頁首, 50x50

let expDocCaps = JSON.parse(JSON.stringify(expDocDefault))
expDocCaps[2] = ['imgBig', 3810000, 2857500, false] //kpWidthMax=600,kpHeightMax=300: ratio=min(600/1200,300/900)=1/3, 縮為400x300
expDocCaps[4] = ['r1', 381000, 381000, false] //kpHeightMax.pic=40: ratio=40/64, 縮為40x40; r2為90x30未超過故不變

let expDocWidthOnly = JSON.parse(JSON.stringify(expDocDefault))
expDocWidthOnly[2] = ['imgBig', 4762500, 3571875, false] //kpWidthMax=500: ratio=min(500/1200,1000/900), 縮為500x375

let sc = {} //各情境之{ buff, sum }
let labels = ['imgA', 'imgB', 'imgBig', 'imgCell', 'r1', 'r2']

//將part內各圖解析為{ label, cx, cy, centered, relType, aCx, aCy }, label由rId→rels→media→來源圖
function resolveDrawings(sum, part) {
    let relsName = `word/_rels/${part.split('/').pop()}.rels`
    return (sum.drawings[part] || []).map((d) => {
        let rel = (sum.rels[relsName] || []).find((r) => r.Id === `rId${d.rId}`)
        let media = rel ? sum.media.find((m) => m.name === `word/${rel.Target}`) : null
        return { ...d, relType: rel ? rel.Type : null, label: media ? media.label : null }
    })
}

function toRows(list) {
    return list.map((d) => [d.label, d.cx, d.cy, d.centered])
}


describe('unit-image', function() {

    before(async function() {
        fs.rmSync(fd, { recursive: true, force: true })
        fs.mkdirSync(fd, { recursive: true })

        let { fp, labelBySha } = writeImages(fd)
        let fpTmpl = `${fd}/multi.docx`
        fs.writeFileSync(fpTmpl, buildTemplate(fs.readFileSync('./test/tmp.docx')))
        let kpData = buildData(fp)

        let run = async(name, fpTmp, data, opt, lbs) => {
            let fpOut = `${fd}/${name}.docx`
            await wtd(data, fpTmp, fpOut, opt)
            let buff = fs.readFileSync(fpOut)
            sc[name] = { buff, sum: summarizeDocx(buff, lbs) }
        }
        await run('multi-default', fpTmpl, kpData, optDefault, labelBySha)
        await run('multi-caps', fpTmpl, kpData, optCaps, labelBySha)
        await run('width-only', fpTmpl, kpData, optWidthOnly, labelBySha)

        let lbsBasic = { [sha256(fs.readFileSync('./test/image.png'))]: 'basic' }
        await run('basic', './test/tmp.docx', { text: 'abc測試中文', image: './test/image.png' }, { kpWidthMax: { image: 300 } }, lbsBasic)
    })

    after(function() {
        fs.rmSync(fd, { recursive: true, force: true })
        try {
            fs.rmdirSync('./test/_tmp') //僅在已空時移除, 其他測試仍使用則保留
        }
        catch (err) {}
    })

    it('一個非空[[%key]]產生一張圖, 依出現順序; 空值不產生圖; 頁首亦可插圖', function() {
        for (let name of ['multi-default', 'multi-caps', 'width-only']) {
            let doc = resolveDrawings(sc[name].sum, 'word/document.xml')
            assert.deepStrictEqual(doc.map((d) => d.label), labels, name) //missing為空值, 不在其中
            let hdr = resolveDrawings(sc[name].sum, 'word/header1.xml')
            assert.deepStrictEqual(hdr.map((d) => d.label), ['hdr'], name)
        }
    })

    it('每張圖之media位元組即為來源圖檔(getImage讀檔), 且共7個media', function() {
        let { sum } = sc['multi-default']
        assert.strictEqual(sum.media.length, 7)
        assert.deepStrictEqual(sum.media.map((m) => m.label).sort(), [...labels, 'hdr'].sort()) //label由來源圖sha256得出, 位元組不同則為unknown:*
    })

    it('尺寸: 預設上限1000x1000, 僅超過者按比例縮放, EMU為px*9525', function() {
        let { sum } = sc['multi-default']
        assert.deepStrictEqual(toRows(resolveDrawings(sum, 'word/document.xml')), expDocDefault)
        assert.deepStrictEqual(toRows(resolveDrawings(sum, 'word/header1.xml')), expHdr)
    })

    it('尺寸: opt.widthMaxDef/heightMaxDef與kpWidthMax/kpHeightMax(依圖鍵)', function() {
        let { sum } = sc['multi-caps']
        assert.deepStrictEqual(toRows(resolveDrawings(sum, 'word/document.xml')), expDocCaps)
        assert.deepStrictEqual(toRows(resolveDrawings(sum, 'word/header1.xml')), expHdr)
    })

    it('尺寸: 僅給kpWidthMax時, 高度上限用預設1000', function() {
        let { sum } = sc['width-only']
        assert.deepStrictEqual(toRows(resolveDrawings(sum, 'word/document.xml')), expDocWidthOnly)
    })

    it('尺寸: 範本test/tmp.docx之[[%image]](566x457)受kpWidthMax.image=300限制, 為300x242.226px', function() {
        let { sum } = sc['basic']
        let list = resolveDrawings(sum, 'word/document.xml')
        assert.strictEqual(list.length, 1)
        assert.strictEqual(list[0].label, 'basic')
        assert.strictEqual(list[0].cx, 2857500) //300*9525
        assert.strictEqual(list[0].cy, 2307204) //round(457*300/566*9525)
    })

    it('行內尺寸wp:extent與圖形尺寸a:ext一致', function() {
        for (let name of ['multi-default', 'multi-caps', 'width-only', 'basic']) {
            for (let part of Object.keys(sc[name].sum.drawings)) {
                for (let d of sc[name].sum.drawings[part]) {
                    assert.strictEqual(d.aCx, d.cx, `${name} ${part}`)
                    assert.strictEqual(d.aCy, d.cy, `${name} ${part}`)
                }
            }
        }
    })

    it('[[%%key]]為置中段落, 僅imgB; 其餘為行內圖', function() {
        let doc = resolveDrawings(sc['multi-default'].sum, 'word/document.xml')
        assert.deepStrictEqual(doc.filter((d) => d.centered).map((d) => d.label), ['imgB'])
    })

    it('關聯: rId皆指向image關聯, 目標media存在, 同一rels內Id不重複; 無獨立rels之頁首會建立自己的rels', function() {
        for (let name of ['multi-default', 'multi-caps', 'width-only', 'basic']) {
            let { sum } = sc[name]
            for (let part of Object.keys(sum.drawings)) {
                for (let d of resolveDrawings(sum, part)) {
                    assert.strictEqual(d.relType, 'image', `${name} ${part} rId${d.rId}`)
                    assert.ok(d.label !== null, `${name} ${part} rId${d.rId} media`)
                }
            }
            for (let [relsName, list] of Object.entries(sum.rels)) {
                let ids = list.map((r) => r.Id)
                assert.strictEqual(new Set(ids).size, ids.length, `${name} ${relsName} Id重複`)
            }
        }
        let { sum } = sc['multi-default']
        assert.ok(sum.rels['word/_rels/header1.xml.rels'], '頁首rels')
        //頁首之rels僅含自己的圖關聯(不帶入document.xml.rels之styles, theme等關聯)
        assert.deepStrictEqual(sum.rels['word/_rels/header1.xml.rels'].map((r) => r.Type), ['image'])
        //document.xml.rels內既有最大rId為rId100, 新增之圖關聯Id為其後遞增
        let docImgIds = sum.rels['word/_rels/document.xml.rels'].filter((r) => r.Type === 'image').map((r) => r.Id)
        assert.deepStrictEqual(docImgIds, ['rId101', 'rId102', 'rId103', 'rId104', 'rId105', 'rId106'])
    })

    it('[Content_Types].xml登錄png預設類型', function() {
        for (let name of ['multi-default', 'basic']) {
            let png = sc[name].sum.contentTypeDefaults.filter((x) => x.Extension === 'png')
            assert.deepStrictEqual(png, [{ Extension: 'png', ContentType: 'image/png' }], name)
        }
    })

    it('文字欄位[[text]]被取代, 文件內不殘留[[ ]]定界符', function() {
        let zip = new PizZip(sc['multi-default'].buff)
        for (let part of ['word/document.xml', 'word/header1.xml']) {
            let xml = zip.file(part).asText()
            assert.ok(xml.includes('abc測試中文'), part)
            assert.ok(!xml.includes('[['), part)
            assert.ok(!xml.includes(']]'), part)
        }
    })

    it('輸出與修改前(golden)一致: 圖數, 關聯, 尺寸, 置中, media對應, 內容類型', function() {
        for (let name of Object.keys(golden)) {
            assert.deepStrictEqual(sc[name].sum, golden[name], name)
        }
    })

    it('非docx模板(pptx)被明確拒絕, 不再套用docx路徑', async function() {
        let fpTmp = `${fd}/mini.pptx`
        fs.writeFileSync(fpTmp, buildMiniPptx())
        await assert.rejects(
            wtd({ imgA: `${fd}/a.png` }, fpTmp, `${fd}/mini_out.pptx`),
            /ImageModule only supports docx templates, got fileType: pptx/
        )
        assert.ok(!fs.existsSync(`${fd}/mini_out.pptx`))
    })

    it('圖檔不存在時套印失敗(reject), 不產出檔案', async function() {
        let fpOut = `${fd}/err_out.docx`
        await assert.rejects(
            wtd({ text: 'x', image: `${fd}/not_exist.png` }, './test/tmp.docx', fpOut),
            /ENOENT/
        )
        assert.ok(!fs.existsSync(fpOut))
    })

    it('ImageModule須給getImage與getSize', function() {
        assert.throws(() => new ImageModule({ getSize: () => [1, 1] }), /You should pass getImage/)
        assert.throws(() => new ImageModule({ getImage: () => Buffer.alloc(0) }), /You should pass getSize/)
    })

})
