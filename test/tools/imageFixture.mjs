import fs from 'fs'
import crypto from 'crypto'
import zlib from 'zlib'
import PizZip from 'pizzip'


//測試輔助: 產製含多種圖片欄位之docx模板, 並將輸出docx摘要為可比對之結構(圖數, 關聯, 尺寸EMU), 供unit-image.test.mjs使用


let crcTable = (() => {
    let t = []
    for (let n = 0; n < 256; n++) {
        let c = n
        for (let k = 0; k < 8; k++) {
            c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1)
        }
        t[n] = c >>> 0
    }
    return t
})()

function crc32(buf) {
    let c = 0xffffffff
    for (let i = 0; i < buf.length; i++) {
        c = crcTable[(c ^ buf[i]) & 0xff] ^ (c >>> 8)
    }
    return (c ^ 0xffffffff) >>> 0
}

function pngChunk(type, data) {
    let len = Buffer.alloc(4)
    len.writeUInt32BE(data.length)
    let body = Buffer.concat([Buffer.from(type, 'ascii'), data])
    let crc = Buffer.alloc(4)
    crc.writeUInt32BE(crc32(body))
    return Buffer.concat([len, body, crc])
}


/**
 * 產製純色PNG二進位資料
 *
 * @param {Integer} width 輸入寬度整數(px)
 * @param {Integer} height 輸入高度整數(px)
 * @param {Array} [rgb=[200,30,30]] 輸入RGB陣列
 * @returns {Buffer} 回傳PNG二進位資料
 */
export function makePng(width, height, rgb = [200, 30, 30]) {
    let ihdr = Buffer.alloc(13)
    ihdr.writeUInt32BE(width, 0)
    ihdr.writeUInt32BE(height, 4)
    ihdr[8] = 8 //bit depth
    ihdr[9] = 2 //color type RGB
    let rowLen = 1 + width * 3
    let raw = Buffer.alloc(rowLen * height)
    for (let y = 0; y < height; y++) {
        let o = y * rowLen
        raw[o] = 0 //filter none
        for (let x = 0; x < width; x++) {
            raw[o + 1 + x * 3] = rgb[0]
            raw[o + 2 + x * 3] = rgb[1]
            raw[o + 3 + x * 3] = rgb[2]
        }
    }
    return Buffer.concat([
        Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
        pngChunk('IHDR', ihdr),
        pngChunk('IDAT', zlib.deflateSync(raw)),
        pngChunk('IEND', Buffer.alloc(0)),
    ])
}


/**
 * 產製僅含尺寸資訊之JPEG檔頭(供image-size讀取寬高, 不保證可顯示)
 *
 * @param {Integer} width 輸入寬度整數(px)
 * @param {Integer} height 輸入高度整數(px)
 * @returns {Buffer} 回傳JPEG二進位資料
 */
export function makeJpeg(width, height) {
    let sof = Buffer.alloc(19)
    sof[0] = 0xff
    sof[1] = 0xc0 //SOF0
    sof.writeUInt16BE(17, 2) //length
    sof[4] = 8 //precision
    sof.writeUInt16BE(height, 5)
    sof.writeUInt16BE(width, 7)
    sof[9] = 3 //components
    for (let i = 0; i < 3; i++) {
        sof[10 + i * 3] = i + 1
        sof[11 + i * 3] = 0x11
        sof[12 + i * 3] = 0
    }
    //APP0(JFIF)段, image-size以首段長度定位下一標記
    let app0 = Buffer.concat([
        Buffer.from([0xff, 0xe0, 0x00, 0x10]),
        Buffer.from('JFIF\0', 'ascii'),
        Buffer.from([0x01, 0x01, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00]),
    ])
    return Buffer.concat([Buffer.from([0xff, 0xd8]), app0, sof, Buffer.from([0xff, 0xd9])])
}


export function sha256(buf) {
    return crypto.createHash('sha256').update(buf).digest('hex')
}


let nsW = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
let nsR = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
let nsWp = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
let typeHeaderCt = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
let typeHeaderRel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header'

function p(text) {
    return `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`
}


/**
 * 以既有docx模板為底, 產製含多種圖片欄位之docx模板
 *
 * 欄位: [[text]]文字, [[%imgA]]行內圖, [[%%imgB]]置中圖, [[%imgBig]]大圖(受最大寬高限制),
 * [[%missing]]空值, 表格內[[%imgCell]], 迴圈[[#rows]]內[[%pic]], 頁首[[%hdr]](頁首無獨立rels檔)
 *
 * @param {Buffer} buffBase 輸入既有docx二進位資料(取其styles, theme與名稱空間)
 * @returns {Buffer} 回傳docx二進位資料
 */
export function buildTemplate(buffBase) {
    let zip = new PizZip(buffBase)

    //document.xml, 保留根標籤與名稱空間, 取代body
    let docXml = zip.file('word/document.xml').asText()
    let head = docXml.slice(0, docXml.indexOf('<w:body>') + '<w:body>'.length)
    let tail = '</w:body></w:document>'
    let tbl = '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>' + p('[[%imgCell]]') + '</w:tc></w:tr></w:tbl>'
    let sect = '<w:sectPr><w:headerReference w:type="default" r:id="rId100"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>'
    let body = [
        p('[[text]]'),
        p('[[%imgA]]'),
        p('[[%%imgB]]'),
        p('[[%imgBig]]'),
        p('[[%missing]]'),
        tbl,
        p('[[#rows]]'),
        p('[[name]]'),
        p('[[%pic]]'),
        p('[[/rows]]'),
        '<w:p/>',
        sect,
    ].join('')
    zip.file('word/document.xml', head + body + tail)

    //頁首, 無word/_rels/header1.xml.rels
    let hdr = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="${nsW}" xmlns:r="${nsR}" xmlns:wp="${nsWp}">${p('[[text]]')}${p('[[%hdr]]')}</w:hdr>`
    zip.file('word/header1.xml', hdr)

    //頁首之內容類型與關聯, 關聯Id故意用rId100, 驗證新增圖關聯之Id為現有最大值+1
    let ct = zip.file('[Content_Types].xml').asText()
    ct = ct.replace('</Types>', `<Override PartName="/word/header1.xml" ContentType="${typeHeaderCt}"/></Types>`)
    zip.file('[Content_Types].xml', ct)
    let rels = zip.file('word/_rels/document.xml.rels').asText()
    rels = rels.replace('</Relationships>', `<Relationship Id="rId100" Type="${typeHeaderRel}" Target="header1.xml"/></Relationships>`)
    zip.file('word/_rels/document.xml.rels', rels)

    return zip.generate({ type: 'nodebuffer' })
}


/**
 * 寫出各測試圖檔, 並回傳各圖鍵對應路徑與sha256對圖鍵標籤
 *
 * 圖: imgA 120x80, imgB 400x200, imgBig 1200x900, imgCell 300x150(JPEG), r1 64x64, r2 90x30, hdr 50x50
 *
 * @param {String} fd 輸入寫出資料夾路徑
 * @returns {Object} 回傳{fp,labelBySha}, fp為圖鍵對路徑, labelBySha為sha256對圖鍵
 */
export function writeImages(fd) {
    let specs = {
        imgA: ['a.png', makePng(120, 80, [200, 30, 30])],
        imgB: ['b.png', makePng(400, 200, [30, 200, 30])],
        imgBig: ['big.png', makePng(1200, 900, [30, 30, 200])],
        imgCell: ['cell.jpg', makeJpeg(300, 150)],
        r1: ['r1.png', makePng(64, 64, [10, 10, 10])],
        r2: ['r2.png', makePng(90, 30, [250, 250, 10])],
        hdr: ['hdr.png', makePng(50, 50, [120, 120, 120])],
    }
    let fp = {}
    let labelBySha = {}
    for (let [k, [fn, buf]] of Object.entries(specs)) {
        fp[k] = `${fd}/${fn}`
        fs.writeFileSync(fp[k], buf)
        labelBySha[sha256(buf)] = k
    }
    return { fp, labelBySha }
}


/**
 * 依buildTemplate模板之欄位組成套印用數據
 *
 * @param {Object} fp 輸入writeImages回傳之fp
 * @returns {Object} 回傳kpData
 */
export function buildData(fp) {
    return {
        text: 'abc測試中文',
        imgA: fp.imgA,
        imgB: fp.imgB,
        imgBig: fp.imgBig,
        missing: '',
        imgCell: fp.imgCell,
        rows: [{ name: 'row1', pic: fp.r1 }, { name: 'row2', pic: fp.r2 }],
        hdr: fp.hdr,
    }
}


/**
 * 產製最小pptx(一張投影片含[[%imgA]]欄位), 供驗證非docx模板被拒絕
 *
 * @returns {Buffer} 回傳pptx二進位資料
 */
export function buildMiniPptx() {
    let z = new PizZip()
    z.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>')
    z.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>')
    z.file('ppt/presentation.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>')
    z.file('ppt/_rels/presentation.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>')
    z.file('ppt/slides/slide1.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:sp><p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="3000" cy="2000"/></a:xfrm></p:spPr><p:txBody><a:p><a:r><a:t>[[%imgA]]</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>')
    return z.generate({ type: 'nodebuffer' })
}


function parseAttrs(s) {
    let r = {}
    for (let m of s.matchAll(/([\w:]+)="([^"]*)"/g)) {
        r[m[1]] = m[2]
    }
    return r
}


/**
 * 將輸出docx摘要為可比對之結構
 *
 * 內容: drawings各part內各圖之rId, 行內尺寸(wp:extent), 圖形尺寸(a:ext), 是否置中; rels各關聯檔內之關聯; media各媒體檔及其來源標籤; contentTypeDefaults
 *
 * @param {Buffer} buff 輸入docx二進位資料
 * @param {Object} [labelBySha={}] 輸入sha256對來源圖標籤之物件, 用於將media檔對應回來源圖
 * @returns {Object} 回傳摘要物件
 */
export function summarizeDocx(buff, labelBySha = {}) {
    let zip = new PizZip(buff)
    let names = Object.keys(zip.files).filter((n) => !zip.files[n].dir).sort()

    //drawings
    let drawings = {}
    for (let n of names.filter((x) => /^word\/(document|header\d*|footer\d*)\.xml$/.test(x))) {
        let xml = zip.file(n).asText()
        let list = []
        for (let para of (xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || [])) {
            let centered = para.includes('<w:jc w:val="center"/>')
            for (let d of (para.match(/<w:drawing>[\s\S]*?<\/w:drawing>/g) || [])) {
                let embed = d.match(/r:embed="rId(\d+)"/)
                let ext = d.match(/<wp:extent cx="(\d+)" cy="(\d+)"\/>/)
                let aext = d.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/)
                list.push({
                    rId: embed ? Number(embed[1]) : null,
                    cx: ext ? Number(ext[1]) : null,
                    cy: ext ? Number(ext[2]) : null,
                    aCx: aext ? Number(aext[1]) : null,
                    aCy: aext ? Number(aext[2]) : null,
                    centered,
                })
            }
        }
        if (list.length > 0) {
            drawings[n] = list
        }
    }

    //rels
    let rels = {}
    for (let n of names.filter((x) => /^word\/_rels\/.*\.rels$/.test(x))) {
        let xml = zip.file(n).asText()
        rels[n] = (xml.match(/<Relationship [^>]*\/>/g) || []).map((t) => {
            let a = parseAttrs(t)
            return { Id: a.Id, Type: String(a.Type).split('/').pop(), Target: a.Target }
        })
    }

    //media
    let media = names.filter((n) => n.startsWith('word/media/')).map((n) => {
        let sha = sha256(zip.file(n).asNodeBuffer())
        return { name: n, label: labelBySha[sha] || `unknown:${sha}` }
    })

    //contentTypeDefaults
    let ctXml = zip.file('[Content_Types].xml').asText()
    let contentTypeDefaults = (ctXml.match(/<Default [^>]*\/>/g) || []).map((t) => {
        let a = parseAttrs(t)
        return { Extension: a.Extension, ContentType: a.ContentType }
    })

    return { drawings, rels, media, contentTypeDefaults }
}
