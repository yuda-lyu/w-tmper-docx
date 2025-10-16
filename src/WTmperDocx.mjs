import fs from 'fs'
import get from 'lodash-es/get.js'
import isnum from 'wsemi/src/isnum.mjs'
import cdbl from 'wsemi/src/cdbl.mjs'
import PizZip from 'pizzip'
import Docxtemplater from 'docxtemplater'
import ImageModule from 'docxtemplater-image-module-free'
import sizeOf from 'image-size'


function silenceConsoleWarn(fn) {
    let originalWarn = console.warn
    console.warn = () => {}
    try {
        return fn()
    }
    finally {
        console.warn = originalWarn
    }
}

/**
 * Docx模板取代器
 *
 * @param {Object} kpData 輸入轉換物件，模板內取代用鍵需用中括號包住，若鍵為keyText，模板內須須取代文字給予[[keyText]]，若鍵為keyImage，模板內須須取代文字須多給予%，也就是給予[[%keyImage]]
 * @param {Object} fpTmp 輸入模板檔案路徑字串
 * @param {Object} fpOut 輸入取代後輸出檔案路徑字串
 * @param {Object} [opt={}] 輸入設定物件，預設{}
 * @param {Integer} [opt.widthMaxDef=1000] 輸入預設圖最大寬度整數，單位px，預設1000
 * @param {Integer} [opt.heightMaxDef=1000] 輸入預設圖最大高度整數，單位px，預設1000
 * @param {Object} [opt.kpWidthMax={}] 輸入指定圖鍵之最大圖寬物件，各值單位px，預設{}
 * @param {Object} [opt.kpHeightMax={}] 輸入指定圖鍵之最大圖高物件，各值單位px，預設{}
 * @returns {Promise} 回傳Promise，resolve代表成功，reject代表執行失敗
 * @example
 *
 * import wtd from './src/WTmperDocx.mjs'
 *
 * let kpData = {
 *     text: 'abc測試中文',
 *     image: './test/image.png',
 * }
 * console.log('kpData', kpData)
 *
 * let fpTmp = './test/tmp.docx'
 * let fpOut = `./test/report.docx`
 * await wtd(kpData, fpTmp, fpOut)
 *
 */
let WTmpertmerx = async(kpData, fpTmp, fpOut, opt = {}) => {

    //widthMaxDef
    let widthMaxDef = get(opt, 'widthMaxDef', null)
    if (!isnum(widthMaxDef)) {
        widthMaxDef = 1000
    }

    //heightMaxDef
    let heightMaxDef = get(opt, 'heightMaxDef', null)
    if (!isnum(heightMaxDef)) {
        heightMaxDef = 1000
    }

    //buffIn, 讀入 Word 檔為二進位
    let buffIn = fs.readFileSync(fpTmp)

    //zip, 使用pizzip解壓縮
    let zip = new PizZip(buffIn)

    //imageModule, 取代圖模組, 記得模板內的key已改用中括號且key前面要添加百分比: [[%key]]
    let optIm = {
        centered: false, //是否置中
        fileType: 'tmerx',
        getImage: function (tagValue, tagName) {
            // console.log('getImage', tagValue, tagName)

            let buff = fs.readFileSync(tagValue)

            return buff

        },
        getSize: function (img, tagValue, tagName) {
            // console.log('getSize', tagValue, tagName)

            //widthMax
            let widthMax = get(opt, `kpWidthMax.${tagName}`, null)
            if (!isnum(widthMax)) {
                widthMax = widthMaxDef
            }
            widthMax = cdbl(widthMax)

            //heightMax
            let heightMax = get(opt, `kpHeightMax.${tagName}`, null)
            if (!isnum(heightMax)) {
                heightMax = heightMaxDef
            }
            heightMax = cdbl(heightMax)

            //width, height
            let dimensions = sizeOf(img)
            let { width, height } = dimensions

            //ratio
            let ratio = 1
            if (width > widthMax || height > heightMax) {
                //圖片寬高按比例縮放
                ratio = Math.min(widthMax / width, heightMax / height)
                width *= ratio
                height *= ratio
            }

            return [width, height]
        }
    }
    let imageModule = new ImageModule(optIm)

    //tmer, 建立docxtemplater實體
    let tmer = new Docxtemplater(zip, {
        modules: [imageModule], //加載imageModule
        // paragraphLoop: true,
        // linebreaks: true,
        delimiters: { start: '[[', end: ']]' }, //預設{{key}}會無法使用, xml內被分拆儲存導致無法解析, 故改用[[key]]
    })

    //setData
    let originalWarn = console.warn
    console.warn = () => {}
    try {
        tmer.setData(kpData) //內部會顯示Deprecated method ".setData", 執行時暫時禁用console.warn
    }
    finally {
        console.warn = originalWarn
    }

    //render, 不能攔錯, 否則tmer轉二進位數據與writeFileSync無法發現render內部錯誤
    tmer.render()
    // console.log('tmer', tmer)

    //buffOut, 轉出新二進位數據
    let buffOut = tmer
        .getZip()
        .generate({ type: 'nodebuffer' })
    // console.log('buffOut', buffOut)

    //writeFileSync
    fs.writeFileSync(fpOut, buffOut)

}


export default WTmpertmerx
