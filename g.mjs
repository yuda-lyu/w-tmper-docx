import wtd from './src/WTmperDocx.mjs'


let kpData = {
    text: 'abc測試中文',
    image: './test/image.png',
}
console.log('kpData', kpData)

let fpTmp = './test/tmp.docx'
let fpOut = `./test/report.docx`
let opt = {
    kpWidthMax: {
        image: 300,
    },
}
await wtd(kpData, fpTmp, fpOut, opt)


//node g.mjs
