import wtd from './src/WTmperDocx.mjs'


let kpData = {
    text: 'abc測試中文',
    image: './test/image.png',
}
console.log('kpData', kpData)

let fpTmp = './test/tmp.docx'
let fpOut = `./test/report.docx`
await wtd(kpData, fpTmp, fpOut)


//node g.mjs
