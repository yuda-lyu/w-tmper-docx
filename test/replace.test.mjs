import fs from 'fs'
import assert from 'assert'
import wtd from '../src/WTmperDocx.mjs'


describe('replace', function() {

    it('test replace', async () => {

        let kpData = {
            text: 'abc測試中文',
            image: './test/image.png',
        }
        // console.log('kpData', kpData)

        let fpTmp = './test/tmp.docx'
        let fpOut = `./test/report.docx`
        await wtd(kpData, fpTmp, fpOut)

        let sizeText = fs.statSync(fpOut).size
        let sizeAns = fs.statSync(`./test/report_ans.docx`).size

        let r = sizeText
        let rr = sizeAns
        assert.strict.deepEqual(r, rr)
    })

})
