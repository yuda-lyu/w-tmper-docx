# w-tmper-docx
A replacer for template of docx.

![language](https://img.shields.io/badge/language-JavaScript-orange.svg) 
[![npm version](http://img.shields.io/npm/v/w-tmper-docx.svg?style=flat)](https://npmjs.org/package/w-tmper-docx) 
[![license](https://img.shields.io/npm/l/w-tmper-docx.svg?style=flat)](https://npmjs.org/package/w-tmper-docx) 
[![npm download](https://img.shields.io/npm/dt/w-tmper-docx.svg)](https://npmjs.org/package/w-tmper-docx) 
[![npm download](https://img.shields.io/npm/dm/w-tmper-docx.svg)](https://npmjs.org/package/w-tmper-docx) 
[![jsdelivr download](https://img.shields.io/jsdelivr/npm/hm/w-tmper-docx.svg)](https://www.jsdelivr.com/package/npm/w-tmper-docx)

## Documentation
To view documentation or get support, visit [docs](https://yuda-lyu.github.io/w-tmper-docx/global.html).

## Installation
### Using npm(ES6 module):
```alias
npm i w-tmper-docx
```

#### Example:
> **Link:** [[dev source code](https://github.com/yuda-lyu/w-tmper-docx/blob/master/g.mjs)]
```alias
import wtd from './src/WTmperDocx.mjs'

let kpData = {
    text: 'abc測試中文',
    image: './test/image.png',
}
console.log('kpData', kpData)

let fpTmp = './test/tmp.docx'
let fpOut = `./test/report.docx`
await wtd(kpData, fpTmp, fpOut)
```
