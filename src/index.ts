import { getAndExtract } from './utils'
import fs from 'fs'
import { csv2json, internalSave2Excel, xlsx2json } from './commonUtils'


/**
 * ネットからZIP化されたファイルをダウンロードして解凍、
 * 取り出したCSVを読んでJSONにし、最後にExcelに書き出すサンプル
 */
function sample1() {
  getAndExtract('http://jusyo.jp/downloads/new/csv/csv_13tokyo.zip').then(() => {
    csv2json('./13tokyo.csv')
      .then((results: any[]) => {
        // 郵便番号が「100-000x」のものに絞ってみた
        const fresults = results.filter((address) => address['郵便番号'].startsWith('100-000'))
        console.table(fresults)
        for (const address of fresults) {
          console.log(address)
        }
        internalSave2Excel(fresults, './13tokyo.xlsx', '', 'Sheet1')
      })
      .finally(() => {
        if (fs.existsSync('./13tokyo.csv')) {
          fs.unlinkSync('./13tokyo.csv')
        } else {
          console.log('あれ？ない')
        }
      })
  })
  // https://tsurutoro.com/pseudo_data/
}


function sample2() {
  const resultPromise = xlsx2json('input.xlsx')
  resultPromise.then((results) => {
    console.table(results)
    internalSave2Excel(results, 'excelResults.xlsx', '', 'Sheet1').then((path) => console.log(path))
  })
}

if (!module.parent) {
  // sample1()
  sample2()
}
