import { getAndExtract } from './utils'
import fs from 'fs'
import { csv2json, excel2json, json2excel } from './commonUtils'

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
        json2excel(fresults, './13tokyo.xlsx')
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
  const resultPromise = excel2json('input.xlsx')
  resultPromise.then((results) => {
    console.table(results)
    json2excel(results, 'excelResults.xlsx').then((path) => console.log(path))
  })
}

/**
 * csvを読み込むサンプル。データは全部文字列で取得できる。
 */
async function sample3() {
  await getAndExtract('http://jusyo.jp/downloads/new/csv/csv_13tokyo.zip')
  const csvDatas: Array<any> = (await csv2json('./13tokyo.csv')).filter((address) =>
    address['郵便番号'].startsWith('100-000'),
  )
  console.table(csvDatas)
}

/**
 * Excelファイルを読み込むサンプル。データは入っているデータに応じて、型変換されて取り込まれる
 * 今時点、日付などは適切にフォーマット変換が必要みたいだ
 */
async function sample4() {
  const excelDatas: Array<any> = await excel2json('input.xlsx')
  console.table(excelDatas)
}

if (!module.parent) {
  // sample1()
  // sample2()
  (async () => {
    await sample3()
    await sample4()
  })()
}
