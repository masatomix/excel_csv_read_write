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

async function sample5() {
  let robots: Array<any> = await csv2json('robotSample.csv')
  robots = robots.map((robot) => Object.assign({}, robot, { now: new Date() })) // 日付列を追加
  console.table(robots)

  // なにも考えずにダンプ
  json2excel(robots, 'output/robots.xlsx')

  // nowというプロパティには、変換をかけたケース
  json2excel(robots, 'output/robots1.xlsx', '', 'Sheet1', {
    now: (value: any) => value,
    // Id: (value: any) => '0' + value,
  })

  // now というプロパティには変換をかけ、さらにその列(M列) は日付フォーマットで出力する
  json2excel(
    robots,
    'output/robots2.xlsx',
    '',
    'Sheet1',
    {
      now: (value: any) => value,
    },
    (instances: any[], workbook: any, sheetName: string) => {
      const rowCount = instances.length
      const sheet = workbook.sheet(sheetName)
      sheet.range(`M2:M${rowCount + 1}`).style('numberFormat', 'yyyy/mm/dd hh:mm') // 書式: 日付+時刻
    },
  ) // プロパティ指定で、変換をかける

  json2excel(robots, 'output/robotsToTemplate.xlsx', 'templateRobots.xlsx', 'Sheet1') // テンプレを指定したケース
}

if (!module.parent) {
  // sample1()
  // sample2()
  ;(async () => {
    // await sample3()
    // await sample4()
    await sample5()
  })()
}
