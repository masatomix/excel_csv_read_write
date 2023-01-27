import * as fs from 'fs'
import * as path from 'path'
import { getAndExtract } from '../src/utils'
import { csv2json, json2excel, excel2json, excelStream2json } from '../src/commonUtils'
import { assertBasicArray } from './utils'
import { Converters, isCSVData } from '../src/data'


describe('テスト', () => {
  jest.setTimeout(20000);

  const url = 'http://jusyo.jp/downloads/new/csv/csv_13tokyo.zip'
  let csvPath = ''
  const baseDir: string = path.resolve('')
  let fullPath = ''
  const excelPath = path.join('', 'excel出力' + '.xlsx')
  const excelPath2 = path.join('', '13tokyoResult.xlsx')
  const excelPath3 = path.join('', 'resultExcel3.xlsx')


  beforeEach(async () => {
    csvPath = await getAndExtract(url)
    fullPath = path.join(baseDir, csvPath)
    await createExcel(fullPath, excelPath)

    console.log(`baseDir: ${baseDir}`)
    console.log(`csvPath: ${csvPath}`)
    console.log(`fullPath: ${fullPath}`)
    console.log(`excelPath: ${excelPath}`)
    console.log(`excelPath: ${excelPath2}`)
  }, 20000)

  it('excel', async () => {
    const results = (await excel2json(excelPath))
      .filter((address) =>
        isCSVData(address) ?
          address.住所CD ? address.住所CD === 100841000 : false
          : false
      )
    // console.table(results)
    assertBasicArray(results, 22)

    for (const result of results) {
      // console.log(JSON.stringify(result))
      console.log(result)
    }
  })


  it('excel2', async () => {
    const results: unknown[] = await excel2json('13tokyo.csv.xlsx')
    console.table(results)
    const filePath1 = await json2excel(results, excelPath2)
    console.log(filePath1)

    const stream = fs.createReadStream('13tokyo.csv.xlsx')
    const results2 = await excelStream2json(stream)
    console.table(results2)
    const filePath2 = await json2excel(results, excelPath2)
    console.log(filePath2)
  })



  it('excel3', async () => {

    const instances = await csv2json(fullPath)
    // プロパティごとに、変換メソッドをかませたケース
    // nowとIdというプロパティには、変換methodを指定
    const converters: Converters = {
      住所CD: (value: string) => 'PREFIX:' + value,
      郵便番号: (value: string) => '〒:' + value,
    }
    await json2excel(instances, excelPath3, '', 'Sheet1', converters)
  })

  // it('excel入力', async () => {
  //   await createExcel(fullPath,excelPath)

  //   const results = await excel2json('13tokyo.csv.xlsx')
  //   console.table(results)
  //   const filepath = await json2excel(results, path.join('output', '13tokyoResult.xlsx'))
  //   console.log(filepath)
  // })

  afterEach(async () => {
    !fs.existsSync(fullPath) ?? fs.unlinkSync(fullPath)
    !fs.existsSync(excelPath) ?? fs.unlinkSync(excelPath)
    !fs.existsSync(excelPath2) ?? fs.unlinkSync(excelPath2)
    !fs.existsSync(excelPath3) ?? fs.unlinkSync(excelPath3)
  })

  const createExcel = async (csvPath: string, excelPath: string): Promise<void> => {
    const instances = await csv2json(csvPath)
    await json2excel(instances, excelPath)
    expect(fs.existsSync(excelPath)).toBeTruthy()
  }
})

