import path from 'path'
import fs, { ReadStream } from 'fs'
import iconv from 'iconv-lite'
import csv from 'csvtojson'
const XlsxPopulate = require('xlsx-populate')

import * as JSZip from 'jszip'

import { getLogger } from './logger'

const logger = getLogger('main')

/**
 * Excelファイルを読み込み、各行をデータとして配列で返すメソッド。
 * @param path Excelファイルパス
 * @param sheet シート名
 * @param sheetName
 * @param format_func フォーマット関数。instanceは各行データが入ってくるので、任意に整形して返せばよい
 */
export const excel2json = async (
  inputFullPath: string,
  sheetName = 'Sheet1',
  format_func?: (instance: any) => any,
): Promise<Array<any>> => {
  const promise = new JSZip.external.Promise((resolve, reject) => {
    fs.readFile(inputFullPath, (err, data) => {
      if (err) return reject(err)
      resolve(data)
    })
  }).then((data) => XlsxPopulate.fromDataAsync(data))

  return excelData2json(await promise, sheetName, format_func)

  // 安定しないので、いったん処理変更
  // const stream: ReadStream = fs.createReadStream(inputFullPath)
  // return excelStream2json(stream, sheetName, format_func)
}

/**
 * Excelファイルを読み込み、各行をデータとして配列で返すメソッド。
 * @param stream
 * @param sheetName
 * @param format_func フォーマット関数。instanceは各行データが入ってくるので、任意に整形して返せばよい
 */
export const excelStream2json = async (
  stream: NodeJS.ReadableStream,
  sheetName = 'Sheet1',
  format_func?: (instance: any) => any,
): Promise<Array<any>> => {
  // cf:https://qiita.com/masakura/items/5683e8e3e655bfda6756
  const promise = new JSZip.external.Promise((resolve, reject) => {
    let buf: any
    stream.on('data', (data) => (buf = data)).on('end', () => resolve(buf))
  }).then((buf) => XlsxPopulate.fromDataAsync(buf))

  return excelData2json(await promise, sheetName, format_func)
}

/**
 * Excelファイルを読み込み、各行をデータとして配列で返すメソッド。
 * @param stream
 * @param sheetName
 * @param format_func フォーマット関数。instanceは各行データが入ってくるので、任意に整形して返せばよい
 */
export const excelData2json = async (
  data: any,
  sheetName = 'Sheet1',
  format_func?: (instance: any) => any,
): Promise<Array<any>> => {
  const workbook: any = data
  const headings: string[] = getHeaders(workbook, sheetName)
  // console.log(headings.length)
  const valuesArray: any[][] = getValuesArray(workbook, sheetName)

  const instances = valuesArray.map((values: any[]) => {
    return values.reduce((box: any, column: any, index: number) => {
      // 列単位で処理してきて、ヘッダの名前で代入する。
      box[headings[index]] = column
      return box
    }, {})
  })

  if (format_func) {
    return instances.map((instance) => format_func(instance))
  }
  return instances
}

/**
 * 指定したパスのcsvファイルをロードして、JSONオブジェクトとしてparseする。
 * 全行読み込んだら完了する Promise を返す。
 * @param filePath
 */
export const csv2json = (filePath: string): Promise<Array<any>> => {
  return csvStream2json(fs.createReadStream(filePath))
  // return new Promise((resolve, reject) => {
  //   const datas: any[] = []

  //   fs.createReadStream(filePath)
  //     .pipe(iconv.decodeStream('Shift_JIS'))
  //     .pipe(iconv.encodeStream('utf-8'))
  //     .pipe(csv().on('data', (data) => datas.push(JSON.parse(data))))
  //     .on('end', () => resolve(datas))
  // })
}

/**
 * 指定したパスのcsvファイルをロードして、JSONオブジェクトとしてparseする。
 * 全行読み込んだら完了する Promise を返す。
 * @param fs
 */
export const csvStream2json = (stream: NodeJS.ReadableStream): Promise<Array<any>> => {
  return new Promise((resolve, reject) => {
    const datas: any[] = []

    stream
      .pipe(iconv.decodeStream('Shift_JIS'))
      .pipe(iconv.encodeStream('utf-8'))
      .pipe(csv().on('data', (data) => datas.push(JSON.parse(data))))
      .on('end', () => resolve(datas))
      // const row = Buffer.isBuffer(data) ? JSON.parse(data.toString()) : JSON.parse(data)

  })
}

/**
 * 引数のJSON配列を、指定したテンプレートを用いて、指定したファイルに出力します。
 * @param instances JSON配列
 * @param outputFullPath 出力Excelのパス
 * @param templateFullPath 元にするテンプレートExcelのパス
 * @param sheetName テンプレートExcelのシート名(シート名で出力する)
 * @param applyStyles 出力時のExcelを書式フォーマットしたい場合に使用する。
 */
export const json2excel = async (
  instances: any[],
  outputFullPath: string,
  templateFullPath = '',
  sheetName = 'Sheet1',
  converters?: any,
  applyStyles?: (instances: any[], workbook: any, sheetName: string) => void,
): Promise<string> => {
  logger.debug(`template path: ${templateFullPath}`)
  // console.log(instances[0])
  // console.table(instances)

  let headings: string[] = [] // ヘッダ名の配列
  let workbook: any
  const fileIsNew: boolean = templateFullPath === '' // templateが指定されない場合新規(fileIsNew = true)、そうでない場合テンプレファイルに出力

  if (!fileIsNew) {
    // 指定された場合は、一行目の文字列群を使ってプロパティを作成する
    workbook = await XlsxPopulate.fromFileAsync(templateFullPath)
    headings = getHeaders(workbook, sheetName)
  } else {
    // templateが指定されない場合は、空ファイルをつくり、オブジェクトのプロパティでダンプする。
    workbook = await XlsxPopulate.fromBlankAsync()
    if (instances.length > 0) {
      headings = Object.keys(instances[0])
    }
  }

  if (instances.length > 0) {
    const csvArrays: any[][] = createCsvArrays(headings, instances, converters)
    // console.table(csvArrays)
    const rowCount = instances.length
    const columnCount = headings.length
    const sheet = workbook.sheet(sheetName)

    if (!fileIsNew && sheet.usedRange()) {
      sheet.usedRange().clear() // Excel上のデータを削除して。
    }
    sheet.cell('A1').value(csvArrays)

    // データがあるところには罫線を引く(細いヤツ)
    const startCell = sheet.cell('A1')
    const endCell = startCell.relativeCell(rowCount, columnCount - 1)

    sheet.range(startCell, endCell).style('border', {
      top: { style: 'hair' },
      left: { style: 'hair' },
      bottom: { style: 'hair' },
      right: { style: 'hair' },
    })

    // よくある整形パタン。
    // sheet.range(`C2:C${rowCount + 1}`).style('numberFormat', '@') // 書式: 文字(コレをやらないと、見かけ上文字だが、F2で抜けると数字になっちゃう)
    // sheet.range(`E2:F${rowCount + 1}`).style('numberFormat', 'yyyy/mm/dd') // 書式: 日付
    // sheet.range(`H2:H${rowCount + 1}`).style('numberFormat', 'yyyy/mm/dd hh:mm') // 書式: 日付+時刻

    if (applyStyles) {
      applyStyles(instances, workbook, sheetName)
    }
  }

  logger.debug(outputFullPath)
  await workbook.toFileAsync(outputFullPath)

  return toFullPath(outputFullPath)
}

const toFullPath = (str: string) => {
  let ret = ''
  if (path.isAbsolute(str)) {
    ret = str
  } else {
    ret = path.join(path.resolve(''), str)
  }
  return ret
}

// 自前実装
function createCsvArrays(headings: string[], instances: any[], converters?: any) {
  const csvArrays: any[][] = instances.map((instance: any): any[] => {
    // console.log(instance)
    const csvArray = headings.reduce((box: any[], header: string): any[] => {
      // console.log(`${instance[header]}: ${instance[header] instanceof Object}`)
      // console.log(converters)
      if (converters && converters[header]) {  // header名に合致するConverterがある場合はそれ優先で適用
        box.push(converters[header](instance[header]))
      } else if (instance[header] instanceof Object) { // Converterがない場合は、文字列に変換
        box.push(JSON.stringify(instance[header])) 
      } else {
        box.push(instance[header]) // あとはそのまま
      }
      return box
    }, [])
    return csvArray
  })
  csvArrays.unshift(headings)
  return csvArrays
}

/**
 * Excelのシリアル値を、Dateへ変換します。
 * @param serialNumber シリアル値
 */
export const dateFromSn = (serialNumber: number): Date => {
  return XlsxPopulate.numberToDate(serialNumber)
}

export const toBoolean = function (boolStr: string | boolean): boolean {
  if (typeof boolStr === 'boolean') {
    return boolStr
  }
  return boolStr.toLowerCase() === 'true'
}

// XlsxPopulate
export const getHeaders = (workbook: any, sheetName: string): string[] => {
  return workbook.sheet(sheetName).usedRange().value().shift()
}

// XlsxPopulate
export const getValuesArray = (workbook: any, sheetName: string): any[][] => {
  const valuesArray: any[][] = workbook.sheet(sheetName).usedRange().value()
  valuesArray.shift() // 先頭除去
  return valuesArray
}
