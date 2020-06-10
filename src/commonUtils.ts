
import fs from 'fs'
import iconv from 'iconv-lite'
import csv from 'csvtojson'

const XlsxPopulate = require('xlsx-populate')


/**
 * Excelファイルを読み込み、各行をデータとして配列で返すメソッド。
 * @param path Excelファイルパス
 * @param sheet シート名
 * @param format_func フォーマット関数。instanceは各行データが入ってくるので、任意に整形して返せばよい
 */
export const xlsx2json = async (
  inputFullPath: string,
  sheetName = 'Sheet1',
  format_func?: (instance: any) => any,
): Promise<Array<any>> => {
  const workbook: any = await XlsxPopulate.fromFileAsync(inputFullPath)
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
  return new Promise((resolve, reject) => {
    const datas: any[] = []

    fs.createReadStream(filePath)
      .pipe(iconv.decodeStream('Shift_JIS'))
      .pipe(iconv.encodeStream('utf-8'))
      .pipe(csv().on('data', (data) => datas.push(JSON.parse(data))))
      .on('end', () => resolve(datas))
  })
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