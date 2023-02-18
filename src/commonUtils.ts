import * as fs from 'fs'
import * as path from 'path'
import * as csv from 'csvtojson'
import * as iconv from 'iconv-lite'
import * as JSZip from 'jszip'
import * as XlsxPopulate from 'xlsx-populate'
import { Converters, CSVData } from './data'
import { getLogger } from './logger'
// const XlsxPopulate = require('xlsx-populate')

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
  formatFunc?: (instance: CSVData) => CSVData,
): Promise<unknown[]> => {
  const promise = new JSZip.external.Promise<Buffer>((resolve, reject) => {
    fs.readFile(inputFullPath, (err, data) => (err ? reject(err) : resolve(data)))
  }).then(async (data: Buffer) => await XlsxPopulate.fromDataAsync(data))

  return excelData2json(await promise, sheetName, formatFunc)

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
  formatFunc?: (instance: CSVData) => CSVData,
): Promise<unknown[]> => {
  // cf:https://qiita.com/masakura/items/5683e8e3e655bfda6756
  const promise = new JSZip.external.Promise((resolve, _) => {
    let buf: any
    stream.on('data', (data) => (buf = data)).on('end', () => resolve(buf))
  }).then(async (buf: any) => await XlsxPopulate.fromDataAsync(buf))

  return excelData2json(await promise, sheetName, formatFunc)
}


/**
 * Excelファイルを読み込み、各行をデータとして配列で返すメソッド。
 * @param stream
 * @param sheetName
 * @param format_func フォーマット関数。instanceは各行データが入ってくるので、任意に整形して返せばよい
 */
export const excelData2json = (
  workbook: XlsxPopulate.Workbook,
  sheetName = 'Sheet1',
  formatFunc?: (instance: CSVData) => CSVData,
): CSVData[] => {
  const headings: string[] = getHeaders(workbook, sheetName)
  // console.log(headings.length)
  const valuesArray: unknown[][] = getValuesArray(workbook, sheetName)

  const instances = valuesArray.map((values: unknown[]) => {
    return values.reduce((box: CSVData, column: unknown, index: number) => {
      // 列単位で処理してきて、ヘッダの名前で代入する。
      box[headings[index]] = column

      return box
    }, {})
  })

  if (formatFunc) {
    return instances.map((instance) => formatFunc(instance))
  }

  return instances
}

/**
 * 指定したパスのcsvファイルをロードして、JSONオブジェクトとしてparseする。
 * 全行読み込んだら完了する Promise を返す。
 * @param filePath
 */
export const csv2json = async (filePath: string, encoding = 'Shift_JIS'): Promise<unknown[]> => {
  return await csvStream2json(fs.createReadStream(filePath), encoding)
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
export const csvStream2json = async (stream: NodeJS.ReadableStream, encoding = 'Shift_JIS'): Promise<unknown[]> => {
  return await new Promise<unknown[]>((resolve, reject) => {
    const datas: unknown[] = []

    stream
      .pipe(iconv.decodeStream(encoding))
      .pipe(iconv.encodeStream('utf-8'))
      .pipe(
        csv()
          .on('data', (data) => datas.push(Buffer.isBuffer(data) ? JSON.parse(data.toString()) : JSON.parse(data)))
          // .on('done', (error) => (error ? reject(error) : resolve(datas)))
          .on('error', (error) => reject(error)),
      )
      .on('end', () => resolve(datas))
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
  instances: unknown[],
  outputFullPath: string,
  templateFullPath = '',
  sheetName = 'Sheet1',
  converters?: Converters,
  applyStyles?: (instances: unknown[], workbook: XlsxPopulate.Workbook, sheetName: string) => void,
): Promise<string> => {
  logger.debug(`template path: ${templateFullPath}`)
  // console.log(instances[0])
  // console.table(instances)

  let headings: string[] = [] // ヘッダ名の配列
  let workbook: XlsxPopulate.Workbook
  const fileIsNew: boolean = templateFullPath === '' // templateが指定されない場合新規(fileIsNew = true)、そうでない場合テンプレファイルに出力

  if (!fileIsNew) {
    // 指定された場合は、一行目の文字列群を使ってプロパティを作成する
    workbook = await XlsxPopulate.fromFileAsync(templateFullPath)
    headings = getHeaders(workbook, sheetName)
  } else {
    // templateが指定されない場合は、空ファイルをつくり、オブジェクトのプロパティでダンプする。
    workbook = await XlsxPopulate.fromBlankAsync()
    if (instances.length > 0) {
      headings = Object.keys(instances[0] as CSVData)
    }
  }

  if (instances.length > 0) {
    const csvArrays: unknown[][] = createCsvArrays(headings, instances, converters)
    // console.table(csvArrays)
    const rowCount = instances.length
    const columnCount = headings.length
    const sheet = workbook.sheet(sheetName) ?? workbook.addSheet(sheetName)

    if (!fileIsNew && sheet.usedRange()) {
      sheet.usedRange()?.clear() // Excel上のデータを削除して。
    }
    sheet.cell('A1').value(csvArrays as unknown as string)

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

export const createWorkbook = async (path?: string): Promise<XlsxPopulate.Workbook> => {
  return (path != null) ? await XlsxPopulate.fromFileAsync(path) : await XlsxPopulate.fromBlankAsync()
}


export const toFileAsync = async (workbook: XlsxPopulate.Workbook, path: string): Promise<void> => {
  logger.debug(path)

  return await workbook.toFileAsync(path)

  // return toFullPath(path)
}



/**
 * 引数のJSON配列を、指定したテンプレートを用いて、指定したファイルに出力します。
 * @param instances JSON配列
 * @param templateFullPath 元にするテンプレートExcelのパス
 * @param sheetName テンプレートExcelのシート名(シート名で出力する)
 * @param applyStyles 出力時のExcelを書式フォーマットしたい場合に使用する。
 */
export const json2wookbook: (arg: {
  instances: unknown[];
  workbook: XlsxPopulate.Workbook;
  sheetName?: string;
  converters?: Converters;
  applyStyles?: (instances: unknown[], workbook: XlsxPopulate.Workbook, sheetName: string) => void;
}) => XlsxPopulate.Workbook = ({
  instances,
  workbook,
  sheetName = 'Sheet1',
  converters,
  applyStyles,
}): XlsxPopulate.Workbook => {

    // logger.debug(`template path: ${templateFullPath}`)
    // console.log(instances[0])
    // console.table(instances)

    let headings: string[] = [] // ヘッダ名の配列
    // let workbook: XlsxPopulate.Workbook
    // const fileIsNew: boolean = templateFullPath === '' // templateが指定されない場合新規(fileIsNew = true)、そうでない場合テンプレファイルに出力
    // if (fs.existsSync(outputFullPath)) {
    //   workbook = await XlsxPopulate.fromFileAsync(outputFullPath)
    //   if (instances.length > 0) {
    //     headings = Object.keys(instances[0] as CSVData)
    //   }
    // } else
    // if (!fileIsNew) {
    // 指定された場合は、一行目の文字列群を使ってプロパティを作成する
    // workbook = await XlsxPopulate.fromFileAsync(templateFullPath)
    // headings = getHeaders(workbook, sheetName)
    // } else {
    // templateが指定されない場合は、空ファイルをつくり、オブジェクトのプロパティでダンプする。
    // if (!workbook) {
    // workbook = await XlsxPopulate.fromBlankAsync()
    // }
    // }

    if (instances.length > 0) {
      headings = Object.keys(instances[0] as CSVData)
      const csvArrays: unknown[][] = createCsvArrays(headings, instances, converters)
      // console.table(csvArrays)
      const rowCount = instances.length
      const columnCount = headings.length
      const sheet = workbook.sheet(sheetName) ?? workbook.addSheet(sheetName)


      // if (!fileIsNew && sheet.usedRange()) {
      // sheet.usedRange()?.clear() // Excel上のデータを削除して。
      // }
      sheet.cell('A1').value(csvArrays as unknown as string)

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

    // logger.debug(outputFullPath)
    // await workbook.toFileAsync(outputFullPath)
    //
    // return toFullPath(outputFullPath)
    return workbook
  }

/**
 * 引数のJSON配列を、指定したテンプレートを用いて、指定したファイルに出力します。
 * @param instances JSON配列
 * @param sheetName テンプレートExcelのシート名(シート名で出力する)
 * @param applyStyles 出力時のExcelを書式フォーマットしたい場合に使用する。
 */
export const json2excelBlob = async (
  instances: unknown[],
  sheetName = 'Sheet1',
  converters?: Converters,
  applyStyles?: (instances: unknown[], workbook: XlsxPopulate.Workbook, sheetName: string) => void,
): Promise<Blob> => {
  let headings: string[] = [] // ヘッダ名の配列
  const workbook = await XlsxPopulate.fromBlankAsync()
  if (instances.length > 0) {
    headings = Object.keys(instances[0] as CSVData)
  }

  if (instances.length > 0) {
    const csvArrays: any[][] = createCsvArrays(headings, instances, converters)
    // console.table(csvArrays)
    const rowCount = instances.length
    const columnCount = headings.length
    const sheet = workbook.sheet(sheetName)

    sheet.cell('A1').value(csvArrays as unknown as string)

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

  const blob: Blob = await workbook.outputAsync() as Blob

  return blob
}

const toFullPath = (str: string): string => (path.isAbsolute(str) ? str : path.join(path.resolve(''), str))



// 自前実装
const createCsvArrays = (headings: string[], instances: unknown[], converters?: Converters): unknown[][] => {
  const csvArrays: unknown[][] = instances.map((tmpInstance: unknown): unknown[] => {
    const instance = tmpInstance as CSVData
    // console.log(instance)
    const csvArray = headings.reduce((box: unknown[], header: string): unknown[] => {
      // console.log(`${instance[header]}: ${instance[header] instanceof Object}`)
      // console.log(converters)
      // if (converters && converters[header]) {
      if (converters?.[header]) {
        // header名に合致するConverterがある場合はそれ優先で適用
        const converter = converters[header] as (value: unknown) => unknown
        box.push(converter(instance[header]))
      } else if (instance[header] instanceof Object) {
        // Converterがない場合は、文字列に変換
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
export const getHeaders = (workbook: XlsxPopulate.Workbook, sheetName: string): string[] => {
  const sheet = workbook.sheet(sheetName)

  if (sheet.usedRange()) {
    // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
    return sheet.usedRange()!.value().shift() as string[]
  }

  return []
}


// XlsxPopulate
export const getValuesArray = (workbook: XlsxPopulate.Workbook, sheetName: string): unknown[][] => {

  const sheet = workbook.sheet(sheetName)
  if (sheet.usedRange()) {
    // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
    const valuesArray: unknown[][] = sheet.usedRange()!.value()
    valuesArray.shift() // 先頭除去

    return valuesArray
  }

  return new Array([])
}
