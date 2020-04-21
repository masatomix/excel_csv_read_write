import fs, { WriteStream, ReadStream } from 'fs'
import iconv from 'iconv-lite'
import csv from 'csvtojson'
import unzip from 'unzip'
import http from 'http'

class CsvReaderSample {
  /**
   * 指定したパスのcsvファイルをロードして、JSONオブジェクトとしてparseする。
   * 全行読み込んだら完了する Promise を返す。
   * @param path
   */
  parse = (path: string): Promise<any[]> => {
    return new Promise((resolve, reject) => {
      let datas: any[] = []
      fs.createReadStream(path)
        .pipe(iconv.decodeStream('Shift_JIS'))
        .pipe(iconv.encodeStream('utf-8'))
        .pipe(csv().on('data', data => datas.push(JSON.parse(data))))
        .on('end', () => resolve(datas))
    })
  }

  getAndExtract = (url: string): Promise<void> => {
    // https://www.npmjs.com/package/unzip
    return new Promise((resolve, reject) => {
      http.get(url, (res: http.IncomingMessage) => {
        res.pipe(unzip.Parse()).on('entry', entry => {
          entry.pipe(fs.createWriteStream(entry.path))
          entry.on('end', () => resolve())
        })
      })

      // const req = require('http').get(url, (res: ReadStream) => {
      //   res.pipe(unzip.Extract({ path: './' }))
      //   res.on('end', () => resolve())
      // })
    })
  }
}

export = CsvReaderSample

if (!module.parent) {
  const reader = new CsvReaderSample()
  reader.getAndExtract('http://jusyo.jp/downloads/new/csv/csv_13tokyo.zip').then(() => {
    reader
      .parse('./13tokyo.csv')
      .then((results: any[]) => {
        // 郵便番号が「100-000x」のものに絞ってみた
        results = results.filter(address => address['郵便番号'].startsWith('100-000'))
        console.table(results)
        for (const address of results) {
          console.log(address)
        }
      })
      .finally(() => {
        if (fs.existsSync('./13tokyo.csv')) {
          fs.unlinkSync('./13tokyo.csv')
        } else {
          console.log('あれ？ない')
        }
      })
  })
}
// https://tsurutoro.com/pseudo_data/
