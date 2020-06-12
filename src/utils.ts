import fs from 'fs'
import unzip from 'unzip'
import http from 'http'

// CSVは関係なくて、URL上のファイルをZip展開する
export const getAndExtract = (url: string): Promise<void> => {
  // https://www.npmjs.com/package/unzip
  return new Promise((resolve, reject) => {
    http.get(url, (res: http.IncomingMessage) => {
      res.pipe(unzip.Parse()).on('entry', (entry) => {
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
