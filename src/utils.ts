import * as fs from 'fs'
import { IncomingMessage } from 'http'
import axios, { AxiosResponse, RawAxiosRequestConfig } from 'axios'
import * as unzip from 'unzip-stream'

export const getAndExtract = async (url: string): Promise<string> => {
  // CSVは関係なくて、URL上のファイルをZip展開する
  const config: RawAxiosRequestConfig = {
    method: 'get',
    url,
    responseType: 'stream'
  }

  return await axios(config).then(async (response: AxiosResponse) => {
    return await new Promise<string>((resolve, reject) => {
      const stream: IncomingMessage = response.data as IncomingMessage
      stream.pipe(unzip.Parse())
        .on('entry', (entry: any): void => {
          const filePath = entry.path as string
          entry.pipe(fs.createWriteStream(filePath))
          entry.on('end', () => {
            resolve(filePath)
          })
        })
    })
  })





  // // https://www.npmjs.com/package/unzip
  // return await new Promise((resolve, reject) => {
  //   http.get(url, (res: http.IncomingMessage) => {
  //     res.pipe(unzip.Parse()).on('entry', (entry) => {
  //       entry.pipe(fs.createWriteStream(entry.path))
  //       entry.on('end', () => resolve())
  //     })
  //   })

  //   // const req = require('http').get(url, (res: ReadStream) => {
  //   //   res.pipe(unzip.Extract({ path: './' }))
  //   //   res.on('end', () => resolve())
  //   // })
  // })
}
