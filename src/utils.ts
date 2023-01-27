import * as fs from 'fs'
import { IncomingMessage } from 'http'
import axios, { AxiosResponse, RawAxiosRequestConfig } from 'axios'
import * as unzip from 'unzip-stream'

export const getAndExtract = async (url: string): Promise<string> => {
  // CSVは関係なくて、URL上のファイルをZip展開する
  const config: RawAxiosRequestConfig = {
    method: 'get',
    url,
    responseType: 'stream',
  }

  return await axios(config).then(async (response: AxiosResponse) => {
    return await new Promise<string>((resolve, reject) => {
      const stream: IncomingMessage = response.data as IncomingMessage
      stream.pipe(unzip.Parse()).on('entry', (entry: unzip.Entry): void => {
        const filePath = entry.path
        entry.pipe(fs.createWriteStream(filePath))
        entry.on('end', () => resolve(filePath)).on('error', (error) => reject(error))
      })
    })
  })
  // https://www.npmjs.com/package/unzip
}
