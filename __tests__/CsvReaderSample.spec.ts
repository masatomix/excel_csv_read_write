import CsvReaderSample from '../src/index'
import fs, { WriteStream } from 'fs'
import path from 'path'

describe('テスト', () => {
  let api: CsvReaderSample
  let instances: any[]

  const url: string = 'http://jusyo.jp/downloads/new/csv/csv_13tokyo.zip'
  const csvPath: string = './13tokyo.csv'
  const baseDir: string = path.resolve('')
  const fullPath: string = path.join(baseDir, csvPath)

  beforeEach(async () => {
    api = new CsvReaderSample()
    await api.getAndExtract(url)
  })

  it('getテスト1', async () => {
    expect(fs.existsSync(fullPath)).toBeTruthy()
    instances = await api.parse(fullPath)
    assertBasicArray(instances, 22)
    for (const obj of instances.filter(instance => instance['郵便番号'] === '100-0000')) {
      console.log(obj)
    }
  })

  it('getテスト2', async () => {
    expect(fs.existsSync(fullPath)).toBeTruthy()
    fs.unlinkSync(fullPath)
    await api.getAndExtract(url)
    expect(fs.existsSync(fullPath)).toBeTruthy()

    instances = await api.parse(fullPath)
    assertBasicArray(instances, 22)
    for (const obj of instances.filter(instance => instance['郵便番号'] === '100-0000')) {
      console.log(obj)
    }
  })

  afterEach(async () => {
    await new Promise((resolve, reject) => {
      fs.unlink(fullPath, error => {
        if (error) {
          console.log(error)
          reject(error)
        }
        resolve()
      })
    })
  })
})

function assertBasicArray(actualInstances: any[], expectedColumnCount: number) {
  expect(actualInstances.length).toBeGreaterThanOrEqual(1) // 1件以上はある

  for (const instance of actualInstances) {
    expect(Object.keys(instance).length).toBe(expectedColumnCount)
    expect(instance['住所CD']).not.toBeUndefined()
    // console.log(instance)
  }
}
