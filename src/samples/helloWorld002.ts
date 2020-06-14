import path from 'path'
import fs from 'fs'
import { excel2json, json2excel, excelStream2json } from '../commonUtils'

/**
 * Excelファイルを読み込んで、ふたたびExcelに書き出すサンプル
 * Excelから読む場合は、型は適宜変換される、空文字はundefinedになるみたい
 * 
┌─────────┬───────────┬────────┬────────┬───────────┬────────────┬────────┬───────┬───────┬──────────┬────────┬────────┬───────────┬────────────┬───────────┬───────────┬───────────┬───────────┬───────────┬───────────┬───────────┬───────────┬───────────┐
│ (index) │   住所CD    │ 都道府県CD │ 市区町村CD │   町域CD    │    郵便番号    │ 事業所フラグ │ 廃止フラグ │ 都道府県  │  都道府県カナ  │  市区町村  │ 市区町村カナ │    町域     │    町域カナ    │   町域補足    │   京都通り名   │    字丁目    │   字丁目カナ   │    補足     │   事業所名    │  事
業所名カナ   │   事業所住所   │   新住所CD   │
├─────────┼───────────┼────────┼────────┼───────────┼────────────┼────────┼───────┼───────┼──────────┼────────┼────────┼───────────┼────────────┼───────────┼───────────┼───────────┼───────────┼───────────┼───────────┼───────────┼───────────┼───────────┤
│    0    │ 100000000 │   13   │ 13101  │ 131010000 │ '100-0000' │   0    │   0   │ '東京都' │ 'トウキョウト' │ '千代田区' │ 'チヨダク' │ undefined │ undefined  │ '（該当なし）'  │ undefined │ undefined │ undefined │ undefined │ undefined │ undefined │ undefined │ undefined │
│    1    │ 100000400 │   13   │ 13101  │ 131010006 │ '100-0004' │   0    │   0   │ '東京都' │ 'トウキョウト' │ '千代田区' │ 'チヨダク' │   '大手町'   │  'オオテマチ'   │ undefined │ undefined │ undefined │ undefined │ undefined │ undefined │ undefined │ undefined │ undefined │
│    2    │ 100000200 │   13   │ 13101  │ 131010039 │ '100-0002' │   0    │   0   │ '東京都' │ 'トウキョウト' │ '千代田区' │ 'チヨダク' │  '皇居外苑'   │ 'コウキョガイエン' │ undefined │ undefined │ undefined │ undefined │ undefined │ undefined │ undefined │ undefined │ undefined │
│    3    │ 100000100 │   13   │ 13101  │ 131010045 │ '100-0001' │   0    │   0   │ '東京都' │ 'トウキョウト' │ '千代田区' │ 'チヨダク' │   '千代田'   │   'チヨダ'    │ undefined │ undefined │ undefined │ undefined │ undefined │ undefined │ undefined │ undefined │ undefined │
│    4    │ 100000300 │   13   │ 13101  │ 131010051 │ '100-0003' │   0    │   0   │ '東京都' │ 'トウキョウト' │ '千代田区' │ 'チヨダク' │   '一ツ橋'   │  'ヒトツバシ'   │ undefined │ undefined │   '１丁目'   │ '０１チョウメ'  │ undefined │ undefined │ undefined │ undefined │ undefined │
│    5    │ 100000500 │   13   │ 13101  │ 131010055 │ '100-0005' │   0    │   0   │ '東京都' │ 'トウキョウト' │ '千代田区' │ 'チヨダク' │   '丸の内'   │  'マルノウチ'   │ undefined │ undefined │ undefined │ undefined │ undefined │ undefined │ undefined │ undefined │ undefined │
│    6    │ 100000600 │   13   │ 13101  │ 131010057 │ '100-0006' │   0    │   0   │ '東京都' │ 'トウキョウト' │ '千代田区' │ 'チヨダク' │   '有楽町'   │ 'ユウラクチョウ'  │ undefined │ undefined │ undefined │ undefined │ undefined │ undefined │ undefined │ undefined │ undefined │
└─────────┴───────────┴────────┴────────┴───────────┴────────────┴────────┴───────┴───────┴──────────┴────────┴────────┴───────────┴────────────┴───────────┴───────────┴───────────┴───────────┴───────────┴───────────┴───────────┴───────────┴───────────┘
 */
function sample2() {
  const resultPromise = excel2json('13tokyo.csv.xlsx')
  resultPromise.then((results) => {
    console.table(results)
    json2excel(results, path.join('output', '13tokyoResult.xlsx')).then((filePath) => console.log(filePath))
  })
}

function sample21() {
  const stream = fs.createReadStream('13tokyo.csv.xlsx')
  const resultPromise = excelStream2json(stream)
  resultPromise.then((results) => {
    console.table(results)
    json2excel(results, path.join('output', '13tokyoResult1.xlsx')).then((filePath) => console.log(filePath))
  })
}

if (!module.parent) {
  sample2()
}
