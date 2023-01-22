import * as path from 'path'
import { csv2json, json2excel } from '../commonUtils'


/**
 * Excel出力の様々なケース.
 *
┌─────────┬─────┬───────────────┬───────────┬─────────────────────┬───────────────────┬─────────────┬──────────────────────┬─────────────┬───────────────┬────────┬────────────────────┬───────────────┬──────────────────────────┐
│ (index) │ Id  │  MachineName  │ MachineId │      Username       │ RobotEnvironments │   Version   │         Name         │ HostingType │ ProvisionType │ UserId │ IsExternalLicensed │     Type      │           now            │
├─────────┼─────┼───────────────┼───────────┼─────────────────────┼───────────────────┼─────────────┼──────────────────────┼─────────────┼───────────────┼────────┼────────────────────┼───────────────┼──────────────────────────┤
│    0    │ '1' │   'WINDOWS'   │    '1'    │   'ad\\masatomi'    │      'test'       │ '19.10.4.0' │      'WINDOWS'       │ 'Standard'  │   'Manual'    │  '3'   │      'FALSE'       │ 'Development' │ 2020-06-12T05:35:01.181Z │
│    1    │ '2' │  'SQLServer'  │    '2'    │      'ad\\ur'       │      'test'       │ '19.10.4.0' │     'SQLServer'      │ 'Standard'  │   'Manual'    │  '6'   │      'FALSE'       │ 'Unattended'  │ 2020-06-12T05:35:01.181Z │
│    2    │ '3' │ 'WINDOWSESXI' │    '3'    │   'ad\\masatomi'    │      'test'       │     ''      │  'WINDOWSESXI_masa'  │ 'Standard'  │   'Manual'    │  '7'   │      'FALSE'       │ 'Development' │ 2020-06-12T05:35:01.181Z │
│    3    │ '4' │ 'WINDOWSESXI' │    '3'    │      'm-kino'       │      'test'       │ '20.4.0.0'  │ 'WINDOWSESXI_m-kino' │ 'Standard'  │   'Manual'    │  '8'   │      'FALSE'       │ 'Development' │ 2020-06-12T05:35:01.181Z │
│    4    │ '5' │   'WINDOWS'   │    '1'    │      'ad\\ur'       │      'test'       │ '19.10.4.0' │     'WINDOWS_UR'     │ 'Standard'  │   'Manual'    │  '9'   │       'TRUE'       │ 'Unattended'  │ 2020-06-12T05:35:01.181Z │
│    5    │ '6' │   'WINDOWS'   │    '1'    │      'm-kino'       │      'test'       │ '19.10.4.0' │   'WINDOWS_m-kino'   │ 'Standard'  │   'Manual'    │  '14'  │      'FALSE'       │ 'Development' │ 2020-06-12T05:35:01.181Z │
│    6    │ '7' │  'SQLServer'  │    '2'    │ 'ad\\administrator' │        ''         │     ''      │    'SQLServer_ad'    │ 'Standard'  │   'Manual'    │  '39'  │      'FALSE'       │ 'Development' │ 2020-06-12T05:35:01.181Z │
│    7    │ '8' │  'SQLServer'  │    '2'    │   'ad\\masatomi'    │        ''         │     ''      │   'SQLServer_masa'   │ 'Standard'  │   'Manual'    │  '40'  │      'FALSE'       │ 'Development' │ 2020-06-12T05:35:01.181Z │
└─────────┴─────┴───────────────┴───────────┴─────────────────────┴───────────────────┴─────────────┴──────────────────────┴─────────────┴───────────────┴────────┴────────────────────┴───────────────┴──────────────────────────┘
 *
 */
async function sample5() {
  let robots: unknown[] = await csv2json(path.join('src/samples', 'robotSample.csv'))
  // robots = robots.map((robot) => ({ ...robot, now: new Date() })) // 日付列を追加
  robots = robots.map((robot) => Object.assign({}, robot, { now: new Date() })) // 日付列を追加
  console.table(robots)

  // なにも考えずにダンプ
  json2excel(robots, 'output/robots.xlsx')

  // プロパティごとに、変換メソッドをかませたケース
  // nowとIdというプロパティには、変換methodを指定
  const converters = {
    now: (value: any) => value,
    Id: (value: any) => '0' + value,
  }
  json2excel(robots, 'output/robots1.xlsx', '', 'Sheet1', converters)

  // プロパティごとに、変換メソッドをかませたケース.
  // nowとIdというプロパティには、変換methodを指定
  // さらにその列(M列) に、日付フォーマットでExcel出力する
  const excelFormatter = (instances: any[], workbook: any, sheetName: string) => {
    const rowCount = instances.length
    const sheet = workbook.sheet(sheetName)
    sheet.range(`M2:M${rowCount + 1}`).style('numberFormat', 'yyyy/mm/dd hh:mm') // 書式: 日付+時刻
    // よくある整形パタン。
    // sheet.range(`C2:C${rowCount + 1}`).style('numberFormat', '@') // 書式: 文字(コレをやらないと、見かけ上文字だが、F2で抜けると数字になっちゃう)
    // sheet.range(`E2:F${rowCount + 1}`).style('numberFormat', 'yyyy/mm/dd') // 書式: 日付
    // sheet.range(`H2:H${rowCount + 1}`).style('numberFormat', 'yyyy/mm/dd hh:mm') // 書式: 日付+時刻
  }
  json2excel(robots, 'output/robots2.xlsx', '', 'Sheet1', converters, excelFormatter) // プロパティ指定で、変換をかける

  json2excel(robots, 'output/robots3.xlsx', path.join('src/samples', 'templateRobots.xlsx'), 'Sheet1') // テンプレを指定したケース
}

if (!module.parent) {
  ; (async () => {
    await sample5()
  })()
}
