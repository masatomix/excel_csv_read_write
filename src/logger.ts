import * as Logger from 'bunyan'
import * as config from 'config'


interface LogConfig {
  logging: Logging[] | undefined
}

interface Logging {
  name: string
}

const isLogConfig = (arg: unknown): arg is LogConfig => {
  const conf = arg as LogConfig

  return conf.logging !== undefined
}

export const getLogger = (name: string): Logger => {
  if (isLogConfig(config)) {
    const settings = config.logging
    const target = settings?.find(setting => setting.name === name)

    if (target) {
      const logger = Logger.createLogger(target)
      logger.info(`[ ${name} ] のログをココに出力します。`)

      return logger
    }
  }
  const nullLogger = Logger.createLogger({
    name,
    level: 'error',
  })

  return nullLogger
}
