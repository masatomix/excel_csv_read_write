export interface CSVData {
    [key: string]: unknown;
}

/**
 *
 * @param arg
 * @returns
 */
export const isCSVData = (arg: unknown): arg is CSVData => {
    // const instance = arg as CSVData
    return true
}

/**
 * 
 * @param arg 
 * @returns 
 */
export const isCSVDatas = (arg: unknown): arg is CSVData[] => {
    const instances = arg as CSVData[]

    return instances.every((instance) => isCSVData(instance))
}


export interface Converters {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    [key: string]: (value?: any) => any
}