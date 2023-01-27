export interface Address {
  // [Key: string]: string;
  住所CD: string
  都道府県CD: string
  市区町村CD: string
  町域CD: string
  郵便番号: string
  事業所フラグ: string
  廃止フラグ: string
  都道府県: string
  都道府県カナ: string
  市区町村: string
  市区町村カナ: string
  町域: string
  町域カナ: string
  町域補足: string
  京都通り名: string
  字丁目: string
  字丁目カナ: string
  補足: string
  事業所名: string
  事業所名カナ: string
  事業所住所: string
  新住所CD: string
}

/**
 *
 * @param arg
 * @returns
 */
export const isAddress = (arg: unknown): arg is Address => {
  const instance = arg as Address

  return instance.住所CD !== undefined
}

/**
 * 
 * @param arg 
 * @returns 
 */
export const isAddresses = (arg: unknown): arg is Address[] => {
  const instances = arg as Address[]

  return instances.every((instance) => isAddress(instance))
}
