const cellAddressPattern = /^[A-Z]+[1-9]\d*$/
const cellRangePattern = /^[A-Z]+[1-9]\d*:[A-Z]+[1-9]\d*$/

export function isValidCellAddress(address: string): boolean {
  return cellAddressPattern.test(address)
}

export function isValidCellRange(range: string): boolean {
  return cellRangePattern.test(range)
}

export function setColumn(address: string, newColumn: number | string) {
  const rowMatch = address.match(/\d+/)
  if (!rowMatch) throw new Error(`Invalid address format: ${address}`)
  const rowPart = rowMatch[0]
  return newColumn + rowPart
}

export function setRow(address: string, newRow: number | string) {
  const columnMatch = address.match(/[A-Z]+/)
  if (!columnMatch) throw new Error(`Invalid address format: ${address}`)
  const columnPart = columnMatch[0]
  return columnPart + newRow
}
