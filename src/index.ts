const cellAddressPattern = /^[A-Z]+[1-9]\d*$/
const cellRangePattern = /^[A-Z]+[1-9]\d*:[A-Z]+[1-9]\d*$/

export function isValidCellAddress(address: string): boolean {
  return cellAddressPattern.test(address)
}

export function isValidCellRange(range: string): boolean {
  return cellRangePattern.test(range)
}

export function setColumn(address: string, newColumn: string): string {
  const rowMatch = address.match(/\d+/)
  if (!rowMatch) throw new Error(`Invalid address format: ${address}`)
  const rowPart = rowMatch[0]
  return newColumn + rowPart
}

export function setRow(address: string, newRow: number | string): string {
  const columnMatch = address.match(/[A-Z]+/)
  if (!columnMatch) throw new Error(`Invalid address format: ${address}`)
  const columnPart = columnMatch[0]
  return columnPart + newRow
}

function parseCellAddress(cellAddress: string): { columnPart: string; rowPart: string } {
  const match = cellAddress.match(/^([A-Z]+)([1-9]\d*)$/)
  if (!match || !match[1] || !match[2]) throw new Error(`Invalid cell address: ${cellAddress}`)
  const [_fullMatch, columnPart, rowPart] = match
  return { columnPart, rowPart }
}

function excelColumnToNumber(excelColumn: string) {
  return excelColumn.split('').reduce((result, currentValue) => result * 26 + parseInt(currentValue, 36) - 9, 0)
}

function numberToExcelColumn(excelColumnAsNumber: number, columnName = '') {
  if (excelColumnAsNumber === 0) return columnName

  const remainder = (excelColumnAsNumber - 1) % 26
  const character = String.fromCharCode('A'.charCodeAt(0) + remainder)
  const newColumnName = character + columnName
  const newExcelColumnAsNumber = Math.floor((excelColumnAsNumber - 1) / 26)

  return numberToExcelColumn(newExcelColumnAsNumber, newColumnName)
}

export function decrementColumn(cellAddress: string): string {
  const { columnPart, rowPart } = parseCellAddress(cellAddress)
  if (columnPart === 'A') throw new Error('Cannot decrement column "A"')
  const decrementedColumnPart = numberToExcelColumn(excelColumnToNumber(columnPart) - 1)
  return decrementedColumnPart + rowPart
}

export function incrementColumn(cellAddress: string): string {
  const { columnPart, rowPart } = parseCellAddress(cellAddress)
  const decrementedColumnPart = numberToExcelColumn(excelColumnToNumber(columnPart) + 1)
  return decrementedColumnPart + rowPart
}
