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

function xlsxColumnToNumber(xlsxColumn: string) {
  return xlsxColumn.split('').reduce((result, currentValue) => result * 26 + parseInt(currentValue, 36) - 9, 0)
}

function numberToXlsxColumn(xlsxColumnAsNumber: number, columnName = '') {
  if (xlsxColumnAsNumber === 0) return columnName

  const remainder = (xlsxColumnAsNumber - 1) % 26
  const character = String.fromCharCode('A'.charCodeAt(0) + remainder)
  const newColumnName = character + columnName
  const newXlsxColumnAsNumber = Math.floor((xlsxColumnAsNumber - 1) / 26)

  return numberToXlsxColumn(newXlsxColumnAsNumber, newColumnName)
}

export function decrementColumn(cellAddress: string): string {
  const { columnPart, rowPart } = parseCellAddress(cellAddress)
  if (columnPart === 'A') throw new Error('Cannot decrement column "A"')
  const decrementedColumnPart = numberToXlsxColumn(xlsxColumnToNumber(columnPart) - 1)
  return decrementedColumnPart + rowPart
}

export function incrementColumn(cellAddress: string): string {
  const { columnPart, rowPart } = parseCellAddress(cellAddress)
  const incrementedColumn = numberToXlsxColumn(xlsxColumnToNumber(columnPart) + 1)
  return incrementedColumn + rowPart
}

export function decrementRow(cellAddress: string): string {
  const { columnPart, rowPart } = parseCellAddress(cellAddress)
  if (rowPart === '1') throw new Error('Cannot decrement row "1"')
  const decrementedRow = Number.parseInt(rowPart) - 1
  return columnPart + decrementedRow
}

export function incrementRow(cellAddress: string): string {
  const { columnPart, rowPart } = parseCellAddress(cellAddress)
  const incrementedRow = Number.parseInt(rowPart) + 1
  return columnPart + incrementedRow
}

export function* cellRangeIterator(range: string) {
  const [beginCellAddress, endCellAddress] = range.split(':')
  if (!beginCellAddress || !isValidCellAddress(beginCellAddress)) throw new Error(`Invalid range format: ${range}`)
  if (!endCellAddress || !isValidCellAddress(endCellAddress)) throw new Error(`Invalid range format: ${range}`)
  const begin = parseCellAddress(beginCellAddress)
  const end = parseCellAddress(endCellAddress)
  const beginColumnIndex = xlsxColumnToNumber(begin.columnPart)
  const beginRowIndex = Number.parseInt(begin.rowPart)
  const endColumnIndex = xlsxColumnToNumber(end.columnPart)
  const endRowIndex = Number.parseInt(end.rowPart)

  for (let column = beginColumnIndex; column <= endColumnIndex; column++) {
    for (let row = beginRowIndex; row <= endRowIndex; row++) {
      yield numberToXlsxColumn(column) + row
    }
  }
}

export function* columnRangeIterator(range: string) {
  const [beginCellAddress, endCellAddress] = range.split(':')
  if (!beginCellAddress || !isValidCellAddress(beginCellAddress)) throw new Error(`Invalid range format: ${range}`)
  if (!endCellAddress || !isValidCellAddress(endCellAddress)) throw new Error(`Invalid range format: ${range}`)
  const begin = parseCellAddress(beginCellAddress)
  const end = parseCellAddress(endCellAddress)
  const beginColumnIndex = xlsxColumnToNumber(begin.columnPart)
  const beginRowIndex = Number.parseInt(begin.rowPart)
  const endColumnIndex = xlsxColumnToNumber(end.columnPart)
  const endRowIndex = Number.parseInt(end.rowPart)

  for (let column = beginColumnIndex; column <= endColumnIndex; column++) {
    yield `${numberToXlsxColumn(column)}${beginRowIndex}:${numberToXlsxColumn(column)}${endRowIndex}`
  }
}
