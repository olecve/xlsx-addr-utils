import { CellAddress } from './types'

const cellRangePattern = /^[A-Z]+[1-9]\d*:[A-Z]+[1-9]\d*$/

export function isValidCellRange(range: string): boolean {
  return cellRangePattern.test(range)
}

function parceCellAddress(cellAddress: string): CellAddress {
  const match = cellAddress.match(/^([A-Z]+)([1-9]\d*)$/)
  if (!match || !match[1] || !match[2]) throw new Error(`Invalid cell address: ${cellAddress}`)
  const [_fullMatch, columnPart, rowPart] = match

  return {
    column: columnPart,
    row: Number.parseInt(rowPart)
  }
}

export function isValidCellAddress(address: string): boolean {
  try {
    parceCellAddress(address)
    return true
  } catch (_exception) {
    return false
  }
}

function cellAddressToString(cellAddress: CellAddress): string {
  return cellAddress.column + cellAddress.row
}

export function setColumn(cellAddress: string, newColumn: string): string {
  return cellAddressToString({ ...parceCellAddress(cellAddress), column: newColumn })
}

export function setRow(cellAddress: string, newRow: number | string): string {
  return cellAddressToString({ ...parceCellAddress(cellAddress), row: Number.parseInt(String(newRow)) })
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
  const { column, row } = parceCellAddress(cellAddress)
  if (column === 'A') throw new Error('Cannot decrement column "A"')
  const decrementedColumn = numberToXlsxColumn(xlsxColumnToNumber(column) - 1)
  return decrementedColumn + row
}

export function incrementColumn(cellAddress: string): string {
  const { column, row } = parceCellAddress(cellAddress)
  const incrementedColumn = numberToXlsxColumn(xlsxColumnToNumber(column) + 1)
  return incrementedColumn + row
}

export function decrementRow(cellAddress: string): string {
  const { column, row } = parceCellAddress(cellAddress)
  if (row === 1) throw new Error('Cannot decrement row "1"')
  const decrementedRow = row - 1
  return column + decrementedRow
}

export function incrementRow(cellAddress: string): string {
  const { column, row } = parceCellAddress(cellAddress)
  const incrementedRow = row + 1
  return column + incrementedRow
}

function parseRange(range: String) {
  const [beginCellAddress, endCellAddress] = range.split(':')
  if (!beginCellAddress || !isValidCellAddress(beginCellAddress)) throw new Error(`Invalid range format: ${range}`)
  if (!endCellAddress || !isValidCellAddress(endCellAddress)) throw new Error(`Invalid range format: ${range}`)
  const begin = parceCellAddress(beginCellAddress)
  const end = parceCellAddress(endCellAddress)
  const beginColumnIndex = xlsxColumnToNumber(begin.column)
  const beginRowIndex = begin.row
  const endColumnIndex = xlsxColumnToNumber(end.column)
  const endRowIndex = end.row
  return { begin, end, beginColumnIndex, beginRowIndex, endColumnIndex, endRowIndex }
}

export function* cellRangeIterator(range: string) {
  const { beginColumnIndex, beginRowIndex, endColumnIndex, endRowIndex } = parseRange(range)
  for (let column = beginColumnIndex; column <= endColumnIndex; column++) {
    for (let row = beginRowIndex; row <= endRowIndex; row++) {
      yield numberToXlsxColumn(column) + row
    }
  }
}

export function cellRangeArray(range: string): string[] {
  return [...cellRangeIterator(range)]
}

export function* columnRangeIterator(range: string) {
  const { beginColumnIndex, beginRowIndex, endColumnIndex, endRowIndex } = parseRange(range)
  for (let column = beginColumnIndex; column <= endColumnIndex; column++) {
    yield `${numberToXlsxColumn(column)}${beginRowIndex}:${numberToXlsxColumn(column)}${endRowIndex}`
  }
}

export function columnRangeArray(range: string): string[] {
  return [...columnRangeIterator(range)]
}

export function* rowRangeIterator(range: string) {
  const { begin, beginRowIndex, end, endRowIndex } = parseRange(range)
  for (let row = beginRowIndex; row <= endRowIndex; row++) {
    yield `${begin.column}${row}:${end.column}${row}`
  }
}

export function rowRangeArray(range: string): string[] {
  return [...rowRangeIterator(range)]
}
