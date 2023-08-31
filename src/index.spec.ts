import { isValidCellAddress, isValidCellRange, setColumn, setRow } from './index'

it('isValidCellAddress', () => {
  expect(isValidCellAddress('A1')).toBe(true)
  expect(isValidCellAddress('A2')).toBe(true)
  expect(isValidCellAddress('Z1')).toBe(true)
  expect(isValidCellAddress('Z2')).toBe(true)
  expect(isValidCellAddress('AA11')).toBe(true)
  expect(isValidCellAddress('AB12')).toBe(true)
  expect(isValidCellAddress('ABC123')).toBe(true)
  expect(isValidCellAddress('A10')).toBe(true)

  expect(isValidCellAddress('A0')).toBe(false)
  expect(isValidCellAddress('A01')).toBe(false)
  expect(isValidCellAddress('A')).toBe(false)
  expect(isValidCellAddress('Z')).toBe(false)
  expect(isValidCellAddress('0')).toBe(false)
  expect(isValidCellAddress('1')).toBe(false)
})

it('isValidCellRange', () => {
  expect(isValidCellRange('A1:B2')).toBe(true)
  expect(isValidCellRange('AA10:AB20')).toBe(true)

  expect(isValidCellRange('A1:B')).toBe(false)
  expect(isValidCellRange('A:B1')).toBe(false)
  expect(isValidCellRange('A1:2')).toBe(false)
  expect(isValidCellRange('2:B2')).toBe(false)
})

it('setRow', () => {
  expect(setRow('A1', 99)).toBe('A99')
  expect(setRow('A1', '99')).toBe('A99')
})

it('setColumn', () => {
  expect(setColumn('A1', 'B')).toBe('B1')
  expect(setColumn('A1', 'AZ')).toBe('AZ1')
})
