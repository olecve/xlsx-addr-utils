import {
  decrementColumn,
  decrementRow,
  incrementColumn,
  incrementRow,
  isValidCellAddress,
  isValidCellRange,
  setColumn,
  setRow
} from './index'

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

it('decrementColumn', () => {
  expect(decrementColumn('B10')).toBe('A10')
  expect(decrementColumn('AA10')).toBe('Z10')
  expect(decrementColumn('AB10')).toBe('AA10')
  expect(decrementColumn('ABA10')).toBe('AAZ10')
  expect(decrementColumn('AAA10')).toBe('ZZ10')
  expect(decrementColumn('AAAA10')).toBe('ZZZ10')

  expect(() => {
    decrementColumn('A1')
  }).toThrow('Cannot decrement column "A"')
})

it('incrementColumn', () => {
  expect(incrementColumn('A10')).toBe('B10')
  expect(incrementColumn('Z10')).toBe('AA10')
  expect(incrementColumn('AA10')).toBe('AB10')
  expect(incrementColumn('AAZ10')).toBe('ABA10')
  expect(incrementColumn('ZZ10')).toBe('AAA10')
  expect(incrementColumn('ZZZ10')).toBe('AAAA10')
})

it('decrementRow', () => {
  expect(decrementRow('A2')).toBe('A1')
  expect(decrementRow('B2')).toBe('B1')
  expect(decrementRow('A10')).toBe('A9')
  expect(decrementRow('A100')).toBe('A99')

  expect(() => {
    decrementRow('A1')
  }).toThrow('Cannot decrement row "1"')
})

it('incrementRow', () => {
  expect(incrementRow('A1')).toBe('A2')
  expect(incrementRow('B1')).toBe('B2')
  expect(incrementRow('A9')).toBe('A10')
  expect(incrementRow('A99')).toBe('A100')
})
