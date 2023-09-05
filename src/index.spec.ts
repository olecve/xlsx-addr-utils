import {
  cellRangeArray,
  cellRangeIterator,
  columnRangeArray,
  columnRangeIterator,
  decrementColumn,
  decrementRow,
  incrementColumn,
  incrementRow,
  isValidCellAddress,
  isValidCellRange,
  rowRangeArray,
  rowRangeIterator,
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

describe('cell ranges functions', () => {
  it('cellRangeIterator', () => {
    expect([...cellRangeIterator('A1:A5')]).toStrictEqual(['A1', 'A2', 'A3', 'A4', 'A5'])
    expect([...cellRangeIterator('A1:B5')]).toStrictEqual(['A1', 'A2', 'A3', 'A4', 'A5', 'B1', 'B2', 'B3', 'B4', 'B5'])
    expect([...cellRangeIterator('A1:E1')]).toStrictEqual(['A1', 'B1', 'C1', 'D1', 'E1'])
    expect([...cellRangeIterator('Z1:AA2')]).toStrictEqual(['Z1', 'Z2', 'AA1', 'AA2'])

    expect(() => {
      cellRangeIterator('AA5').next()
    }).toThrow('Invalid range format: AA5')

    expect(() => {
      cellRangeIterator('A:A5').next()
    }).toThrow('Invalid range format: A:A5')

    expect(() => {
      cellRangeIterator('A5:A').next()
    }).toThrow('Invalid range format: A5:A')
  })

  it('cellRangeArray', () => {
    expect(cellRangeArray('A1:A5')).toStrictEqual(['A1', 'A2', 'A3', 'A4', 'A5'])
    expect(cellRangeArray('A1:B5')).toStrictEqual(['A1', 'A2', 'A3', 'A4', 'A5', 'B1', 'B2', 'B3', 'B4', 'B5'])
    expect(cellRangeArray('A1:E1')).toStrictEqual(['A1', 'B1', 'C1', 'D1', 'E1'])
    expect(cellRangeArray('Z1:AA2')).toStrictEqual(['Z1', 'Z2', 'AA1', 'AA2'])

    expect(() => {
      cellRangeArray('AA5')
    }).toThrow('Invalid range format: AA5')

    expect(() => {
      cellRangeArray('A:A5')
    }).toThrow('Invalid range format: A:A5')

    expect(() => {
      cellRangeArray('A5:A')
    }).toThrow('Invalid range format: A5:A')
  })
})

describe('column ranges functions', () => {
  it('columnRangeIterator', () => {
    expect([...columnRangeIterator('A1:A5')]).toStrictEqual(['A1:A5'])
    expect([...columnRangeIterator('A1:F1')]).toStrictEqual(['A1:A1', 'B1:B1', 'C1:C1', 'D1:D1', 'E1:E1', 'F1:F1'])
    expect([...columnRangeIterator('A1:F5')]).toStrictEqual(['A1:A5', 'B1:B5', 'C1:C5', 'D1:D5', 'E1:E5', 'F1:F5'])
    expect([...columnRangeIterator('Z1:AA5')]).toStrictEqual(['Z1:Z5', 'AA1:AA5'])

    expect(() => {
      columnRangeIterator('AA5').next()
    }).toThrow('Invalid range format: AA5')

    expect(() => {
      columnRangeIterator('A:A5').next()
    }).toThrow('Invalid range format: A:A5')

    expect(() => {
      columnRangeIterator('A5:A').next()
    }).toThrow('Invalid range format: A5:A')
  })

  it('columnRangeArray', () => {
    expect(columnRangeArray('A1:A5')).toStrictEqual(['A1:A5'])
    expect(columnRangeArray('A1:F1')).toStrictEqual(['A1:A1', 'B1:B1', 'C1:C1', 'D1:D1', 'E1:E1', 'F1:F1'])
    expect(columnRangeArray('A1:F5')).toStrictEqual(['A1:A5', 'B1:B5', 'C1:C5', 'D1:D5', 'E1:E5', 'F1:F5'])
    expect(columnRangeArray('Z1:AA5')).toStrictEqual(['Z1:Z5', 'AA1:AA5'])

    expect(() => {
      columnRangeArray('AA5')
    }).toThrow('Invalid range format: AA5')

    expect(() => {
      columnRangeArray('A:A5')
    }).toThrow('Invalid range format: A:A5')

    expect(() => {
      columnRangeArray('A5:A')
    }).toThrow('Invalid range format: A5:A')
  })
})

describe('row ranges functions', () => {
  it('rowRangeIterator', () => {
    expect([...rowRangeIterator('A1:A5')]).toStrictEqual(['A1:A1', 'A2:A2', 'A3:A3', 'A4:A4', 'A5:A5'])
    expect([...rowRangeIterator('A1:F1')]).toStrictEqual(['A1:F1'])
    expect([...rowRangeIterator('A1:F5')]).toStrictEqual(['A1:F1', 'A2:F2', 'A3:F3', 'A4:F4', 'A5:F5'])
    expect([...rowRangeIterator('Z1:AA5')]).toStrictEqual(['Z1:AA1', 'Z2:AA2', 'Z3:AA3', 'Z4:AA4', 'Z5:AA5'])

    expect(() => {
      rowRangeIterator('AA5').next()
    }).toThrow('Invalid range format: AA5')

    expect(() => {
      rowRangeIterator('A:A5').next()
    }).toThrow('Invalid range format: A:A5')

    expect(() => {
      rowRangeIterator('A5:A').next()
    }).toThrow('Invalid range format: A5:A')
  })

  it('rowRangeArray', () => {
    expect(rowRangeArray('A1:A5')).toStrictEqual(['A1:A1', 'A2:A2', 'A3:A3', 'A4:A4', 'A5:A5'])
    expect(rowRangeArray('A1:F1')).toStrictEqual(['A1:F1'])
    expect(rowRangeArray('A1:F5')).toStrictEqual(['A1:F1', 'A2:F2', 'A3:F3', 'A4:F4', 'A5:F5'])
    expect(rowRangeArray('Z1:AA5')).toStrictEqual(['Z1:AA1', 'Z2:AA2', 'Z3:AA3', 'Z4:AA4', 'Z5:AA5'])

    expect(() => {
      rowRangeArray('AA5')
    }).toThrow('Invalid range format: AA5')

    expect(() => {
      rowRangeArray('A:A5')
    }).toThrow('Invalid range format: A:A5')

    expect(() => {
      rowRangeArray('A5:A')
    }).toThrow('Invalid range format: A5:A')
  })
})
