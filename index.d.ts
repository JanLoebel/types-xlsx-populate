// Type definitions for XLSX-Populate
// Project: https://github.com/dtjohnson/xlsx-populate
// Definitions: https://github.com/DefinitelyTyped/DefinitelyTyped
// TypeScript Version: 3.4
export = XlsxPopulate
declare class XlsxPopulate {
  static MIME_TYPE: string
  static dateToNumber(date: Date): number
  static fromBlankAsync(): Promise<XlsxPopulate.Workbook>
  static fromDataAsync(data: string | number[] | ArrayBuffer | Uint8Array | Buffer | Blob | Promise<any>, opts?: object): Promise<XlsxPopulate.Workbook>
  static fromFileAsync(path: string, opts?: any): Promise<XlsxPopulate.Workbook>
  static numberToDate(number: number): Date
}

class StyleAble {
  style<K extends keyof XlsxPopulate.Style>(name: K): XlsxPopulate.Style[K]
  style<K extends keyof XlsxPopulate.Style>(names: K[]): { [key in K]: XlsxPopulate.Style[K] }
  style<K extends keyof XlsxPopulate.Style>(name: K, value: XlsxPopulate.Style[K]): this
  style(style: XlsxPopulate.Style): this
}

declare namespace XlsxPopulate {
  class Workbook {
    private constructor(...args: any[]): this;

    activeSheet(): Sheet
    activeSheet(sheet: Sheet | string | number): Workbook
    addSheet(name: string, indexOrBeforeSheet?: number | string | Sheet): Sheet
    definedName(name: string): undefined | string | Cell | Range | Row | Column
    definedName(name: string, refersTo: string | Cell | Range | Row | Column): Workbook
    deleteSheet(sheet: Sheet | string | number): Workbook
    find(pattern: string | RegExp, replacement?: string | Function): boolean
    moveSheet(sheet: Sheet | string | number, indexOrBeforeSheet?: number | string | Sheet): Workbook
    outputAsync(type?: string | Uint8Array | ArrayBuffer | Blob | Buffer): Promise<string | Uint8Array | ArrayBuffer | Blob | Buffer>
    outputAsync(opts?: object): string | Uint8Array | ArrayBuffer | Blob | Buffer
    sheet(sheetNameOrIndex: number | string): Sheet
    sheets(): Sheet[]
    property(name: string): any
    property(names: string[]): {[key: string]: any}
    property(properties: {[key: string]: any}): Workbook
    property(name: string, value: any): Workbook
    properties(): CoreProperties
    toFileAsync(path: string, opts?: object): Promise<void>
    cloneSheet(from: Sheet, name: string, indexOrBeforeSheet?: number | string | Sheet): Sheet
  }

  class Sheet {
    private constructor(...args: any[]): this;

    _rows: Row[] // private field

    active(): boolean
    active(active: boolean): Sheet
    activeCell(): Cell
    activeCell(cell: string | Cell): Sheet
    activeCell(rowNumber: number, columnNameOrNumber: string | number): Sheet
    cell(address: string): Cell
    cell(rowNumber: number, columnNameOrNumber: string | number): Cell
    column(columnNameOrNumber: string | number): Column
    definedName(name: string): undefined | string | Cell | Range | Row | Column
    definedName(name: string, refersTo: string | Cell | Range | Row | Column): Workbook
    delete(): Workbook
    find(pattern: string | RegExp, replacement?: string | Function): Array<Cell>
    gridLinesVisible(): boolean
    gridLinesVisible(selected: boolean): Sheet
    hidden(): boolean | string
    hidden(hidden: boolean): Sheet
    move(indexOrBeforeSheet?: number | string | Sheet): Sheet
    name(): string
    name(name: string): Sheet
    range(address: string): Range
    range(startCell: string | Cell, endCell: string | Cell): Range
    range(startRowNumber: number, startColumnNameOrNumber: string | number, endRowNumber: number, endColumnNameOrNumber: string | number): Range
    autoFilter(): Sheet
    autoFilter(range: Range): Sheet
    row(rowNumber: number): Row
    tabColor(): undefined | Color
    tabColor(): Color | string | number
    tabSelected(): boolean
    tabSelected(selected: boolean): Sheet
    usedRange(): Range | undefined
    workbook(): Workbook
    pageBreaks(): Object
    verticalPageBreaks(): PageBreaks
    horizontalPageBreaks(): PageBreaks
    hyperlink(address: string): string | undefined
    hyperlink(address: string, hyperlink: string, internal?: boolean): Sheet
    hyperlink(address: string, opts: object | Cell): Sheet
    printOptions(attributeName: string): boolean
    printOptions(attributeName: string, attributeEnabled: undefined | boolean): Sheet
    printGridLines(): boolean
    printGridLines(enabled: undefined | boolean): Sheet
    panes(opts : PanesOptions): Sheet
    freezePanes(xSplit: number, ySplit: number): Sheet
    freezePanes(topLeftCell: string): Sheet
    splitPanes(xSplit : number, ySplit : number): Sheet
    resetPanes(): Sheet
    pageMargins(attributeName: string): number
    pageMargins(attributeName: string, attributeStringValue: undefined | number | string): Sheet
    pageMarginsPreset(): string
    pageMarginsPreset(presetName: undefined | string): Sheet
    pageMarginsPreset(presetName: string, presetAttributes: object): Sheet
  }

  class Row extends StyleAble {
    private constructor(...args: any[]): this;

    _cells: Cell[];
    address(opts?: object): string
    cell(columnNameOrNumber: string | number ): Cell
    height(): undefined | number
    height(height: number): Row
    hidden(): boolean
    hidden(hidden: boolean): Row
    rowNumber(): number
    sheet(): Sheet
    workbook(): Workbook
    addPageBreak(): Row
  }

  type cellValue = string | boolean | number | Date | undefined | null;

  class Cell extends StyleAble {
    private constructor(...args: any[]): this;

    active(): boolean
    active(active: boolean): Cell
    address(opts?: object): string
    column(): Column
    clear(): Cell
    columnName(): number
    columnNumber(): number
    find(pattern: string | RegExp, replacement?: string | Function): boolean
    formula(): string
    formula(formula: string): Cell
    hyperlink(): string | undefined
    hyperlink(hyperlink: string | Cell | undefined): Cell
    hyperlink(opts: Object | Cell): Cell
    dataValidation(): object | undefined
    dataValidation(dataValidation: object | undefined): Cell
    tap(callback: Function): Cell
    thru(callback: Function): any
    rangeTo(cell: Cell | string): Range
    relativeCell(rowOffset: number, columnOffset: number): Cell
    row(): Row
    rowNumber(): number
    sheet(): Sheet
    value(): cellValue
    value(value: cellValue): Cell
    workbook(): Workbook
    addHorizontalPageBreak(): Cell
  }

  class Column extends StyleAble {
    private constructor(...args: any[]): this;

    address(opts?: object): string
    cell(rowNumber: number): Cell
    columnName(): string
    columnNumber(): number
    hidden(): boolean
    hidden(hidden: boolean): Column
    sheet(): Sheet
    width(): undefined | number
    width(width: number): Column
    workbook(): Workbook
    addPageBreak(): Column
  }

  class PanesOptions {
    activePane: string
    state: string
    topLeftCell: string
    xSplit: number
    ySplit: number
  }

  class CoreProperties {
    [key: string]: any
  }

  class Range extends StyleAble {
    address(opts?: object): string
    cell(ri: number, ci: number): Cell
    autoFilter(): Range
    cells(): [Cell][]
    clear(): Range
    endCell(): Cell
    forEach(callback: (cell: Cell, rowIndex: number, columnIndex: number, range: this) => void): Range
    formula(): string | undefined
    formula(formula: string): Range
    map<T>(callback: (cell: Cell, rowIndex: number, columnIndex: number, range: this) => T): T[][]
    merged(): boolean
    merged(merged: boolean): Range
    dataValidation(): object | undefined
    dataValidation(dataValidation: object | undefined): Range
    reduce<T>(callback: (obj: T, cell: Cell, rowIndex: number, columnIndex: number, range: this) => T, initialValue?: T): T
    sheet(): Sheet
    style(styles: {
      [K in keyof Style]: ((cell: Cell, rowIndex: number, columnIndex: number, range: this) => Style[K]) | Style[K][][] | Style[K]
    }): Range
    startCell(): Cell
    tap(callback: (cell: Cell, rowIndex: number, columnIndex: number, range: this) => void): Range
    thru<T>(callback: (cell: Cell, rowIndex: number, columnIndex: number, range: this) => T): T
    value(): cellValue[][]
    value(callback: (cell: Cell, rowIndex: number, columnIndex: number, range: this) => cellValue): Range
    value(values: cellValue[][]): Range
    value(value: cellValue): Range
    workbook(): Workbook
  }

  class PageBreaks {
    count: number
    list: any[]
    add(id: number): PageBreaks
    remove(index: number): PageBreaks
  }

  class Color {
    rgb?: string
    theme?: number
    tint?: number
  }

  type BorderStyle = 'hair' | 'dotted' | 'dashDotDot' | 'dashed' | 'mediumDashDotDot' | 'thin' | 'slantDashDot' | 'mediumDashDot' | 'mediumDashed' | 'medium' | 'thick' | 'double';

  class Style {
    bold?: boolean
    italic?: boolean
    underline?: boolean | string
    strikethrough?: boolean
    subscript?: boolean
    superscript?: boolean
    fontSize?: number
    fontFamily?: string
    fontColor?: Color | string
    horizontalAlignment?: 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed'
    justifyLastLine?: boolean
    indent?: number
    verticalAlignment?: 'top' | 'center' | 'bottom' | 'justify' | 'distributed'
    wrapText?: boolean
    shrinkToFit?: boolean
    textDirection?: 'left-to-right' | 'right-to-left'
    textRotation?: number
    angleTextCounterclockwise?: boolean
    angleTextClockwise?: boolean
    rotateTextUp?: boolean
    rotateTextDown?: boolean
    verticalText?: boolean
    fill?: SolidFill | PatternFill | GradientFill
    border?: Borders | Border
    borderColor?: Color | string | number
    borderStyle?: BorderStyle
    leftBorderColor?: Color | string | number
    rightBorderColor?: Color | string | number
    topBorderColor?: Color | string | number
    bottomBorderColor?: Color | string | number
    diagonalBorderColor?: Color | string | number
    leftBorderStyle?: BorderStyle
    rightBorderStyle?: BorderStyle
    topBorderStyle?: BorderStyle
    bottomBorderStyle?: BorderStyle
    diagonalBorderStyle?: BorderStyle
    diagonalBorderDirection?: 'up' | 'down' | 'both'
  }

  class SolidFill {
    type: 'solid'
    color: Color | string
  }

  class PatternFill {
    type: 'pattern'
    pattern: 'gray125' | 'darkGray' | 'mediumGray' | 'lightGray' | 'gray0625' | 'darkHorizontal' | 'darkVertical' | 'darkDown' | 'darkUp' | 'darkGrid' | 'darkTrellis' | 'lightHorizontal' | 'lightVertical' | 'lightDown' | 'lightUp' | 'lightGrid' | 'lightTrellis'
    foreground: Color | string
    background: Color | string
  }

  class Border {
    style: BorderStyle
    color: Color | string
    direction?: string
  }

  class Borders {
    left?: Border | BorderStyle
    right?: Border | BorderStyle
    top?: Border | BorderStyle
    bottom?: Border | BorderStyle
    diagonal?: Border | BorderStyle
  }

  class GradientFill {
    type:	'gradient'
    gradientType?: 'linear' | 'path'
    stops: {
      position: number
      color: Color | string
    }[]
    angle?: number
    left?: number
    right?: number
    top?: number
    bottom?: number
  }

  class FormulaError {
    error(): string
  }
  namespace FormulaError {
    const DIV0: FormulaError
    const NA: FormulaError
    const NAME: FormulaError
    const NULL: FormulaError
    const NUM: FormulaError
    const REF: FormulaError
    const VALUE: FormulaError
  }
}
