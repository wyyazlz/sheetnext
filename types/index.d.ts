// Type definitions for SheetNext
// Project: https://www.sheetnext.com
// Definitions by: SheetNext Team

export interface CellNum {
  r: number;
  c: number;
}

export interface RangeNum {
  s: CellNum;
  e: CellNum;
}

export type CellRef = string | CellNum;
export type RangeRef = string | RangeNum;

export interface SheetNextOptions {
  AI_URL?: string;
  AI_TOKEN?: string;
}

export interface ICellConfig {
  v?: string;
  w?: number;
  h?: number;
  b?: boolean;
  s?: number;
  fg?: string;
  a?: 'l' | 'r' | 'c';
  c?: string;
  mr?: number;
  mb?: number;
}

export interface SortItem {
  type: 'column' | 'row' | 'custom';
  order?: 'asc' | 'desc' | 'value';
  index: string;
  sortData?: any[];
  cb?: (rowsArray: Cell[][], sortIndex: number) => Cell[][];
}

export class Cell {
  isMerged: boolean;
  master: CellNum | null;
  hyperlink: { target?: string; location?: string; tooltip?: string } | null;
  dataValidation: any;
  editVal: string;
  readonly calcVal: any;
  readonly showVal: string;
  numFmt: string;
  font: any;
  alignment: any;
  border: any;
  fill: any;
}

export class Row {
  cells: Cell[];
  height: number;
  hidden: boolean;
  readonly rIndex: number;
  numFmt: string;
  font: any;
  alignment: any;
  border: any;
  fill: any;
  getCell(c: number): Cell;
}

export class Col {
  cells: Cell[];
  hidden: boolean;
  width: number;
  readonly cIndex: number;
  numFmt: string;
  font: any;
  alignment: any;
  border: any;
  fill: any;
  getCell(r: number): Cell;
}

export class Drawing {
  id: string;
  type: 'chart' | 'image' | 'connector' | 'shape';
  startCell: CellNum;
  offsetX: number;
  offsetY: number;
  width: number;
  height: number;
  option?: any;
  imageBase64?: string;
  updRender: boolean;
  updIndex(direction: 'up' | 'down' | 'top' | 'bottom'): void;
}

export class Sheet {
  name: string;
  hidden: boolean;
  merges: RangeNum[];
  defaultColWidth: number;
  defaultRowHeight: number;
  showGridLines: boolean;
  showRowColHeaders: boolean;
  activeCell: CellNum;
  activeAreas: RangeNum[];
  rows: Row[];
  cols: Col[];
  xSplit: number;
  ySplit: number;
  drawings: Drawing[];
  readonly rowCount: number;
  readonly colCount: number;

  getRow(r: number): Row;
  getCol(c: number): Col;
  getCell(r: number, c: number): Cell;
  getCellByStr(cellStr: string): Cell;
  rangeStrToNum(rangeStr: string): RangeNum;
  eachArea(rangeRef: RangeRef, callback: (r: number, c: number, index: number) => void, reverse?: boolean): void;
  showAllHidRows(): void;
  showAllHidCols(): void;
  addRows(startR: number, num: number): void;
  addCols(startC: number, num: number): void;
  delRows(startR: number, num: number): void;
  delCols(startC: number, num: number): void;
  mergeCells(rangeRef: RangeRef): void;
  unMergeCells(cellRef: CellRef): void;
  rangeSort(sortItems: SortItem[], range?: RangeRef): void;
  insertTable(data: (ICellConfig | string | number)[][], startCell: CellRef, globalConfig?: any): RangeNum;
  addDrawing(config: any): Drawing;
  getDrawingsByCell(cellRef: CellRef): Drawing[];
  removeDrawing(id: string): void;
}

export class Utils {
  static numToChar(num: number): string;
  static charToNum(char: string): number;
  static rangeNumToStr(rangeNum: RangeNum): string;
  static cellStrToNum(cellStr: string): CellNum;
  static cellNumToStr(cellNum: CellNum): string;
}

export class UndoRedo {
  undo(): void;
  redo(): void;
}

export default class SheetNext {
  constructor(dom: HTMLElement, options?: SheetNextOptions);

  activeSheet: Sheet;
  sheets: Sheet[];
  readonly sheetNames: string[];
  containerDom: HTMLElement;
  namespace: string;
  locked: boolean;
  Utils: typeof Utils;
  UndoRedo: UndoRedo;

  addSheet(name?: string): Sheet;
  delSheet(name: string): void;
  getSheetByName(name: string): Sheet | null;
  getVisibleSheetByIndex(index: number): Sheet;
  import(file: File): Promise<void>;
  export(type: 'XLSX'): void;
  r(): void;
  getLicenseInfo(): any;
}
