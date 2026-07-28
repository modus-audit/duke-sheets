/* tslint:disable */
/* eslint-disable */

declare global {
    interface SymbolConstructor {
        readonly dispose: unique symbol;
    }
}

export interface JsRowCell {
    col: number;
    value: string;
    style?: any;
    mergeSpan?: JsMergeSpan;
    isMergedSecondary?: boolean;
    hyperlink?: any;
    comment?: any;
    formula?: string;
    image?: any;
}

export interface JsRow {
    index: number;
    cells: Array<JsRowCell>;
}

export interface JsMergeSpan {
    rowSpan: number;
    colSpan: number;
}

export interface JsRowsOptions {
    useFormattedValues?: boolean;
    useCalculatedValues?: boolean;
    includeStyles?: boolean;
    includeMergeInfo?: boolean;
    includeHyperlinks?: boolean;
    includeComments?: boolean;
    includeFormulas?: boolean;
    includeImages?: boolean;
    skipEmptyValues?: boolean;
    skipBlankValues?: boolean;
}

export class RowIterator implements IterableIterator<JsRow> {
    constructor(ws: Worksheet, opts?: JsRowsOptions, maxRow?: number);
    [Symbol.iterator](): RowIterator;
    next(): IteratorResult<JsRow>;
}

export interface Worksheet {
    iterateRows(opts?: JsRowsOptions): RowIterator;
}



export class CellValue {
    private constructor();
    free(): void;
    [Symbol.dispose](): void;
    asBoolean(): boolean | undefined;
    asError(): string | undefined;
    asNumber(): number | undefined;
    asText(): string | undefined;
    toJs(): any;
    toString(): string;
    readonly is_boolean: boolean;
    readonly is_empty: boolean;
    readonly is_error: boolean;
    readonly is_number: boolean;
    readonly is_text: boolean;
}

export class Workbook {
    free(): void;
    [Symbol.dispose](): void;
    addSheet(name: string): number;
    calculate(options?: any | null): any;
    defineName(name: string, refers_to: string): void;
    /**
     * Load a workbook from bytes, auto-detecting the format (XLSX or XLS).
     */
    static fromBytes(data: Uint8Array): Workbook;
    /**
     * Load a password-protected workbook from bytes, auto-detecting
     * the format (XLSX or XLS). `skipIntegrityCheck` skips the HMAC
     * integrity check on Agile-encrypted files (matches Office
     * behaviour); defaults to false.
     */
    static fromBytesWithPassword(data: Uint8Array, password: string, skip_integrity_check?: boolean | null): Workbook;
    static fromCsvString(csv: string): Workbook;
    getNamedRange(name: string): string | undefined;
    getSheet(index: number): Worksheet;
    getSheetByName(name: string): Worksheet;
    static loadCsvString(csv: string): Workbook;
    constructor();
    removeSheet(index: number): void;
    saveCsvString(): string;
    /**
     * Save the workbook as encrypted XLS bytes. `profile` selects
     * the FilePass variant; `null` defaults to RC4 CryptoAPI 128.
     * Valid values: `"rc4-cryptoapi"`, `"rc4-legacy"`, `"xor"`.
     * `keyBits` controls RC4 CryptoAPI key size (40 or 128). XOR is
     * not certified to interoperate with modern Excel.
     */
    saveXlsBytesEncrypted(password: string, profile?: string | null, key_bits?: number | null): Uint8Array;
    saveXlsbBytes(): Uint8Array;
    saveXlsxBytes(): Uint8Array;
    /**
     * Save the workbook as encrypted XLSX bytes. `profile` selects
     * the encryption variant; passing `null`/`undefined` uses the
     * Agile-256 default. Valid values: `"agile"`, `"standard"`.
     * `keyBits` and `spinCount` override the defaults where the
     * profile supports them.
     */
    saveXlsxBytesEncrypted(password: string, profile?: string | null, key_bits?: number | null, spin_count?: number | null): Uint8Array;
    sheetIndex(name: string): number | undefined;
    readonly activeSheet: number;
    readonly chartsheetCount: number;
    readonly chartsheets: any;
    readonly isEmpty: boolean;
    readonly namedRanges: any;
    readonly settings: any;
    readonly sheetCount: number;
    readonly sheetNames: string[];
    readonly sheetOrder: any;
    readonly totalSheetCount: number;
}

export class Worksheet {
    private constructor();
    free(): void;
    [Symbol.dispose](): void;
    getCalculatedValue(address: string): CellValue;
    getCalculatedValueAt(row: number, col: number): CellValue;
    getCell(address: string): CellValue;
    getCellAt(row: number, col: number): CellValue;
    getCellStyle(address: string): any;
    getCellStyleAt(row: number, col: number): any;
    getColumnOutlineLevel(col: number): number;
    getColumnWidth(col: number): number | undefined;
    getComment(address: string): any;
    getCommentAt(row: number, col: number): any;
    getFormattedValue(address: string): string;
    getFormattedValueAt(row: number, col: number): string;
    getFormulaAt(row: number, col: number): string | undefined;
    getHyperlink(address: string): any;
    getHyperlinkAt(row: number, col: number): any;
    getImageAt(row: number, col: number): any;
    getMergeSpan(row: number, col: number): any;
    getRowHeight(row: number): number | undefined;
    getRowOutlineLevel(row: number): number;
    getSpillSource(row: number, col: number): any;
    getTableByName(name: string): any;
    hasComment(address: string): boolean;
    hasCommentAt(row: number, col: number): boolean;
    isColumnCollapsed(col: number): boolean;
    isColumnHidden(col: number): boolean;
    isMergedSecondary(row: number, col: number): boolean;
    isRowCollapsed(row: number): boolean;
    isRowHidden(row: number): boolean;
    isSpillSource(row: number, col: number): boolean;
    isSpillTarget(row: number, col: number): boolean;
    mergeCells(range_str: string): void;
    setCell(address: string, value: any): void;
    setColumnWidth(col: number, width: number): void;
    setFormula(address: string, formula: string): void;
    setRowHeight(row: number, height: number): void;
    unmergeCells(range_str: string): boolean;
    usedRange(): any;
    readonly autoFilter: any;
    readonly cellCount: number;
    readonly chartCount: number;
    readonly chartExCount: number;
    readonly charts: any;
    readonly chartsEx: any;
    readonly colBreaks: any;
    readonly commentAuthors: string[];
    readonly commentCount: number;
    readonly comments: any;
    readonly conditionalFormatCount: number;
    readonly conditionalFormats: any;
    readonly dataValidationCount: number;
    readonly dataValidations: any;
    readonly date1904: boolean;
    readonly defaultColumnWidth: number;
    readonly defaultRowHeight: number;
    readonly formulaCells: any;
    /**
     * Get the number of formula cells in this worksheet.
     */
    readonly formulaCount: number;
    readonly freezePanes: any;
    readonly hyperlinkCount: number;
    readonly hyperlinks: any;
    readonly imageCount: number;
    readonly images: any;
    readonly isEmpty: boolean;
    readonly isSelected: boolean;
    readonly mergedRegions: any;
    readonly name: string;
    readonly pageSetup: any;
    readonly printArea: string | undefined;
    readonly protection: any;
    readonly repeatCols: Uint32Array | undefined;
    readonly repeatRows: Uint32Array | undefined;
    readonly rowBreaks: any;
    readonly selections: any;
    readonly splitPanes: any;
    readonly tabColor: any;
    readonly tableCount: number;
    readonly tables: any;
    readonly visibility: string;
    readonly zoomScale: number | undefined;
}

export function init(): void;

export type InitInput = RequestInfo | URL | Response | BufferSource | WebAssembly.Module;

export interface InitOutput {
    readonly memory: WebAssembly.Memory;
    readonly __wbg_cellvalue_free: (a: number, b: number) => void;
    readonly __wbg_workbook_free: (a: number, b: number) => void;
    readonly __wbg_worksheet_free: (a: number, b: number) => void;
    readonly cellvalue_asBoolean: (a: number) => number;
    readonly cellvalue_asError: (a: number) => [number, number];
    readonly cellvalue_asNumber: (a: number) => [number, number];
    readonly cellvalue_asText: (a: number) => [number, number];
    readonly cellvalue_is_boolean: (a: number) => number;
    readonly cellvalue_is_empty: (a: number) => number;
    readonly cellvalue_is_error: (a: number) => number;
    readonly cellvalue_is_number: (a: number) => number;
    readonly cellvalue_is_text: (a: number) => number;
    readonly cellvalue_toJs: (a: number) => any;
    readonly cellvalue_toString: (a: number) => [number, number];
    readonly init: () => void;
    readonly workbook_addSheet: (a: number, b: number, c: number) => [number, number, number];
    readonly workbook_calculate: (a: number, b: number) => [number, number, number];
    readonly workbook_defineName: (a: number, b: number, c: number, d: number, e: number) => [number, number];
    readonly workbook_fromBytes: (a: number, b: number) => [number, number, number];
    readonly workbook_fromBytesWithPassword: (a: number, b: number, c: number, d: number, e: number) => [number, number, number];
    readonly workbook_fromCsvString: (a: number, b: number) => [number, number, number];
    readonly workbook_getNamedRange: (a: number, b: number, c: number) => [number, number, number, number];
    readonly workbook_getSheet: (a: number, b: number) => [number, number, number];
    readonly workbook_getSheetByName: (a: number, b: number, c: number) => [number, number, number];
    readonly workbook_loadCsvString: (a: number, b: number) => [number, number, number];
    readonly workbook_new: () => number;
    readonly workbook_removeSheet: (a: number, b: number) => [number, number];
    readonly workbook_saveCsvString: (a: number) => [number, number, number, number];
    readonly workbook_saveXlsBytesEncrypted: (a: number, b: number, c: number, d: number, e: number, f: number) => [number, number, number, number];
    readonly workbook_saveXlsbBytes: (a: number) => [number, number, number, number];
    readonly workbook_saveXlsxBytes: (a: number) => [number, number, number, number];
    readonly workbook_saveXlsxBytesEncrypted: (a: number, b: number, c: number, d: number, e: number, f: number, g: number) => [number, number, number, number];
    readonly workbook_sheetCount: (a: number) => [number, number, number];
    readonly workbook_sheetNames: (a: number) => [number, number, number, number];
    readonly worksheet_getCalculatedValue: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_getCalculatedValueAt: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_getCell: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_getCellAt: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_getColumnWidth: (a: number, b: number) => [number, number, number, number];
    readonly worksheet_getImageAt: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_getRowHeight: (a: number, b: number) => [number, number, number, number];
    readonly worksheet_mergeCells: (a: number, b: number, c: number) => [number, number];
    readonly worksheet_name: (a: number) => [number, number, number, number];
    readonly worksheet_setCell: (a: number, b: number, c: number, d: any) => [number, number];
    readonly worksheet_setColumnWidth: (a: number, b: number, c: number) => [number, number];
    readonly worksheet_setFormula: (a: number, b: number, c: number, d: number, e: number) => [number, number];
    readonly worksheet_setRowHeight: (a: number, b: number, c: number) => [number, number];
    readonly worksheet_unmergeCells: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_usedRange: (a: number) => [number, number, number];
    readonly worksheet_autoFilter: (a: number) => [number, number, number];
    readonly worksheet_cellCount: (a: number) => [number, number, number];
    readonly worksheet_chartCount: (a: number) => [number, number, number];
    readonly worksheet_chartExCount: (a: number) => [number, number, number];
    readonly worksheet_charts: (a: number) => [number, number, number];
    readonly worksheet_chartsEx: (a: number) => [number, number, number];
    readonly worksheet_colBreaks: (a: number) => [number, number, number];
    readonly worksheet_commentAuthors: (a: number) => [number, number, number, number];
    readonly worksheet_commentCount: (a: number) => [number, number, number];
    readonly worksheet_comments: (a: number) => [number, number, number];
    readonly worksheet_conditionalFormatCount: (a: number) => [number, number, number];
    readonly worksheet_conditionalFormats: (a: number) => [number, number, number];
    readonly worksheet_dataValidationCount: (a: number) => [number, number, number];
    readonly worksheet_dataValidations: (a: number) => [number, number, number];
    readonly worksheet_date1904: (a: number) => [number, number, number];
    readonly worksheet_defaultColumnWidth: (a: number) => [number, number, number];
    readonly worksheet_defaultRowHeight: (a: number) => [number, number, number];
    readonly worksheet_formulaCells: (a: number) => [number, number, number];
    readonly worksheet_formulaCount: (a: number) => [number, number, number];
    readonly worksheet_freezePanes: (a: number) => [number, number, number];
    readonly worksheet_getCellStyle: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_getCellStyleAt: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_getColumnOutlineLevel: (a: number, b: number) => [number, number, number];
    readonly worksheet_getComment: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_getCommentAt: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_getFormattedValue: (a: number, b: number, c: number) => [number, number, number, number];
    readonly worksheet_getFormattedValueAt: (a: number, b: number, c: number) => [number, number, number, number];
    readonly worksheet_getFormulaAt: (a: number, b: number, c: number) => [number, number, number, number];
    readonly worksheet_getHyperlink: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_getHyperlinkAt: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_getMergeSpan: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_getRowOutlineLevel: (a: number, b: number) => [number, number, number];
    readonly worksheet_getRowsBatch: (a: number, b: number, c: number, d: any) => [number, number, number];
    readonly worksheet_getSpillSource: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_getTableByName: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_hasComment: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_hasCommentAt: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_hyperlinkCount: (a: number) => [number, number, number];
    readonly worksheet_hyperlinks: (a: number) => [number, number, number];
    readonly worksheet_imageCount: (a: number) => [number, number, number];
    readonly worksheet_images: (a: number) => [number, number, number];
    readonly worksheet_isColumnCollapsed: (a: number, b: number) => [number, number, number];
    readonly worksheet_isColumnHidden: (a: number, b: number) => [number, number, number];
    readonly worksheet_isEmpty: (a: number) => [number, number, number];
    readonly worksheet_isMergedSecondary: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_isRowCollapsed: (a: number, b: number) => [number, number, number];
    readonly worksheet_isRowHidden: (a: number, b: number) => [number, number, number];
    readonly worksheet_isSelected: (a: number) => [number, number, number];
    readonly worksheet_isSpillSource: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_isSpillTarget: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_mergedRegions: (a: number) => [number, number, number];
    readonly worksheet_pageSetup: (a: number) => [number, number, number];
    readonly worksheet_printArea: (a: number) => [number, number, number, number];
    readonly worksheet_protection: (a: number) => [number, number, number];
    readonly worksheet_repeatCols: (a: number) => [number, number, number, number];
    readonly worksheet_repeatRows: (a: number) => [number, number, number, number];
    readonly worksheet_rowBreaks: (a: number) => [number, number, number];
    readonly worksheet_selections: (a: number) => [number, number, number];
    readonly worksheet_splitPanes: (a: number) => [number, number, number];
    readonly worksheet_tabColor: (a: number) => [number, number, number];
    readonly worksheet_tableCount: (a: number) => [number, number, number];
    readonly worksheet_tables: (a: number) => [number, number, number];
    readonly worksheet_visibility: (a: number) => [number, number, number, number];
    readonly worksheet_zoomScale: (a: number) => [number, number, number];
    readonly workbook_activeSheet: (a: number) => [number, number, number];
    readonly workbook_chartsheetCount: (a: number) => [number, number, number];
    readonly workbook_chartsheets: (a: number) => [number, number, number];
    readonly workbook_isEmpty: (a: number) => [number, number, number];
    readonly workbook_namedRanges: (a: number) => [number, number, number];
    readonly workbook_settings: (a: number) => [number, number, number];
    readonly workbook_sheetIndex: (a: number, b: number, c: number) => [number, number, number];
    readonly workbook_sheetOrder: (a: number) => [number, number, number];
    readonly workbook_totalSheetCount: (a: number) => [number, number, number];
    readonly __wbindgen_malloc: (a: number, b: number) => number;
    readonly __wbindgen_realloc: (a: number, b: number, c: number, d: number) => number;
    readonly __wbindgen_exn_store: (a: number) => void;
    readonly __externref_table_alloc: () => number;
    readonly __wbindgen_externrefs: WebAssembly.Table;
    readonly __wbindgen_free: (a: number, b: number, c: number) => void;
    readonly __externref_table_dealloc: (a: number) => void;
    readonly __externref_drop_slice: (a: number, b: number) => void;
    readonly __wbindgen_start: () => void;
}

export type SyncInitInput = BufferSource | WebAssembly.Module;

/**
 * Instantiates the given `module`, which can either be bytes or
 * a precompiled `WebAssembly.Module`.
 *
 * @param {{ module: SyncInitInput }} module - Passing `SyncInitInput` directly is deprecated.
 *
 * @returns {InitOutput}
 */
export function initSync(module: { module: SyncInitInput } | SyncInitInput): InitOutput;

/**
 * If `module_or_path` is {RequestInfo} or {URL}, makes a request and
 * for everything else, calls `WebAssembly.instantiate` directly.
 *
 * @param {{ module_or_path: InitInput | Promise<InitInput> }} module_or_path - Passing `InitInput` directly is deprecated.
 *
 * @returns {Promise<InitOutput>}
 */
export default function __wbg_init (module_or_path?: { module_or_path: InitInput | Promise<InitInput> } | InitInput | Promise<InitInput>): Promise<InitOutput>;
