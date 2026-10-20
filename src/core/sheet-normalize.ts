import type { WorkSheet } from 'xlsx'
import { utils } from 'xlsx'
import { COLUMN_TYPE, SUBJECT_ID, SUBJECT_ID_META, type JSONSheetRow, type JSONSheetRows } from './constants'
import { getColumnType, getSubjectId, normalizeColumnName } from './utils'

/** 0-based sheet normalize options (功能层；UI 稍后接范围选择) */
export interface SheetNormalizeOptions {
    /**
     * 作为表头的行索引（0-based，可多行，如双层表头 [1, 2]）。
     * 省略时自动检测。
     */
    headerRows?: number[]
    /**
     * 数据起始行（0-based）。省略时为 `max(headerRows) + 1` 或自动检测。
     */
    dataStartRow?: number
}

export interface NormalizeSheetResult {
    rows: JSONSheetRows
    headerRows: number[]
    dataStartRow: number
    columnNames: string[]
}

/**
 * 将 AOA 按 !merges 填充：合并区内非左上角的空单元格写入左上角值。
 */
export function fillMergedCells(aoa: unknown[][], merges: WorkSheet['!merges']): unknown[][] {
    const grid = aoa.map((row) => (Array.isArray(row) ? [...row] : []))
    if (!merges?.length) return grid

    let maxCol = 0
    for (const row of grid) {
        maxCol = Math.max(maxCol, row.length)
    }
    for (const range of merges) {
        maxCol = Math.max(maxCol, range.e.c + 1)
    }
    for (const row of grid) {
        while (row.length < maxCol) row.push(null)
    }

    for (const range of merges) {
        const top = range.s.r
        const left = range.s.c
        while (grid.length <= range.e.r) {
            grid.push(Array(maxCol).fill(null))
        }
        const value = grid[top]?.[left] ?? null
        for (let r = range.s.r; r <= range.e.r; r++) {
            if (!grid[r]) grid[r] = Array(maxCol).fill(null)
            while (grid[r].length < maxCol) grid[r].push(null)
            for (let c = range.s.c; c <= range.e.c; c++) {
                if (r === top && c === left) continue
                const cur = grid[r][c]
                if (cur == null || cur === '') {
                    grid[r][c] = value
                }
            }
        }
    }

    return grid
}

function nonEmptyCells(row: unknown[] | undefined): unknown[] {
    if (!row) return []
    return row.filter((c) => c != null && c !== '')
}

/** 整行标题 / 横幅（如「详细得分表」「第一次考试」） */
export function isTitleOrBannerRow(row: unknown[] | undefined): boolean {
    const cells = nonEmptyCells(row)
    if (cells.length === 0) return true

    const unique = [...new Set(cells.map((c) => String(c).trim()).filter(Boolean))]
    if (unique.length !== 1) return false

    const width = row?.length ?? 0
    // 同一文案铺满多列，或整行几乎只有一个标签
    if (cells.length >= Math.max(2, Math.ceil(width * 0.4))) return true
    if (cells.length === 1) return true
    return false
}

function isHeaderLikeValue(val: unknown): boolean {
    if (val == null || val === '') return false
    if (typeof val === 'number') return false

    const s = String(val).trim()
    if (!s) return false
    if (/^\d+(\.\d+)?$/.test(s)) return false

    if (getColumnType(normalizeColumnName(s)) !== COLUMN_TYPE.NONE) return true
    if (getSubjectId(normalizeColumnName(s)) !== SUBJECT_ID.NONE) return true

    const extras = ['总分', '客观分', '主观分', '分数', '得分', '排名', '成绩']
    if (extras.some((x) => s === x || s.includes(x))) return true

    // 非纯数字文本更可能是表头
    return true
}

function isDataLikeValue(val: unknown): boolean {
    if (val == null || val === '') return false
    if (typeof val === 'number') return true
    const s = String(val).trim()
    if (/^\d+(\.\d+)?$/.test(s)) return true
    return false
}

/**
 * 自动检测表头行与数据起始行（跳过顶部标题行）。
 */
export function detectHeaderRegion(grid: unknown[][]): { headerRows: number[]; dataStartRow: number } {
    if (!grid.length) {
        return { headerRows: [0], dataStartRow: 1 }
    }

    let row = 0
    while (row < grid.length && isTitleOrBannerRow(grid[row])) {
        row++
    }

    if (row >= grid.length) {
        return { headerRows: [0], dataStartRow: 1 }
    }

    const headerRows: number[] = []
    while (row < grid.length) {
        const current = grid[row] ?? []
        const nonEmpty = nonEmptyCells(current)
        if (nonEmpty.length === 0) {
            row++
            continue
        }

        let headerLike = 0
        let dataLike = 0
        for (const cell of nonEmpty) {
            if (isDataLikeValue(cell)) dataLike++
            else if (isHeaderLikeValue(cell)) headerLike++
        }

        // 数据行：数字/学号类占优
        if (dataLike > 0 && dataLike >= headerLike) {
            break
        }

        headerRows.push(row)
        row++

        // 表头不会无限长
        if (headerRows.length >= 10) break
    }

    if (headerRows.length === 0) {
        // 无标题时默认第一行表头
        const fallback = Math.min(row, grid.length - 1)
        return { headerRows: [Math.max(0, fallback)], dataStartRow: Math.max(0, fallback) + 1 }
    }

    return { headerRows, dataStartRow: row }
}

/**
 * 多层表头合成列名。
 * - 科目 + 总分/分数 → 科目名（如 计算机）
 * - 科目 + 客观分/主观分 → 科目-子列（如 计算机-客观分）
 * - 其余用 `-` 连接并去重相邻段
 */
export function composeColumnName(parts: unknown[]): string {
    const cleaned: string[] = []
    for (const p of parts) {
        if (p == null || p === '') continue
        const s = String(p).trim()
        if (!s) continue
        if (cleaned[cleaned.length - 1] === s) continue
        cleaned.push(s)
    }

    if (cleaned.length === 0) return ''
    if (cleaned.length === 1) return cleaned[0]

    let subjectDisplay: string | null = null
    for (const p of cleaned) {
        const id = getSubjectId(normalizeColumnName(p))
        if (id !== SUBJECT_ID.NONE) {
            subjectDisplay = SUBJECT_ID_META[id as keyof typeof SUBJECT_ID_META].display
            break
        }
    }

    const last = cleaned[cleaned.length - 1]
    const lastType = getColumnType(normalizeColumnName(last))
    const lastIsScoreLabel =
        lastType === COLUMN_TYPE.SCORE ||
        last === '总分' ||
        last === '分数' ||
        last === '得分'

    if (subjectDisplay) {
        if (lastIsScoreLabel) {
            return subjectDisplay
        }
        if (last !== subjectDisplay && getSubjectId(normalizeColumnName(last)) === SUBJECT_ID.NONE) {
            return `${subjectDisplay}-${last}`
        }
        return subjectDisplay
    }

    return cleaned.join('-')
}

export function uniquifyColumnNames(names: string[]): string[] {
    const seen = new Map<string, number>()
    return names.map((raw, index) => {
        let name = raw?.trim() || `列${index + 1}`
        const count = seen.get(name) ?? 0
        seen.set(name, count + 1)
        if (count === 0) return name
        return `${name}_${count}`
    })
}

export function buildColumnNames(grid: unknown[][], headerRows: number[]): string[] {
    if (!headerRows.length) return []

    let colCount = 0
    for (const r of headerRows) {
        colCount = Math.max(colCount, grid[r]?.length ?? 0)
    }
    for (const row of grid) {
        colCount = Math.max(colCount, row?.length ?? 0)
    }

    const names: string[] = []
    for (let c = 0; c < colCount; c++) {
        const parts: unknown[] = []
        for (const r of headerRows) {
            parts.push(grid[r]?.[c] ?? null)
        }
        names.push(composeColumnName(parts))
    }

    return uniquifyColumnNames(names)
}

export function gridToJsonRows(
    grid: unknown[][],
    columnNames: string[],
    dataStartRow: number
): JSONSheetRows {
    const rows: JSONSheetRows = []

    for (let r = dataStartRow; r < grid.length; r++) {
        const line = grid[r] ?? []
        const nonEmpty = nonEmptyCells(line)
        if (nonEmpty.length === 0) continue

        const obj: JSONSheetRow = {}
        let hasAny = false
        for (let c = 0; c < columnNames.length; c++) {
            const key = columnNames[c]
            const val = c < line.length ? line[c] : null
            if (val !== undefined) {
                obj[key] = val
                if (val != null && val !== '') hasAny = true
            }
        }
        if (hasAny) rows.push(obj)
    }

    return rows
}

/**
 * 将 worksheet 规范为 JSON 行：支持多层合并表头与可选行配置。
 */
export function normalizeWorksheet(
    sheet: WorkSheet,
    options: SheetNormalizeOptions = {}
): NormalizeSheetResult {
    const aoa = utils.sheet_to_json(sheet, {
        header: 1,
        defval: null,
        raw: true,
    }) as unknown[][]

    const grid = fillMergedCells(aoa, sheet['!merges'])

    let headerRows: number[]
    let dataStartRow: number

    if (options.headerRows && options.headerRows.length > 0) {
        headerRows = [...options.headerRows].sort((a, b) => a - b)
        dataStartRow =
            typeof options.dataStartRow === 'number'
                ? options.dataStartRow
                : Math.max(...headerRows) + 1
    } else {
        const detected = detectHeaderRegion(grid)
        headerRows = detected.headerRows
        dataStartRow =
            typeof options.dataStartRow === 'number'
                ? options.dataStartRow
                : detected.dataStartRow
    }

    const columnNames = buildColumnNames(grid, headerRows)
    const rows = gridToJsonRows(grid, columnNames, dataStartRow)

    return { rows, headerRows, dataStartRow, columnNames }
}
