import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { read, utils, write } from 'xlsx'
import {
    buildColumnNames,
    composeColumnName,
    detectHeaderRegion,
    fillMergedCells,
    normalizeWorksheet,
} from './sheet-normalize'
import { ExcelAdapter } from './core'
import { SUBJECT_ID, type SubjectId } from './constants'

const examplesDir = resolve(import.meta.dirname, '../../examples')

function loadWorkbook(name: string) {
    return read(readFileSync(resolve(examplesDir, name)))
}

function fileFromPath(name: string): File {
    const buf = readFileSync(resolve(examplesDir, name))
    return new File([buf], name, {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    })
}

/** 内存构建 xlsx File，便于构造「缺学生」等场景 */
function fileFromAoa(
    name: string,
    sheets: { sheetName: string; aoa: unknown[][] }[]
): File {
    const wb = utils.book_new()
    for (const { sheetName, aoa } of sheets) {
        utils.book_append_sheet(wb, utils.aoa_to_sheet(aoa), sheetName)
    }
    const buf = write(wb, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer
    return new File([buf], name, {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    })
}

function rowsToObjects(rows: (string | object)[][]) {
    const header = rows[0] as string[]
    return rows.slice(1).map((row) => {
        const obj: Record<string, unknown> = {}
        header.forEach((h, idx) => {
            obj[h] = row[idx]
        })
        return obj
    })
}

function applyMergedHeadersSubjects(ea: ExcelAdapter) {
    const subjects = ea.getSubjectColumns()
    for (const [id, map] of subjects.entries()) {
        const parsed = ea.parseId(id) as { filename: string; sheetname: string }
        const name = parsed.filename

        // 数学成绩表（含 *-math / *-math-missing）
        if (name.includes('math')) {
            for (const col of map.keys()) {
                if (col === '总分') map.set(col, SUBJECT_ID.MATH)
                else map.set(col, SUBJECT_ID.NONE)
            }
            continue
        }

        // template / 花名册：无科目映射
        if (name.includes('template') || name.includes('fixed-roster') || name.includes('roster')) {
            for (const col of map.keys()) map.set(col, SUBJECT_ID.NONE)
            continue
        }

        // multi-header 计算机/语文（含 merged-column-headers.xlsx / *_missing.xlsx / other-scores）
        const isMultiHeaderScores =
            name === 'merged-column-headers.xlsx' ||
            name === 'merged-column-headers_missing.xlsx' ||
            (name.startsWith('merged-column-headers') &&
                !name.includes('math') &&
                !name.includes('template') &&
                !name.includes('result')) ||
            name.includes('scores')

        if (isMultiHeaderScores) {
            for (const col of map.keys()) {
                if (col === '计算机') map.set(col, SUBJECT_ID.COMPUTER)
                else if (col === '语文') map.set(col, SUBJECT_ID.CHINESE)
                else map.set(col, SUBJECT_ID.NONE)
            }
        }
    }
    ea.setSubjectColumns(subjects)
}

/** 从 result-missing 解析期望的科目分（null 表示缺分） */
function loadResultMissingExpected() {
    const wb = loadWorkbook('merged-column-headers_result-missing.xlsx')
    const sheet = wb.Sheets[wb.SheetNames[0]]
    // 结构与 template 相同：标题两行 + 表头，数据从 R3
    const aoa = utils.sheet_to_json(sheet, { header: 1, defval: null }) as unknown[][]
    // R2 = 表头
    const header = aoa[2] as string[]
    const col = (name: string) => header.indexOf(name)
    const iSerial = col('流水号')
    const iComputer = col('计算机')
    const iChinese = col('语文')
    const iMath = col('数学')

    return aoa.slice(3).filter((r) => r && r[iSerial] != null).map((r) => ({
        流水号: Number(r[iSerial]),
        计算机: r[iComputer] == null || r[iComputer] === '' ? null : Number(r[iComputer]),
        语文: r[iChinese] == null || r[iChinese] === '' ? null : Number(r[iChinese]),
        数学: r[iMath] == null || r[iMath] === '' ? null : Number(r[iMath]),
    }))
}

function expectScoreOrEmpty(actual: unknown, expected: number | null, label: string) {
    if (expected == null) {
        expect(actual === '' || actual == null, `${label} should be empty`).toBe(true)
    } else {
        expect(actual, label).toBe(expected)
    }
}

describe('sheet-normalize helpers', () => {
    it('composeColumnName: 科目+总分 → 科目名', () => {
        expect(composeColumnName(['计算机', '总分'])).toBe('计算机')
        expect(composeColumnName(['语文', '总分'])).toBe('语文')
    })

    it('composeColumnName: 科目+客观/主观 → 前缀列', () => {
        expect(composeColumnName(['计算机', '客观分'])).toBe('计算机-客观分')
        expect(composeColumnName(['语文', '主观分'])).toBe('语文-主观分')
    })

    it('detect + build multi-header sheet', () => {
        const wb = loadWorkbook('merged-column-headers.xlsx')
        const sheet = wb.Sheets[wb.SheetNames[0]]
        const aoa = utils.sheet_to_json(sheet, { header: 1, defval: null }) as unknown[][]
        const grid = fillMergedCells(aoa, sheet['!merges'])
        const { headerRows, dataStartRow } = detectHeaderRegion(grid)

        expect(headerRows).toEqual([1, 2])
        expect(dataStartRow).toBe(3)

        const cols = buildColumnNames(grid, headerRows)
        expect(cols).toEqual([
            '姓名',
            '流水号',
            '班级',
            '计算机',
            '计算机-客观分',
            '计算机-主观分',
            '语文',
            '语文-客观分',
            '语文-主观分',
        ])
    })
})

describe('normalizeWorksheet examples', () => {
    it('merged-column-headers.xlsx auto multi-header', () => {
        const wb = loadWorkbook('merged-column-headers.xlsx')
        const r = normalizeWorksheet(wb.Sheets[wb.SheetNames[0]])
        expect(r.columnNames).toContain('计算机')
        expect(r.columnNames).toContain('语文')
        expect(r.rows[0]).toMatchObject({
            姓名: '张三',
            流水号: 200101,
            班级: 2001,
            计算机: 130,
            语文: 110,
        })
        expect(r.rows).toHaveLength(6)
    })

    it('merged-column-headers.xlsx explicit headerRows', () => {
        const wb = loadWorkbook('merged-column-headers.xlsx')
        const r = normalizeWorksheet(wb.Sheets[wb.SheetNames[0]], {
            headerRows: [1, 2],
            dataStartRow: 3,
        })
        expect(r.headerRows).toEqual([1, 2])
        expect(r.rows[0]['计算机']).toBe(130)
    })

    it('flat math workbook', () => {
        const wb = loadWorkbook('merged-column-headers_math.xlsx')
        for (const name of wb.SheetNames) {
            const r = normalizeWorksheet(wb.Sheets[name])
            expect(r.columnNames).toEqual(['姓名', '流水号', '班级', '总分', '客观分', '主观分'])
            expect(r.rows.length).toBe(3)
        }
    })

    it('class9-math-title single title row', () => {
        const wb = loadWorkbook('class9-math-title.xlsx')
        const r = normalizeWorksheet(wb.Sheets[wb.SheetNames[0]])
        expect(r.columnNames).toEqual(['姓名', '班级', '总分', '客观分', '主观分'])
        expect(r.rows[0]).toMatchObject({ 姓名: '张三', 总分: 54 })
    })

    it('class1 flat multi-subject', () => {
        const wb = loadWorkbook('class1.xlsx')
        const r = normalizeWorksheet(wb.Sheets[wb.SheetNames[0]])
        expect(r.columnNames).toEqual(['姓名', '班级', '学号', '语文', '数学', '英语'])
        expect(r.rows).toHaveLength(3)
    })
})

describe('ExcelAdapter merge: merged-column-headers scenario', () => {
    it('主键流水号 + 数学/计算机/语文 → 与 result 分数一致', async () => {
        const ea = new ExcelAdapter()
        const errors = await ea.read([
            fileFromPath('merged-column-headers_math.xlsx'),
            fileFromPath('merged-column-headers.xlsx'),
        ])
        expect(errors).toHaveLength(0)

        const [shared] = ea.getSharedColumns()
        expect(shared).toEqual(expect.arrayContaining(['姓名', '流水号', '班级']))

        ea.setPrimaryColumns(['流水号'])

        const subjects = ea.getSubjectColumns()
        // 手动标记：math 的 总分→数学；multi 的 计算机/语文 应已自动识别
        for (const [id, map] of subjects.entries()) {
            const parsed = ea.parseId(id) as { filename: string; sheetname: string }
            if (parsed.filename.includes('math')) {
                for (const col of map.keys()) {
                    if (col === '总分') map.set(col, SUBJECT_ID.MATH)
                    else map.set(col, SUBJECT_ID.NONE)
                }
            }
            // multi-header 文件：仅保留 计算机/语文 作为科目（客观主观不选）
            if (parsed.filename === 'merged-column-headers.xlsx') {
                for (const col of map.keys()) {
                    if (col === '计算机') map.set(col, SUBJECT_ID.COMPUTER)
                    else if (col === '语文') map.set(col, SUBJECT_ID.CHINESE)
                    else map.set(col, SUBJECT_ID.NONE)
                }
            }
        }
        ea.setSubjectColumns(subjects)

        ea.addTotalScoreColumn({
            subjects: [SUBJECT_ID.COMPUTER, SUBJECT_ID.CHINESE, SUBJECT_ID.MATH] as SubjectId[],
            rank: true,
        })

        const [, rows, mergeErrors] = ea.merge()
        expect(mergeErrors).toHaveLength(0)

        const header = rows[0] as string[]
        expect(header).toEqual(
            expect.arrayContaining(['流水号', '计算机', '语文', '数学'])
        )

        const bySerial = new Map<number, Record<string, unknown>>()
        for (let i = 1; i < rows.length; i++) {
            const row = rows[i]
            const obj: Record<string, unknown> = {}
            header.forEach((h, idx) => {
                obj[h] = row[idx]
            })
            bySerial.set(Number(obj['流水号']), obj)
        }

        // 与 merged-column-headers_result.xlsx 数据行一致
        const expected: Record<number, { 计算机: number; 语文: number; 数学: number; 总分: number; 排名: number }> = {
            200101: { 计算机: 130, 语文: 110, 数学: 110, 总分: 350, 排名: 1 },
            200102: { 计算机: 81, 语文: 120, 数学: 120, 总分: 321, 排名: 2 },
            200103: { 计算机: 91, 语文: 87, 数学: 87, 总分: 265, 排名: 3 },
            200201: { 计算机: 120, 语文: 47, 数学: 47, 总分: 214, 排名: 5 },
            200202: { 计算机: 67, 语文: 57, 数学: 57, 总分: 181, 排名: 6 },
            200203: { 计算机: 88, 语文: 88, 数学: 88, 总分: 264, 排名: 4 },
        }

        expect(bySerial.size).toBe(6)

        for (const [serial, exp] of Object.entries(expected)) {
            const row = bySerial.get(Number(serial))
            expect(row, `missing serial ${serial}`).toBeTruthy()
            expect(row!['计算机']).toBe(exp.计算机)
            expect(row!['语文']).toBe(exp.语文)
            expect(row!['数学']).toBe(exp.数学)
        }

        // 总分/排名为公式单元格，检查公式存在
        const firstData = rows[1]
        const totalIdx = header.findIndex((h) => String(h).includes('总分'))
        const rankIdx = header.findIndex((h) => String(h).includes('排名'))
        expect(totalIdx).toBeGreaterThan(-1)
        expect(rankIdx).toBeGreaterThan(-1)
        expect((firstData[totalIdx] as { f?: string }).f).toMatch(/SUM/i)
        expect((firstData[rankIdx] as { f?: string }).f).toMatch(/RANK/i)
    })

    it('setSheetNormalizeOptions 可指定 headerRows', async () => {
        const ea = new ExcelAdapter()
        await ea.read([fileFromPath('merged-column-headers.xlsx')])
        const id = ea.getSheetIds()[0]

        ea.setSheetNormalizeOptions(id, { headerRows: [1, 2], dataStartRow: 3 })
        const sheet = ea.getNormalizedSheet(id)!
        expect(Object.keys(sheet.rows[0])).toEqual(
            expect.arrayContaining(['计算机', '语文', '流水号'])
        )
        expect(sheet.rows[0]['计算机']).toBe(130)
    })
})

describe('ExcelAdapter fixed source row order', () => {
    it('fixed=完整名单时，math 缺部分学生 → 顺序不变且数学为空', async () => {
        // fixed 名单顺序（故意与成绩表不同）
        const fixedOrder = [200203, 200101, 200102, 200103, 200201, 200202]
        const fixed = fileFromAoa('fixed-roster.xlsx', [
            {
                sheetName: '名单',
                aoa: [
                    ['姓名', '流水号', '班级'],
                    ['王五', 200203, 2002],
                    ['张三', 200101, 2001],
                    ['李四', 200102, 2001],
                    ['王五', 200103, 2001],
                    ['赵六', 200201, 2002],
                    ['李四', 200202, 2002],
                ],
            },
        ])

        // math 故意缺少 200102、200201
        const mathPartial = fileFromAoa('partial-math.xlsx', [
            {
                sheetName: '成绩',
                aoa: [
                    ['姓名', '流水号', '班级', '总分'],
                    ['张三', 200101, 2001, 110],
                    ['王五', 200103, 2001, 87],
                    ['李四', 200202, 2002, 57],
                    ['王五', 200203, 2002, 88],
                ],
            },
        ])

        // 计算机/语文完整
        const scores = fileFromAoa('other-scores.xlsx', [
            {
                sheetName: '成绩',
                aoa: [
                    ['姓名', '流水号', '班级', '计算机', '语文'],
                    ['张三', 200101, 2001, 130, 110],
                    ['李四', 200102, 2001, 81, 120],
                    ['王五', 200103, 2001, 91, 87],
                    ['赵六', 200201, 2002, 120, 47],
                    ['李四', 200202, 2002, 67, 57],
                    ['王五', 200203, 2002, 88, 88],
                ],
            },
        ])

        const ea = new ExcelAdapter()
        // 先读成绩再读 fixed，若无 fixed 会按成绩插入序；有 fixed 应被覆盖
        await ea.read([mathPartial, scores, fixed])
        ea.setPrimaryColumns(['流水号'])
        ea.setFixedSource('fixed-roster.xlsx')

        const subjects = ea.getSubjectColumns()
        for (const [id, map] of subjects.entries()) {
            const parsed = ea.parseId(id) as { filename: string }
            if (parsed.filename.includes('math')) {
                for (const col of map.keys()) {
                    map.set(col, col === '总分' ? SUBJECT_ID.MATH : SUBJECT_ID.NONE)
                }
            } else if (parsed.filename.includes('scores')) {
                for (const col of map.keys()) {
                    if (col === '计算机') map.set(col, SUBJECT_ID.COMPUTER)
                    else if (col === '语文') map.set(col, SUBJECT_ID.CHINESE)
                    else map.set(col, SUBJECT_ID.NONE)
                }
            } else {
                for (const col of map.keys()) map.set(col, SUBJECT_ID.NONE)
            }
        }
        ea.setSubjectColumns(subjects)
        ea.addTotalScoreColumn({
            subjects: [SUBJECT_ID.COMPUTER, SUBJECT_ID.CHINESE, SUBJECT_ID.MATH],
            rank: false,
        })

        const [, rows, mergeErrors] = ea.merge()
        expect(mergeErrors).toHaveLength(0)

        const data = rowsToObjects(rows as (string | object)[][])
        expect(data.map((r) => Number(r['流水号']))).toEqual(fixedOrder)

        // 有数学成绩的
        expect(data.find((r) => r['流水号'] === 200101)!['数学']).toBe(110)
        expect(data.find((r) => r['流水号'] === 200203)!['数学']).toBe(88)

        // math 缺失：数学列为空字符串（merge 默认 fill）
        const missingMath = [200102, 200201]
        for (const serial of missingMath) {
            const row = data.find((r) => r['流水号'] === serial)!
            expect(row['数学'], `serial ${serial} math should be empty`).toBe('')
            // 其他科仍有
            expect(row['计算机']).not.toBe('')
            expect(row['语文']).not.toBe('')
        }
    })

    it('fixed=multi-header 文件名时顺序与该表一致，且仅含 fixed 学生', async () => {
        const ea = new ExcelAdapter()
        await ea.read([
            fileFromPath('merged-column-headers_math.xlsx'),
            fileFromPath('merged-column-headers.xlsx'),
        ])
        ea.setPrimaryColumns(['流水号'])
        ea.setFixedSource('merged-column-headers.xlsx')
        applyMergedHeadersSubjects(ea)
        ea.addTotalScoreColumn({
            subjects: [SUBJECT_ID.COMPUTER, SUBJECT_ID.CHINESE, SUBJECT_ID.MATH],
            rank: true,
        })

        const [, rows] = ea.merge()
        const data = rowsToObjects(rows as (string | object)[][])

        // multi-header 表内顺序
        expect(data.map((r) => Number(r['流水号']))).toEqual([
            200101, 200102, 200103, 200201, 200202, 200203,
        ])
        expect(data).toHaveLength(6)
        expect(data[0]['计算机']).toBe(130)
        expect(data[0]['数学']).toBe(110)
    })

    it('fixed 无效时 merge 抛错', async () => {
        const ea = new ExcelAdapter()
        await ea.read([fileFromPath('merged-column-headers_math.xlsx')])
        ea.setPrimaryColumns(['流水号'])
        ea.setFixedSource('not-exist.xlsx')
        expect(() => ea.merge()).toThrow(/fixed source not found/)
    })

    it('examples merged-column-headers_*-missing: fixed=template，缺分留空且顺序不变', async () => {
        const ea = new ExcelAdapter()
        const errors = await ea.read([
            fileFromPath('merged-column-headers_math-missing.xlsx'),
            fileFromPath('merged-column-headers_missing.xlsx'),
            fileFromPath('merged-column-headers_template-missing.xlsx'),
        ])
        expect(errors).toHaveLength(0)

        ea.setPrimaryColumns(['流水号'])
        // 行序与名单以 template 为准
        ea.setFixedSource('merged-column-headers_template-missing.xlsx')
        applyMergedHeadersSubjects(ea)
        // 与 result-missing 一致：该总分列合计≤0 不写总分/排名
        ea.addTotalScoreColumn({
            subjects: [SUBJECT_ID.COMPUTER, SUBJECT_ID.CHINESE, SUBJECT_ID.MATH],
            rank: true,
            skipAtOrBelow: 0,
        })

        const [, rows, mergeErrors] = ea.merge()
        expect(mergeErrors).toHaveLength(0)

        const header = rows[0] as string[]
        const totalIdx = header.findIndex((h) => String(h).includes('总分'))
        const rankIdx = header.findIndex((h) => String(h).includes('排名'))
        expect(totalIdx).toBeGreaterThan(-1)
        expect(rankIdx).toBeGreaterThan(-1)

        const data = rowsToObjects(rows as (string | object)[][])
        const expected = loadResultMissingExpected()

        // 顺序 = template / result-missing
        expect(data.map((r) => Number(r['流水号']))).toEqual(expected.map((e) => e.流水号))
        expect(data).toHaveLength(6)

        for (const exp of expected) {
            const row = data.find((r) => Number(r['流水号']) === exp.流水号)
            expect(row, `missing serial ${exp.流水号}`).toBeTruthy()
            expectScoreOrEmpty(row!['计算机'], exp.计算机, `${exp.流水号} 计算机`)
            expectScoreOrEmpty(row!['语文'], exp.语文, `${exp.流水号} 语文`)
            expectScoreOrEmpty(row!['数学'], exp.数学, `${exp.流水号} 数学`)
        }

        // 200102：三科皆空 → 总分/排名也应为空（不写公式）
        const emptyRow = data.find((r) => Number(r['流水号']) === 200102)!
        expect(emptyRow['计算机']).toBe('')
        expect(emptyRow['语文']).toBe('')
        expect(emptyRow['数学']).toBe('')
        const emptyTotal = emptyRow[header[totalIdx]]
        const emptyRank = emptyRow[header[rankIdx]]
        expect(emptyTotal === '' || emptyTotal == null).toBe(true)
        expect(emptyRank === '' || emptyRank == null).toBe(true)

        // 有分的人：科目分 + 总分/排名公式
        const r101 = data.find((r) => Number(r['流水号']) === 200101)!
        expect(r101).toMatchObject({ 计算机: 130, 语文: 110, 数学: 110 })
        expect((r101[header[totalIdx]] as { f?: string }).f).toMatch(/SUM/i)
        expect((r101[header[rankIdx]] as { f?: string }).f).toMatch(/RANK/i)

        expect(data.find((r) => Number(r['流水号']) === 200103)!['数学']).toBe(87)
    })

    it('examples merged-column-headers_*-missing: 无 fixed 时按成绩表并集，200102 三科仍空', async () => {
        const ea = new ExcelAdapter()
        await ea.read([
            fileFromPath('merged-column-headers_missing.xlsx'),
            fileFromPath('merged-column-headers_math-missing.xlsx'),
        ])
        ea.setPrimaryColumns(['流水号'])
        applyMergedHeadersSubjects(ea)

        const [, rows] = ea.merge()
        const bySerial = new Map(
            rowsToObjects(rows as (string | object)[][]).map((r) => [Number(r['流水号']), r])
        )

        expect(bySerial.get(200101)).toMatchObject({ 计算机: 130, 语文: 110, 数学: 110 })
        expect(bySerial.get(200103)).toMatchObject({ 计算机: 91, 语文: 87, 数学: 87 })
        // multi-header 有 200102 主键行但成绩为空；math 无此人
        const r102 = bySerial.get(200102)
        expect(r102).toBeTruthy()
        expect(r102!['计算机'] === '' || r102!['计算机'] == null).toBe(true)
        expect(r102!['语文'] === '' || r102!['语文'] == null).toBe(true)
        expect(r102!['数学'] === '' || r102!['数学'] == null).toBe(true)
        // math-missing 无 李四 200102，有其余人
        expect(bySerial.get(200202)!['数学']).toBe(57)
        expect(bySerial.has(200102) && bySerial.get(200102)!['数学'] !== 57).toBe(true)
    })

    it('无 fixed 时不丢成绩表独有学生（对照）', async () => {
        const roster = fileFromAoa('small-roster.xlsx', [
            {
                sheetName: 's',
                aoa: [
                    ['流水号', '姓名'],
                    [1, '甲'],
                    [2, '乙'],
                ],
            },
        ])
        const math = fileFromAoa('math-extra.xlsx', [
            {
                sheetName: 's',
                aoa: [
                    ['流水号', '姓名', '总分'],
                    [1, '甲', 90],
                    [2, '乙', 80],
                    [3, '丙', 70], // roster 没有
                ],
            },
        ])

        const ea = new ExcelAdapter()
        await ea.read([roster, math])
        ea.setPrimaryColumns(['流水号'])

        const subjects = ea.getSubjectColumns()
        for (const [id, map] of subjects.entries()) {
            const parsed = ea.parseId(id) as { filename: string }
            for (const col of map.keys()) {
                if (parsed.filename.includes('math') && col === '总分') {
                    map.set(col, SUBJECT_ID.MATH)
                } else {
                    map.set(col, SUBJECT_ID.NONE)
                }
            }
        }
        ea.setSubjectColumns(subjects)

        // 无 fixed：丙应出现
        const [, rowsNoFixed] = ea.merge()
        const serialsNoFixed = rowsToObjects(rowsNoFixed as (string | object)[][]).map((r) =>
            Number(r['流水号'])
        )
        expect(serialsNoFixed).toContain(3)

        // 有 fixed=roster：只有甲乙，顺序 1,2；丙不出现
        ea.setFixedSource('small-roster.xlsx')
        const [, rowsFixed] = ea.merge()
        const dataFixed = rowsToObjects(rowsFixed as (string | object)[][])
        expect(dataFixed.map((r) => Number(r['流水号']))).toEqual([1, 2])
        expect(dataFixed.find((r) => r['流水号'] === 1)!['数学']).toBe(90)
    })
})

describe('ExcelAdapter TotalScoreColumn.skipAtOrBelow (column-level)', () => {
    function totalRankCells(rows: (string | object)[][]) {
        const header = rows[0] as string[]
        const totalKey = header.find((h) => String(h).includes('总分'))!
        const rankKey = header.find((h) => String(h).includes('排名'))
        const data = rowsToObjects(rows)
        return { header, totalKey, rankKey, data }
    }

    it('未设置 skipAtOrBelow：全空分学生仍写总分/排名公式', async () => {
        const ea = new ExcelAdapter()
        await ea.read([
            fileFromPath('merged-column-headers_math-missing.xlsx'),
            fileFromPath('merged-column-headers_missing.xlsx'),
            fileFromPath('merged-column-headers_template-missing.xlsx'),
        ])
        ea.setPrimaryColumns(['流水号'])
        ea.setFixedSource('merged-column-headers_template-missing.xlsx')
        applyMergedHeadersSubjects(ea)
        ea.addTotalScoreColumn({
            subjects: [SUBJECT_ID.COMPUTER, SUBJECT_ID.CHINESE, SUBJECT_ID.MATH],
            rank: true,
            skipAtOrBelow: null,
        })

        const [, rows] = ea.merge()
        const { totalKey, rankKey, data } = totalRankCells(rows as (string | object)[][])
        const empty = data.find((r) => Number(r['流水号']) === 200102)!
        expect((empty[totalKey] as { f?: string }).f).toMatch(/SUM/i)
        expect(rankKey && (empty[rankKey] as { f?: string }).f).toMatch(/RANK/i)
    })

    it('skipAtOrBelow=0：合计≤0 不写总分/排名', async () => {
        const file = fileFromAoa('skip-total.xlsx', [
            {
                sheetName: 's',
                aoa: [
                    ['流水号', '语文', '数学'],
                    [1, 50, 50], // 合计 100
                    [2, '', ''], // 合计 0 → skip
                    [3, 0, 0], // 合计 0 → skip
                    [4, 10, 5], // 合计 15
                ],
            },
        ])

        const ea = new ExcelAdapter()
        await ea.read([file])
        ea.setPrimaryColumns(['流水号'])

        const subjects = ea.getSubjectColumns()
        for (const map of subjects.values()) {
            for (const col of map.keys()) {
                if (col === '语文') map.set(col, SUBJECT_ID.CHINESE)
                else if (col === '数学') map.set(col, SUBJECT_ID.MATH)
                else map.set(col, SUBJECT_ID.NONE)
            }
        }
        ea.setSubjectColumns(subjects)
        ea.addTotalScoreColumn({
            subjects: [SUBJECT_ID.CHINESE, SUBJECT_ID.MATH],
            rank: true,
            skipAtOrBelow: 0,
        })

        const [, rows] = ea.merge()
        const { totalKey, rankKey, data } = totalRankCells(rows as (string | object)[][])

        const r1 = data.find((r) => r['流水号'] === 1)!
        expect((r1[totalKey] as { f?: string }).f).toMatch(/SUM/i)
        expect(rankKey && (r1[rankKey!] as { f?: string }).f).toMatch(/RANK/i)

        for (const serial of [2, 3]) {
            const r = data.find((row) => row['流水号'] === serial)!
            expect(r[totalKey]).toBe('')
            expect(rankKey && r[rankKey!]).toBe('')
        }

        const r4 = data.find((r) => r['流水号'] === 4)!
        expect((r4[totalKey] as { f?: string }).f).toMatch(/SUM/i)
    })

    it('skipAtOrBelow=100：合计≤100 跳过，>100 保留', async () => {
        const file = fileFromAoa('skip-n100.xlsx', [
            {
                sheetName: 's',
                aoa: [
                    ['流水号', '语文', '数学'],
                    [1, 40, 40], // 80 ≤100 skip
                    [2, 60, 50], // 110 >100 keep
                    [3, 100, 0], // 100 ≤100 skip
                ],
            },
        ])

        const ea = new ExcelAdapter()
        await ea.read([file])
        ea.setPrimaryColumns(['流水号'])

        const subjects = ea.getSubjectColumns()
        for (const map of subjects.values()) {
            for (const col of map.keys()) {
                if (col === '语文') map.set(col, SUBJECT_ID.CHINESE)
                else if (col === '数学') map.set(col, SUBJECT_ID.MATH)
                else map.set(col, SUBJECT_ID.NONE)
            }
        }
        ea.setSubjectColumns(subjects)
        ea.addTotalScoreColumn({
            subjects: [SUBJECT_ID.CHINESE, SUBJECT_ID.MATH],
            rank: true,
            skipAtOrBelow: 100,
        })

        const [, rows] = ea.merge()
        const { totalKey, data } = totalRankCells(rows as (string | object)[][])

        expect(data.find((r) => r['流水号'] === 1)![totalKey]).toBe('')
        expect((data.find((r) => r['流水号'] === 2)![totalKey] as { f?: string }).f).toMatch(/SUM/i)
        expect(data.find((r) => r['流水号'] === 3)![totalKey]).toBe('')
    })

    it('多总分列可分别配置 skipAtOrBelow', async () => {
        const file = fileFromAoa('skip-multi-col.xlsx', [
            {
                sheetName: 's',
                aoa: [
                    ['流水号', '语文', '数学', '英语'],
                    [1, 0, 0, 80], // 语数合计 0；语数英 80
                ],
            },
        ])

        const ea = new ExcelAdapter()
        await ea.read([file])
        ea.setPrimaryColumns(['流水号'])
        const subjects = ea.getSubjectColumns()
        for (const map of subjects.values()) {
            for (const col of map.keys()) {
                if (col === '语文') map.set(col, SUBJECT_ID.CHINESE)
                else if (col === '数学') map.set(col, SUBJECT_ID.MATH)
                else if (col === '英语') map.set(col, SUBJECT_ID.ENGLISH)
                else map.set(col, SUBJECT_ID.NONE)
            }
        }
        ea.setSubjectColumns(subjects)
        ea.addTotalScoreColumn({
            subjects: [SUBJECT_ID.CHINESE, SUBJECT_ID.MATH],
            rank: false,
            skipAtOrBelow: 0, // 跳过
        })
        ea.addTotalScoreColumn({
            subjects: [SUBJECT_ID.CHINESE, SUBJECT_ID.MATH, SUBJECT_ID.ENGLISH],
            rank: false,
            skipAtOrBelow: null, // 不跳过
        })

        const [, rows] = ea.merge()
        const header = rows[0] as string[]
        const dataRow = rows[1]
        // 语数总分应空；语数英总分应有 SUM
        const yuShuIdx = header.findIndex((h) => h === '语数总分')
        const allIdx = header.findIndex((h) => h === '语数英总分')
        expect(yuShuIdx).toBeGreaterThan(-1)
        expect(allIdx).toBeGreaterThan(-1)
        expect(dataRow[yuShuIdx]).toBe('')
        expect((dataRow[allIdx] as { f?: string }).f).toMatch(/SUM/i)
    })
})
