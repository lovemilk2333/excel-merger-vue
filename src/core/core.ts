import { read, type WorkSheet, utils, type CellObject, writeFile } from 'xlsx'
import { COLUMN_TYPE_PRIMARY_KEY, type ExcelAdapterCacheTypeValues, type ColumnName, type ExcelSheetId, type ExcelAdapterCacheTypes, EXCEL_ADAPTER_CACHE_TYPE, type NormalizedSheet, type SubjectId, SUBJECT_ID_META, SUBJECT_ID, COLUMN_TYPE, type TotalScoreColumn, type ExcelSheetIdParsed, type SubjectColumnsType, type JSONSheetRow, type Score, VALID_SCORE_RANGE, TOTAL_SCORE_NAME, RANK_SCORE_NAME } from './constants'
import { getColumnType, getExcelColumnName, getSubjectId, normalizeColumnName } from './utils'
import { normalizeWorksheet, type SheetNormalizeOptions } from './sheet-normalize'

export type { SheetNormalizeOptions }

export class ExcelAdapter {
    private sheets: Map<ExcelSheetId, NormalizedSheet> = new Map()
    private cache: Map<ExcelAdapterCacheTypes, ExcelAdapterCacheTypeValues> = new Map()
    private primaryColumns: ColumnName[] = []
    private subjectColumns: SubjectColumnsType = new Map()
    private totalScoreColumns: TotalScoreColumn[] = []
    /** per-sheet normalize options (headerRows / dataStartRow); applied on next read or reNormalize */
    private sheetNormalizeOptions: Map<ExcelSheetId, SheetNormalizeOptions> = new Map()
    private defaultNormalizeOptions: SheetNormalizeOptions = {}
    /** raw worksheets retained so options can re-normalize without re-reading files */
    private rawSheets: Map<ExcelSheetId, WorkSheet> = new Map()
    /**
     * Fixed 源：导出行顺序与名单以此为准。
     * - `null`：不固定，按各表首次出现顺序
     * - 文件名：该 Excel 下全部 sheet（按读入顺序）决定行序
     * - sheetId（`sheetId(filename, sheetname)`）：仅该 sheet
     */
    private fixedSource: string | null = null

    constructor() {
        this.init()
    }

    init() {
        this.primaryColumns.length = 0
        this.subjectColumns.clear()
        this.totalScoreColumns.length = 0
        this.sheetNormalizeOptions.clear()
        this.defaultNormalizeOptions = {}
        this.fixedSource = null
        this.initCache()
    }

    initCache() {
        this.sheets.clear()
        this.rawSheets.clear()
        this.cache.clear()
    }

    /**
     * 指定 fixed 源（文件名或 sheetId）。导出时：
     * - 行顺序 = fixed 中学生出现顺序（同主键首次出现为准）
     * - 其他表缺成绩时对应科目留空，不改变顺序
     * - 不在 fixed 中的学生不会出现在结果中
     */
    setFixedSource(source: string | null) {
        this.fixedSource = source
    }

    getFixedSource(): string | null {
        return this.fixedSource
    }

    clearFixedSource() {
        this.fixedSource = null
    }

    private isFixedSheet(sheetId: ExcelSheetId, sheet: NormalizedSheet): boolean {
        if (!this.fixedSource) return false
        if (sheetId === this.fixedSource) return true
        if (sheet.filename === this.fixedSource) return true
        return false
    }

    /** 将科目成绩折合为可加总分的数值（空/非数字按 0） */
    private scoreAsNumber(score: Score | string | undefined): number {
        if (typeof score === 'number' && !Number.isNaN(score)) return score
        if (score === '' || score == null) return 0
        const n = Number(score)
        return Number.isFinite(n) ? n : 0
    }

    private sumSubjectsForTotal(
        scores: Map<Exclude<SubjectId, 'none'>, Score | string>,
        subjects: SubjectId[]
    ): number {
        let sum = 0
        for (const sid of subjects) {
            if (sid === SUBJECT_ID.NONE) continue
            sum += this.scoreAsNumber(scores.get(sid as Exclude<SubjectId, 'none'>))
        }
        return sum
    }

    /**
     * 默认表头/数据行配置（0-based）。未对单个 sheet 单独配置时使用。
     * 仅影响之后的 read / reNormalizeSheets。
     */
    setDefaultNormalizeOptions(options: SheetNormalizeOptions) {
        this.defaultNormalizeOptions = { ...options }
    }

    getDefaultNormalizeOptions(): SheetNormalizeOptions {
        return { ...this.defaultNormalizeOptions }
    }

    /**
     * 为指定 sheet 设置表头行 / 数据起始行（0-based）。
     * @param sheetId `sheetId(filename, sheetname)` 的返回值；也可用 filename+sheetname
     */
    setSheetNormalizeOptions(sheetId: ExcelSheetId, options: SheetNormalizeOptions) {
        this.sheetNormalizeOptions.set(sheetId, { ...options })
        if (this.rawSheets.has(sheetId)) {
            this.reNormalizeSheet(sheetId)
        }
    }

    getSheetNormalizeOptions(sheetId: ExcelSheetId): SheetNormalizeOptions | undefined {
        const opt = this.sheetNormalizeOptions.get(sheetId)
        return opt ? { ...opt } : undefined
    }

    clearSheetNormalizeOptions(sheetId?: ExcelSheetId) {
        if (typeof sheetId === 'undefined') {
            this.sheetNormalizeOptions.clear()
            this.defaultNormalizeOptions = {}
        } else {
            this.sheetNormalizeOptions.delete(sheetId)
        }
    }

    getSheetIds(): ExcelSheetId[] {
        return [...this.sheets.keys()]
    }

    getNormalizedSheet(sheetId: ExcelSheetId): NormalizedSheet | undefined {
        return this.sheets.get(sheetId)
    }

    /**
     * 在已 read 的前提下，用当前 options 重新规范化全部 sheet（会清 subject/primary 缓存语义相关 cache）。
     */
    reNormalizeSheets() {
        for (const id of this.rawSheets.keys()) {
            this.reNormalizeSheet(id)
        }
        this.primaryColumns.length = 0
        this.subjectColumns.clear()
        this.totalScoreColumns.length = 0
        this.cache.clear()
    }

    private resolveNormalizeOptions(sheetId: ExcelSheetId): SheetNormalizeOptions {
        return {
            ...this.defaultNormalizeOptions,
            ...this.sheetNormalizeOptions.get(sheetId),
        }
    }

    private reNormalizeSheet(sheetId: ExcelSheetId) {
        const raw = this.rawSheets.get(sheetId)
        const existing = this.sheets.get(sheetId)
        if (!raw || !existing) return

        const normalized = normalizeWorksheet(raw, this.resolveNormalizeOptions(sheetId))
        this.sheets.set(sheetId, {
            filename: existing.filename,
            sheetname: existing.sheetname,
            rows: normalized.rows,
        })
    }

    sheetId(filename: string, sheetname: string) {
        return JSON.stringify({ filename, sheetname } as ExcelSheetIdParsed)
    }

    studentId(row: JSONSheetRow) {
        const primaryKeyPairs: JSONSheetRow = {}

        for (const col of Object.values(this.primaryColumns)) {
            const val = row[col];
            if (typeof val === 'undefined') {
                throw new Error(`cannot get student id for row: ${JSON.stringify(row)}`)
            }

            primaryKeyPairs[col] = val
        }

        return JSON.stringify(primaryKeyPairs)
    }

    parseId(id: string): { [key: string]: any } | ExcelSheetIdParsed {
        return JSON.parse(id)
    }

    /**
     * @returns [total score column name, rank column name]
     * @example ['语数总分', '语数排名']
     */
    getTotalScoreName(totalScoreColumn: TotalScoreColumn): [string, string] {
        const subjectNames = totalScoreColumn.subjects.map(id => SUBJECT_ID_META[id as keyof typeof SUBJECT_ID_META]).sort((a, b) => a.order - b.order).map(meta => meta.display.slice(0, 1))

        return [subjectNames.join('') + TOTAL_SCORE_NAME, subjectNames.join('') + RANK_SCORE_NAME]
    }

    private normalizeSheet(
        filename: NormalizedSheet['filename'],
        sheetname: NormalizedSheet['sheetname'],
        sheet: WorkSheet,
        sheetId: ExcelSheetId
    ): NormalizedSheet {
        const { rows } = normalizeWorksheet(sheet, this.resolveNormalizeOptions(sheetId))
        return {
            filename,
            sheetname,
            rows,
        }
    }

    private setCache<T extends ExcelAdapterCacheTypes, R = ExcelAdapterCacheTypeValues[T]>(type: T): (val: R) => R {
        return (val: R) => {
            this.cache.set(type, structuredClone(val) as any)
            return val
        }
    }

    private getCache<T extends ExcelAdapterCacheTypes, R = ExcelAdapterCacheTypeValues[T]>(type: T): R | undefined {
        return this.cache.get(type) as R | undefined
    }

    private rmCache<T extends ExcelAdapterCacheTypes>(type: T): boolean {
        return this.cache.delete(type)
    }

    /**
     * @returns [shared columns (column which contains in all sheets), prefer columns (column which most possible to be the primary key)]
     */
    getSharedColumns(): [ColumnName[], number[]] {
        const cached = this.getCache(EXCEL_ADAPTER_CACHE_TYPE.PRIMARY_COLUMNS)
        if (!!cached) {
            return cached
        }

        const returnAndCache = this.setCache(EXCEL_ADAPTER_CACHE_TYPE.PRIMARY_COLUMNS)

        const sheets = this.sheets.values()
        const firstRows = sheets.next()

        if (firstRows.done || !firstRows.value || !firstRows.value?.rows?.length) {
            console.warn('no sheet or columns found')
            return returnAndCache([[], []])
        }

        let sharedColumns: ColumnName[] = Object.keys(firstRows.value.rows[0]);

        for (const sheet of sheets) {
            const rows = sheet.rows
            if (!rows || !rows.length) continue

            const currentColumns = Object.keys(rows[0])
            if (!currentColumns.length) {
                throw new Error('any sheet must have at least one column')
            }

            sharedColumns = [...sharedColumns].filter(col => currentColumns.includes(col))
        }

        const preferColumns = []
        for (let i = 0; i < sharedColumns.length; i++) {
            const columnType = getColumnType(sharedColumns[i])
            if (COLUMN_TYPE_PRIMARY_KEY.has(columnType as any)) {
                preferColumns.push(i)
            }
        }

        return returnAndCache([sharedColumns, preferColumns])
    }

    getPrimaryColumns(): ColumnName[] {
        return this.primaryColumns
    }

    setPrimaryColumns(columns: ColumnName[]) {
        const [sharedColumns, _] = this.getSharedColumns()
        const invalidColumns = columns.filter((columnName) => !sharedColumns.includes(columnName))
        if (invalidColumns.length) {
            throw new Error(`these primary columns are not found in shared columns: ${invalidColumns.join(', ')}`)
        }

        this.primaryColumns.length = 0
        this.primaryColumns.push(...columns)
    }

    getSubjectColumns(): SubjectColumnsType {
        const cached = this.getCache(EXCEL_ADAPTER_CACHE_TYPE.SUBJECT_COLUMNS)
        if (!!cached) {
            return cached
        }

        // const subjects = new Map() as SubjectColumnsType
        const subjects = this.subjectColumns
        subjects.clear()

        for (const [id, sheet] of this.sheets.entries()) {
            const subjectIdsMap = new Map() as NonNullable<ReturnType<SubjectColumnsType['get']>>;
            if (!sheet || !sheet.rows || !sheet.rows.length) {
                subjects.set(id, subjectIdsMap)
                continue
            }

            const rows = sheet.rows

            let subjectColumnFound = false
            // 查找单个文件中有多个 Subject Column 的 Column
            for (const columnName of Object.keys(rows[0])) {
                const subjectId = getSubjectId(normalizeColumnName(columnName))
                subjectColumnFound = subjectColumnFound || subjectId != SUBJECT_ID.NONE

                subjectIdsMap.set(columnName, subjectId)
            }

            // 如果没有匹配, 视为单文件, 从文件名/表名获取科目
            if (!subjectColumnFound) {
                const subjectIdSheetname = getSubjectId(normalizeColumnName(sheet.sheetname), true)
                const subjectIdFilename = getSubjectId(normalizeColumnName(sheet.filename), true)

                const subjectId = subjectIdFilename != SUBJECT_ID.NONE ? subjectIdFilename : subjectIdSheetname

                // 还是找不到, 那就算了 (xd
                if (subjectId == SUBJECT_ID.NONE) {
                    subjects.set(id, subjectIdsMap)
                    continue
                }

                // 如果匹配到了, 就找 总分/分数 的 Column
                for (const columnName of Object.keys(rows[0])) {
                    const columnType = getColumnType(normalizeColumnName(columnName))

                    // 现在 `subjectIdsMap` 已经全部为 `SUBJECT_ID.NONE` 了, 只更新找到的 subjectId 就行
                    if (columnType == COLUMN_TYPE.SCORE) {
                        subjectIdsMap.set(columnName, subjectId)
                    }
                }
            }

            subjects.set(id, subjectIdsMap)
        }

        return this.setCache(EXCEL_ADAPTER_CACHE_TYPE.SUBJECT_COLUMNS)(subjects)
    }

    private mergeSubjectMaps(parent: NonNullable<ReturnType<SubjectColumnsType['get']>>, source: NonNullable<ReturnType<SubjectColumnsType['get']>>) {
        for (const [columnName, subjectId] of source.entries()) {
            parent.set(columnName, subjectId);
        }

        this.rmCache(EXCEL_ADAPTER_CACHE_TYPE.SUBJECT_IDS)
    }

    setSubjectColumnsSheet(sheetId: ExcelSheetId, subjectMaps: NonNullable<ReturnType<SubjectColumnsType['get']>>) {
        const parent = this.subjectColumns.get(sheetId);
        if (!parent) {
            throw new Error('no such sheet id: ' + sheetId);
        }

        this.mergeSubjectMaps(parent, subjectMaps);
    }

    setSubjectColumns(subjects: SubjectColumnsType) {
        for (const [id, subjectMaps] of subjects.entries()) {
            if (!this.subjectColumns.has(id)) {
                // this.subjectColumns.set(id, new Map() as NonNullable<ReturnType<SubjectColumnsType['get']>>);
                continue
            }

            const parent = this.subjectColumns.get(id)!;
            this.mergeSubjectMaps(parent, subjectMaps);
        }
    }

    addTotalScoreColumn(meta: TotalScoreColumn) {
        const subjects = JSON.stringify(meta.subjects)
        const presentMeta = this.totalScoreColumns.find((val) => JSON.stringify(val.subjects) === subjects)

        if (presentMeta) {
            Object.assign(presentMeta, meta)
        } else {
            this.totalScoreColumns.push(meta)
        }
    }

    clearTotalScoreColumn() {
        this.totalScoreColumns.length = 0
    }

    getSubjectIds() {
        const cached = this.getCache(EXCEL_ADAPTER_CACHE_TYPE.SUBJECT_IDS)
        if (!!cached) {
            return cached
        }

        const subjectIds: Exclude<SubjectId, 'none'>[] = []
        for (const maps of this.subjectColumns.values()) {
            for (const subjectId of maps.values()) {
                if (subjectId === SUBJECT_ID.NONE || subjectIds.includes(subjectId)) continue

                subjectIds.push(subjectId)
            }
        }

        return this.setCache(EXCEL_ADAPTER_CACHE_TYPE.SUBJECT_IDS)(subjectIds)
    }

    /**
     * clear cached excel sheets and read excel sheets
     * @returns the files that cannot load to excel
     */
    async read(files: File[]): Promise<{ file: File, error: Error }[]> {
        this.init()

        const errors = []

        for (const file of files) {
            let workbook
            try {
                workbook = read(await file.arrayBuffer())
            } catch (e) {
                console.error('cannot read excel', file.name, e)
                errors.push({ file, error: e as Error })
                continue
            }

            for (const sheetname of workbook.SheetNames) {
                const id = this.sheetId(file.name, sheetname)
                const worksheet = workbook.Sheets[sheetname]
                this.rawSheets.set(id, worksheet)
                this.sheets.set(id, this.normalizeSheet(file.name, sheetname, worksheet, id))
            }
        }

        return errors
    }

    /**
     * 按 fixed 源收集学生主键顺序；无 fixed 时返回 null（调用方用 Map 插入序）。
     */
    private collectFixedStudentOrder(errors: { row: JSONSheetRow, error: Error }[]): string[] | null {
        if (!this.fixedSource) return null

        const order: string[] = []
        const seen = new Set<string>()
        let matchedSheet = false

        for (const [id, sheet] of this.sheets.entries()) {
            if (!this.isFixedSheet(id, sheet)) continue
            matchedSheet = true

            for (const row of sheet.rows) {
                let studentId: string
                try {
                    studentId = this.studentId(row)
                } catch (e) {
                    errors.push({ row, error: e as Error })
                    continue
                }
                if (seen.has(studentId)) continue
                seen.add(studentId)
                order.push(studentId)
            }
        }

        if (!matchedSheet) {
            throw new Error(
                `fixed source not found among loaded sheets: ${JSON.stringify(this.fixedSource)}`
            )
        }

        return order
    }

    /**
     * merge all data with the config
     */
    merge(): [WorkSheet, (string | number | CellObject)[][], { row: JSONSheetRow, error: Error }[]] {
        if (!this.sheets.size) {
            throw new Error('at least one excel sheet is required')
        }
        if (!this.getPrimaryColumns().length) {
            throw new Error('at least one primary column is required')
        }

        const errors: { row: JSONSheetRow, error: Error }[] = []
        const students = new Map<string, Map<Exclude<SubjectId, 'none'>, Score>>()

        for (const [id, sheet] of this.sheets.entries()) {
            for (const row of sheet.rows) {
                let studentId
                try {
                    studentId = this.studentId(row)
                } catch (e) {
                    errors.push({ row, error: e as Error })
                    continue
                }

                if (!students.has(studentId)) {
                    students.set(studentId, new Map())
                }

                const studentData = students.get(studentId)!
                const subjectMaps = this.subjectColumns.get(id)
                if (!subjectMaps) {
                    continue
                }

                for (const [columnName, subjectId] of subjectMaps.entries()) {
                    if (subjectId === SUBJECT_ID.NONE) continue

                    const score = row[columnName]

                    if (typeof score === 'number') {
                        if (!(VALID_SCORE_RANGE[0] <= score && score <= VALID_SCORE_RANGE[1])) {
                            // not in valid range
                            continue
                        }
                    } else if (score === '') {
                        // empty string is valid as zero in excel function `SUM()`
                    } else {
                        // not a valid score
                        continue
                    }

                    studentData.set(subjectId, score)
                }
            }
        }

        const fixedOrder = this.collectFixedStudentOrder(errors)
        // fixed：仅输出 fixed 名单且保持其顺序；缺成绩科目保持为空
        // 无 fixed：保持各表首次发现顺序
        const orderedStudentIds: string[] = fixedOrder
            ? fixedOrder
            : [...students.keys()]

        const studentCount = orderedStudentIds.length

        // subjectIds: {chinese, math, ...}
        // students: {'{"姓名": "aaa", "班级": "0000"}': Map => {chinese: 50, math: 30, ...}}
        const totalScoreColumnsMeta = this.totalScoreColumns.map(val => ({ ...val, names: this.getTotalScoreName(val) }))

        const subjectIds = this.getSubjectIds()
        const subjectColumnsMeta = subjectIds.map((id) => ({ ...SUBJECT_ID_META[id as keyof typeof SUBJECT_ID_META], id })).sort((a, b) => a.order - b.order)
        const subjectColumns = subjectColumnsMeta.map(meta => meta.display)
        const subjectColumnsIndexMap: Partial<Record<Exclude<SubjectId, 'none'>, number>> = {}
        for (const [index, val] of Object.entries(subjectColumnsMeta)) {
            subjectColumnsIndexMap[val.id] = this.primaryColumns.length + parseInt(index)
        }
        const totalScoreColumns = totalScoreColumnsMeta.flatMap(val => {
            const [scoreName, rankName] = val.names
            return val.rank ? [scoreName, rankName] : [scoreName];
        })

        const header = [
            ...this.primaryColumns,
            ...subjectColumns,
            ...totalScoreColumns
        ]

        const RANK_START_ROW = 2
        const RANK_END_ROW = studentCount > 0 ? RANK_START_ROW + studentCount - 1 : RANK_START_ROW
        const totalScoreColumnsStart = this.primaryColumns.length + subjectColumns.length
        let totalScoreCurrentIndex = totalScoreColumnsStart
        // 每个总分列一组模板；跳过条件按「列配置 skipAtOrBelow」+ 学生合计判定
        const totalScoreColumnSpecs = totalScoreColumnsMeta.map(({ subjects, rank, skipAtOrBelow }) => {
            const subjectColumnIndex: number[] = subjects
                .map(subjectId => subjectColumnsIndexMap[subjectId as Partial<Exclude<SubjectId, 'none'>>])
                .filter((val): val is number => typeof val === 'number')
            const subjectColumnExcelName: string[] = subjectColumnIndex.map(getExcelColumnName)

            const totalColumnName = getExcelColumnName(totalScoreCurrentIndex)
            totalScoreCurrentIndex += rank ? 2 : 1

            return {
                subjects,
                rank,
                skipAtOrBelow: typeof skipAtOrBelow === 'number' && !Number.isNaN(skipAtOrBelow)
                    ? skipAtOrBelow
                    : null as number | null,
                totalFormula: `SUM(${subjectColumnExcelName.map(col => `${col}{ROWNUMBER}`).join(',')})`,
                rankFormula: rank
                    ? `RANK(${totalColumnName}{ROWNUMBER}, \$${totalColumnName}\$${RANK_START_ROW}:\$${totalColumnName}\$${RANK_END_ROW})`
                    : null as string | null,
            }
        })

        const rowColumns: (string | number | CellObject)[][] = [
            header,
        ]

        for (const studentId of orderedStudentIds) {
            const scores = students.get(studentId) ?? new Map<Exclude<SubjectId, 'none'>, Score>()
            const rowNumber = rowColumns.length + 1  // Excel Row Number (based 1)
            const row: (string | number | CellObject)[] = new Array(totalScoreColumnsStart).fill('')

            for (const [primaryKey, primaryValue] of Object.entries(this.parseId(studentId))) {
                const index = header.indexOf(primaryKey)
                if (index <= -1) continue
                row[index] = primaryValue
            }

            for (const [subjectId, score] of scores.entries()) {
                const displayName = SUBJECT_ID_META[subjectId as keyof typeof SUBJECT_ID_META].display
                const index = header.indexOf(displayName)
                if (index <= -1) continue
                row[index] = score
            }

            for (const spec of totalScoreColumnSpecs) {
                const shouldSkip =
                    spec.skipAtOrBelow !== null &&
                    this.sumSubjectsForTotal(scores, spec.subjects) <= spec.skipAtOrBelow

                if (shouldSkip) {
                    row.push('')
                    if (spec.rank) row.push('')
                    continue
                }

                row.push({
                    t: 'n',
                    f: spec.totalFormula.replaceAll('{ROWNUMBER}', `${rowNumber}`),
                })
                if (spec.rank && spec.rankFormula) {
                    row.push({
                        t: 'n',
                        f: spec.rankFormula.replaceAll('{ROWNUMBER}', `${rowNumber}`),
                    })
                }
            }

            rowColumns.push(row)
        }

        const worksheet = utils.aoa_to_sheet(rowColumns)

        return [worksheet, rowColumns, errors]
    }

    private generateOutputFilename(): string | null {
        if (!this.sheets || this.sheets.size === 0) return null;

        const filenames: string[] = [];
        for (const sheet of this.sheets.values()) {
            if (sheet.filename) {
                // 【优化】在这里直接提取纯文件名，避免路径干扰（例如去掉 "examples/"）
                const parts = sheet.filename.split('/');
                filenames.push(parts[parts.length - 1]);
            }
        }

        if (filenames.length === 0) return null;

        // === 修改核心逻辑：从原有的首尾前缀比对，改为全量字符共有性检查 ===

        // 3. 挑选长度最短的一个文件名作为基准字符串（减少循环次数）
        const base = filenames.reduce((s1, s2) => s1.length < s2.length ? s1 : s2);

        let common = "";

        // 4. 遍历基准字符串的每一个字符，检查它是否在所有文件名中都存在
        for (const char of base) {
            const isCommon = filenames.every(name => name.includes(char));
            if (isCommon) {
                common += char;
            }
        }

        // 5. 【新增】移除因后缀名相同而被误提取的扩展名（如 .xls, .xlsx, .csv）
        // 否则 "得分表.xls" 和 "得分表.xls" 会被提取出 "得分表.xls"
        common = common.replace(/\.(xlsx?|csv)$/i, '');

        // =============================================================

        // 6. 去掉末尾可能残留的连字符或下划线
        // 1. 定义 NTFS 禁用的字符正则：\ / : * ? " < > |
        const ntfsInvalidChars = /[\\/:*?"<>|]/g;

        // 2. 将共有部分中的所有 NTFS 禁用字符替换为下划线 '_'
        let sanitized = common.replace(ntfsInvalidChars, '_');

        // 3. 处理末尾符号：
        // 去掉末尾的连字符、下划线、括号、空格或中文符号等多余杂质
        // 这里加入了对中文范围的兼容，防止误删末尾的中文字符
        const result = sanitized.replace(/[^a-zA-Z0-9\u4e00-\u9fa5]+$/, '');

        console.debug('generateOutputFilename', filenames, '->', result);
        return result.length > 0 ? result : null;
    }

    export(worksheet: WorkSheet) {
        const workbook = utils.book_new();
        utils.book_append_sheet(workbook, worksheet);
        const filename = this.generateOutputFilename()
        writeFile(workbook, filename ? filename + '.xlsx' : 'result.xlsx');
    }
}
