import { read, type WorkSheet, utils, type CellObject, writeFile } from 'xlsx'
import { COLUMN_TYPE_PRIMARY_KEY, type ExcelAdapterCacheTypeValues, type ColumnName, type ExcelSheetId, type ExcelAdapterCacheTypes, EXCEL_ADAPTER_CACHE_TYPE, type NormalizedSheet, type SubjectId, SUBJECT_ID_META, SUBJECT_ID, COLUMN_TYPE, type TotalScoreColumn, type ExcelSheetIdParsed, type SubjectColumnsType, type JSONSheetRow, type Score, VALID_SCORE_RANGE, TOTAL_SCORE_NAME, RANK_SCORE_NAME } from './constants'
import { getColumnType, getExcelColumnName, getSubjectId, normalizeColumnName } from './utils'


export class ExcelAdapter {
    private sheets: Map<ExcelSheetId, NormalizedSheet> = new Map()
    private cache: Map<ExcelAdapterCacheTypes, ExcelAdapterCacheTypeValues> = new Map()
    private primaryColumns: ColumnName[] = []
    private subjectColumns: SubjectColumnsType = new Map()
    private totalScoreColumns: TotalScoreColumn[] = []

    constructor() {
        this.init()
    }

    init() {
        this.primaryColumns.length = 0
        this.subjectColumns.clear()
        this.totalScoreColumns.length = 0
        this.initCache()
    }

    initCache() {
        this.sheets.clear()
        this.cache.clear()
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

    private findLastMergedRow(sheet: WorkSheet) {
        const merges = sheet['!merges'];
        if (!merges || merges.length === 0) return -1;

        // 1. 过滤出所有“跨列合并”（水平合并）的区域
        // 并且只关心行范围
        const mergedRowRanges = merges
            .filter(range => range.s.c !== range.e.c) // 必须跨列
            .map(range => ({ start: range.s.r, end: range.e.r }))
            .sort((a, b) => a.start - b.start); // 按起始行升序排序

        // 2. 验证是否从第 0 行（第一行）开始连续
        let currentEndRow = -1;

        // 如果没有从第 0 行开始的合并，直接返回 -1
        if (mergedRowRanges.length === 0 || mergedRowRanges[0].start !== 0) {
            return -1;
        }

        currentEndRow = mergedRowRanges[0].end;

        // 3. 贪心检查后续的合并块是否与当前范围连续
        for (let i = 1; i < mergedRowRanges.length; i++) {
            const nextChunk = mergedRowRanges[i];

            // 如果下一个合并块的起始行，紧跟在当前结束行的下一行（或在当前范围内，即重叠）
            if (nextChunk.start <= currentEndRow + 1) {
                // 更新当前连续合并能到达的最远行
                currentEndRow = Math.max(currentEndRow, nextChunk.end);
            } else {
                // 出现断层，直接中断循环
                break;
            }
        }

        return currentEndRow;
    }

    private normalizeSheet(filename: NormalizedSheet['filename'], sheetname: NormalizedSheet['sheetname'], sheet: WorkSheet): NormalizedSheet {
        const range = this.findLastMergedRow(sheet) + 1
        return {
            filename,
            sheetname,
            rows: utils.sheet_to_json(sheet, { range })
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
                this.sheets.set(this.sheetId(file.name, sheetname), this.normalizeSheet(file.name, sheetname, workbook.Sheets[sheetname]))
            }
        }

        return errors
    }

    /**
     * merge all data with the config
     */
    merge(): [WorkSheet, (string | CellObject)[][], { row: JSONSheetRow, error: Error }[]] {
        if (!this.sheets.size) {
            throw new Error('at least one excel sheet is required')
        }
        if (!this.getPrimaryColumns().length) {
            throw new Error('at least one primary column is required')
        }

        const errors: { row: JSONSheetRow, error: Error }[] = []
        const students = new Map<string, Map<Exclude<SubjectId, 'none'>, Score>>()
        let studentCount = 0

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
                    studentCount += 1
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
        const RANK_END_ROW = RANK_START_ROW + studentCount - 1
        const totalScoreColumnsStart = this.primaryColumns.length + subjectColumns.length
        let totalScoreCurrentIndex = totalScoreColumnsStart
        const totalScoreColumnsFormulas: (Partial<CellObject> | null)[] = totalScoreColumnsMeta.flatMap(({ subjects, rank }) => {
            const subjectColumnIndex: number[] = subjects.map(subjectId => subjectColumnsIndexMap[subjectId as Partial<Exclude<SubjectId, 'none'>>]).filter(val => typeof val === 'number')
            const subjectColumnExcelName: string[] = subjectColumnIndex.map(getExcelColumnName)

            const totalColumnName = getExcelColumnName(totalScoreCurrentIndex)
            // const rankColumnName = getExcelColumnName(totalScoreCurrentIndex + 1)

            totalScoreCurrentIndex += rank ? 2 : 1
            return [
                { t: 'n', f: `SUM(${subjectColumnExcelName.map(col => `${col}{ROWNUMBER}`).join(',')})` },
                rank ? { t: 'n', f: `RANK(${totalColumnName}{ROWNUMBER}, \$${totalColumnName}\$${RANK_START_ROW}:\$${totalColumnName}\$${RANK_END_ROW})` } : null
            ]
        })


        const rowColumns: (string | CellObject)[][] = [
            header,
        ]

        for (const [studentId, scores] of students.entries()) {
            const totalScoreColumnsStart = this.primaryColumns.length + subjectColumns.length
            const rowNumber = rowColumns.length + 1  // Excel Row Number (based 1)
            const row = new Array(totalScoreColumnsStart).fill('')

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

            row.push(
                ...totalScoreColumnsFormulas
                    .filter((val): val is Partial<CellObject> => val !== null)
                    .map(val => ({
                        ...val,
                        f: val.f?.replaceAll('{ROWNUMBER}', `${rowNumber}`) || ''
                    }))
            )

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
