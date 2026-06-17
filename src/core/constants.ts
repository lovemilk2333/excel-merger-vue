export type ValuesOf<T> = T[keyof T]

export type ColumnName = string

export const COLUMN_TYPE = {
    NONE: 'none',
    NAME: 'name',
    ID: 'id',
    CLASS: 'class',
    SCORE: 'score',
} as const

export type ColumnType = ValuesOf<typeof COLUMN_TYPE>

export const COLUMN_TYPE_PRIMARY_KEY = Object.freeze(new Set([
    COLUMN_TYPE.NAME, COLUMN_TYPE.ID, COLUMN_TYPE.CLASS
])) satisfies Set<ColumnType>

export const COLUMN_TYPES = {
    [COLUMN_TYPE.NAME]: Object.freeze(["*名*", "name", "学生"]),
    [COLUMN_TYPE.ID]: Object.freeze(["*号*", "number", "id", "student", "no", "no.", "学号"]),
    [COLUMN_TYPE.CLASS]: Object.freeze(["*班*", "class", "班级"]),
    [COLUMN_TYPE.SCORE]: Object.freeze(["分数", "总分", "?分"]),
} satisfies Record<Exclude<ColumnType, 'none'>, readonly string[]>  // glob

export interface SUBJECT_ID_META_VALUE {
    order: number,
    display: string
    matchers: readonly string[] // glob
}

export const SUBJECT_ID = {
    NONE: 'none',
    CHINESE: 'chinese',
    MATH: 'math',
    ENGLISH: 'english',
    COMPUTER: 'computer',
    IT: 'it',
    PHYSICS: 'physics',
    CHEMISTRY: 'chemistry',
    BIOLOGY: 'biology',
    POLITICS: 'politics',
    HISTORY: 'history',
    GEOGRAPHY: 'geography',
    PE: 'pe',
    ART: 'art',
    MUSIC: 'music',
    MAJOR: 'major',
    SKILL: 'skill'
} as const

export type SubjectId = ValuesOf<typeof SUBJECT_ID>

export const SUBJECT_ID_META = {
    [SUBJECT_ID.CHINESE]: {
        order: 0,
        display: '语文',
        matchers: Object.freeze(['语文', 'chinese', 'yuwen'])
    },

    [SUBJECT_ID.MATH]: {
        order: 1,
        display: '数学',
        matchers: Object.freeze(['数学', 'math', 'maths', 'shuxue'])
    },

    [SUBJECT_ID.ENGLISH]: {
        order: 2,
        display: '英语',
        matchers: Object.freeze(['英语', 'english', 'yingyu'])
    },

    [SUBJECT_ID.COMPUTER]: {
        order: 3,
        display: '计算机',
        matchers: Object.freeze(['计算机', '计算机应用', 'computer', 'cs', 'jisuanji'])
    },

    [SUBJECT_ID.PHYSICS]: {
        order: 4,
        display: '物理',
        matchers: Object.freeze(['物理', 'physics', 'phy', 'wuli'])
    },

    [SUBJECT_ID.CHEMISTRY]: {
        order: 5,
        display: '化学',
        matchers: Object.freeze(['化学', 'chemistry', 'chem', 'huaxue'])
    },

    [SUBJECT_ID.BIOLOGY]: {
        order: 6,
        display: '生物',
        matchers: Object.freeze(['生物学', 'biology', 'bio', 'shengwu'])
    },

    [SUBJECT_ID.POLITICS]: {
        order: 7,
        display: '政治',
        matchers: Object.freeze([
            '道法',
            '道德与法治',
            '思想政治',
            'politics',
            'political science',
            'sizheng',
            'sixiangzhengzhi',
            'daofa',
            'daodeyufazhi'
        ])
    },

    [SUBJECT_ID.HISTORY]: {
        order: 7,
        display: '历史',
        matchers: Object.freeze(['历史', 'history', 'hist', 'lishi'])
    },

    [SUBJECT_ID.GEOGRAPHY]: {
        order: 8,
        display: '地理',
        matchers: Object.freeze(['地理', 'geography', 'geo', 'dili'])
    },

    [SUBJECT_ID.IT]: {
        order: 9,
        display: '信息技术',
        matchers: Object.freeze([
            '信息技术',
            '信息',
            'it',
            'information technology',
            'xinxijishu'
        ])
    },

    [SUBJECT_ID.PE]: {
        order: 10,
        display: '体育',
        matchers: Object.freeze([
            '体育',
            'p.e.',
            'pe',
            'physical education',
            'tiyu'
        ])
    },

    [SUBJECT_ID.ART]: {
        order: 11,
        display: '美术',
        matchers: Object.freeze([
            '艺术',
            'art',
            'arts',
            'meishu'
        ])
    },

    [SUBJECT_ID.MUSIC]: {
        order: 12,
        display: '音乐',
        matchers: Object.freeze(['音乐', 'music', 'yinyue'])
    },

    [SUBJECT_ID.MAJOR]: {
        order: 13,
        display: '专业课',
        matchers: Object.freeze([
            '专业',
            '专业课',
            '专业理论',
            'major',
            'zhuanye',
            'zhuanyeke'
        ])
    },

    [SUBJECT_ID.SKILL]: {
        order: 14,
        display: '技能',
        matchers: Object.freeze([
            '技术',  // 浙江省
            '技能',
            '专业技能',
            'skill',
            'skills',
            'jineng',
            'jishu'
        ])
    }
} satisfies Record<Exclude<SubjectId, 'none'>, SUBJECT_ID_META_VALUE>

export const VALID_SCORE_RANGE: readonly [number, number] = [-99999, 99999] as const

// /**
//  * fuck IEEE 754 !!!
//  * '93.5' -> '9350'
//  * [:-POINT_INDEX] + '.' + [-POINT_INDEX:]
//  */
// export const SUBJECT_SCORE_TIMES: number = 100
// export const SUBJECT_SCORE_POINT_INDEX: number = Math.log10(SUBJECT_SCORE_TIMES)


export const EXCEL_ADAPTER_CACHE_TYPE = {
    PRIMARY_COLUMNS: 'primary-columns',
    SUBJECT_COLUMNS: 'subject-columns',
    SUBJECT_IDS: 'subject-ids'
} as const

export type ExcelAdapterCacheTypes = ValuesOf<typeof EXCEL_ADAPTER_CACHE_TYPE>

export type ExcelSheetId = string  // `{filename: <filename>, sheetname: <sheetname>}`
export type ExcelSheetIdParsed = { filename: string, sheetname: string }
export type SubjectColumnsType = Map<ExcelSheetId, Map<ColumnName, SubjectId>>

export type ExcelAdapterCacheTypeValues = {
    [EXCEL_ADAPTER_CACHE_TYPE.PRIMARY_COLUMNS]: [ColumnName[], number[]],
    [EXCEL_ADAPTER_CACHE_TYPE.SUBJECT_COLUMNS]: SubjectColumnsType,
    [EXCEL_ADAPTER_CACHE_TYPE.SUBJECT_IDS]: Exclude<SubjectId, 'none'>[]
}

export type JSONSheetRow = {
    [key: string]: any
}
export type JSONSheetRows = JSONSheetRow[]

export type NormalizedSheet = { filename: string, sheetname: string, rows: JSONSheetRows }

export const TOTAL_SCORE_NAME = '总分'
export const RANK_SCORE_NAME = '排名'

export interface TotalScoreColumn {
    subjects: SubjectId[]
    rank: boolean  // 是否包含排名列
}

export type Score = number

export const LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
