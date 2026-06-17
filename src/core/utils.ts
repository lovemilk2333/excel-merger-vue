// import { SUBJECT_SCORE_POINT_INDEX } from "./constants";

import { COLUMN_TYPE, COLUMN_TYPES, LETTERS, SUBJECT_ID, SUBJECT_ID_META, type ColumnType, type SubjectId } from "./constants";
import { isMatch } from 'picomatch';

// export function score2display(score: number): string {
//     if (Number.isNaN(score) || Number.isFinite(score)) {
//         return ''
//     }

//     const strScore = score.toString()
//     return strScore.slice(0, strScore.length - SUBJECT_SCORE_POINT_INDEX) + "." + strScore.slice(strScore.length - SUBJECT_SCORE_POINT_INDEX)
// }

export function dropDecimal(num: number): number {
    return num | 0
}

export const formatSize = (bytes: number) => {
    const k = 1024;
    const dm = 3;
    const sizes = ['B', 'KiB', 'MiB', 'GiB', 'TiB'];

    if (bytes === 0) {
        return `0 ${sizes[0]}`;
    }

    const i = Math.floor(Math.log(bytes) / Math.log(k));
    const formattedSize = parseFloat((bytes / Math.pow(k, i)).toFixed(dm));

    return `${formattedSize} ${sizes[i]}`;
};

export function normalizeColumnName(str: string): string {
    return str.toLowerCase().trim()
}

export function getColumnType(columnName: string): ColumnType {
    columnName = normalizeColumnName(columnName)

    for (const [type, globs] of Object.entries(COLUMN_TYPES)) {
        if (isMatch(columnName, globs as string[])) {
            return type as ColumnType
        }
    }

    return COLUMN_TYPE.NONE
}

export function getSubjectId(name: string, contains: boolean = false) {
    for (const [subjectId, { matchers }] of Object.entries(SUBJECT_ID_META)) {
        if (isMatch(name, (contains ? matchers.map((val) => `*${val}*`) : matchers) as string[])) {
            return subjectId as SubjectId
        }
    }

    return SUBJECT_ID.NONE
}

/**
 * @param index the column index, the first (left) one is `0`
 * @returns 0 -> `A`, 26 -> `AA`, 52 -> `BA`
 */
export function getExcelColumnName(index: number): string {
    if (!(index >= 0) || index % 1 !== 0) {  // NOTE: `!(index >= 0)` is required when index is `NaN`
        throw new Error('`index` must be a non-negative integer')
    }

    const length = LETTERS.length

    let result = ''
    while (index >= 0) {
        const letterIndex = index % length

        result = LETTERS[letterIndex] + result
        index = dropDecimal(index / length) - 1
    }

    return result
}
