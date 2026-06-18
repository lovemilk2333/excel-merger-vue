<template>
    <div class="subject-selector">
        <Card>
            <template #title>
                <div class="flex gap-4">
                    <div><span v-html="title"></span></div>
                    <div class="ml-auto flex justify-start">
                        <Button :label="canAddRow ? '添加科目-列对应关系' : '无可用列'" icon="pi pi-plus" size="small" outlined
                            :disabled="!canAddRow" @click="addRow" />
                    </div>
                </div>
                <Divider />
            </template>
            <template #content>
                <div class="subject-selector-content flex flex-col gap-4">
                    <div class="subject-selector-item p-x-2 flex gap-2 items-center" v-for="(row, index) in localRows"
                        :key="index">
                        <Select placeholder="选择一个列" :options="getFilteredColumns(row.column, row.columnFilterQuery)"
                            v-model="row.column" @change="handleRowChange" class="flex-1">
                            <template #header>
                                <div class="p-2 border-b border-surface-200 dark:border-surface-700">
                                    <div class="relative w-full">
                                        <i
                                            class="pi pi-search absolute left-3 top-1/2 -translate-y-1/2 text-surface-400" />
                                        <input type="text" v-model="row.columnFilterQuery" placeholder="输入列名进行过滤..."
                                            class="w-full pl-9 pr-3 py-1.5 text-sm bg-transparent border-none outline-none focus:outline-none focus:ring-0 shadow-none focus:shadow-none"
                                            @click.stop />
                                    </div>
                                </div>
                            </template>

                            <template #empty>
                                <div class="p-3 text-center text-gray-400 text-sm">
                                    没有找到匹配的列
                                </div>
                            </template>
                        </Select>

                        <span>作为</span>

                        <Select placeholder="选择一个科目" :options="getFilteredSubjects(row.subjectFilterQuery)"
                            v-model="row.subject" @change="handleRowChange"
                            :option-label="(subjectId: SubjectId) => SUBJECT_ID_META[subjectId as keyof typeof SUBJECT_ID_META]?.display"
                            class="flex-1">
                            <template #header>
                                <div class="p-2 border-b border-surface-200 dark:border-surface-700">
                                    <div class="relative w-full">
                                        <i
                                            class="pi pi-search absolute left-3 top-1/2 -translate-y-1/2 text-surface-400" />
                                        <input type="text" v-model="row.subjectFilterQuery" placeholder="输入科目或拼音进行过滤..."
                                            class="w-full pl-9 pr-3 py-1.5 text-sm bg-transparent border-none outline-none focus:outline-none focus:ring-0 shadow-none focus:shadow-none"
                                            @click.stop />
                                    </div>
                                </div>
                            </template>

                            <template #empty>
                                <div class="p-3 text-center text-gray-400 text-sm">
                                    没有找到匹配的科目
                                </div>
                            </template>
                        </Select>

                        <span>成绩</span>

                        <Button icon="pi pi-trash" severity="danger" variant="text" rounded @click="removeRow(index)"
                            aria-label="删除科目-列对应关系" />
                    </div>

                    <div v-if="localRows.length === 0" class="text-center text-gray-4 text-sm p-y-2">
                        请添加科目-列对应关系
                    </div>
                </div>
            </template>
        </Card>
    </div>
</template>

<script setup lang="ts">
import { ref, computed, watch, onMounted } from 'vue';
import { Card, Select, Button, Divider } from 'primevue';
import { SUBJECT_ID, SUBJECT_ID_META, type ColumnName, type SubjectId } from '../core/constants';
import { normalizeColumnName } from '../core/utils';

interface MappingRow {
    column: ColumnName | null;
    subject: SubjectId | null;
    columnFilterQuery: string;
    subjectFilterQuery: string;
}

const props = defineProps({
    title: { type: String, required: true },
    modelValue: { type: Map<ColumnName, SubjectId>, required: true }
})

const emit = defineEmits<{
    (e: 'update:modelValue'): void
}>()

const localRows = ref<MappingRow[]>([]);

let isSelfUpdating = false;

const allColumns = computed(() => Array.from(props.modelValue.keys()));

const initLocalRows = () => {
    localRows.value = Array.from(props.modelValue.entries())
        .filter(([_, subject]) => subject !== SUBJECT_ID.NONE && typeof subject !== 'undefined')
        .map(([column, subject]) => ({
            column,
            subject,
            columnFilterQuery: '',
            subjectFilterQuery: ''
        }));
};

onMounted(() => {
    initLocalRows();
});

watch(() => props.modelValue, () => {
    if (isSelfUpdating) {
        isSelfUpdating = false;
        return;
    }
    initLocalRows();
}, { deep: true });

const usedColumns = computed(() => {
    return localRows.value.map((val) => val.column).filter((col): col is ColumnName => col !== null);
});

const unselectedColumns = computed(() => {
    return allColumns.value.filter(val => !usedColumns.value.includes(val));
});

const baseSubjects = Object.keys(SUBJECT_ID_META).filter((val) => val !== SUBJECT_ID.NONE) as SubjectId[];

const getFilteredColumns = (currentRowColumn: ColumnName | null, filterValue: string): ColumnName[] => {
    const available = currentRowColumn ? [...unselectedColumns.value, currentRowColumn] : unselectedColumns.value;

    const query = normalizeColumnName(filterValue || "").replaceAll(" ", "");
    if (!query) return available;

    return available.filter((colName) => {
        const normalizedCol = normalizeColumnName(colName).replaceAll(" ", "");
        return normalizedCol.includes(query);
    });
};


const getFilteredSubjects = (filterValue: string): SubjectId[] => {
    const query = normalizeColumnName(filterValue || "").replaceAll(" ", "");

    if (!query) return baseSubjects;

    return baseSubjects.filter((subjectId) => {
        const meta = SUBJECT_ID_META[subjectId as keyof typeof SUBJECT_ID_META];
        if (!meta) return false;

        return meta.matchers.some((val) => {
            const normalizedMatcher = normalizeColumnName(val).replaceAll(" ", "");
            return normalizedMatcher.includes(query) || query.includes(normalizedMatcher);
        });
    });
};

const syncToParent = () => {
    allColumns.value.forEach(col => props.modelValue.set(col, SUBJECT_ID.NONE));

    localRows.value.forEach(row => {
        if (row.column && row.subject) {
            props.modelValue.set(row.column, row.subject);
        }
    });

    isSelfUpdating = true;
    emit('update:modelValue');
};

const addRow = () => {
    if (!canAddRow.value) return;

    const nextAvailableColumn = unselectedColumns.value[0] || null;

    localRows.value.push({
        column: nextAvailableColumn,
        subject: null,
        columnFilterQuery: '',
        subjectFilterQuery: ''
    });

    if (nextAvailableColumn) {
        syncToParent();
    }
};

const removeRow = (index: number) => {
    localRows.value.splice(index, 1);
    syncToParent();
};

const handleRowChange = () => {
    syncToParent();
};

const canAddRow = computed(() => unselectedColumns.value.length > 0);
</script>
