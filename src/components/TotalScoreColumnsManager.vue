<template>
    <div class="total-score-columns">
        <Card>
            <template #title>
                <div class="flex">
                    {{ title }}
                    <div class="ml-auto flex justify-start">
                        <Button label="新增" icon="pi pi-plus" size="small" outlined @click="addRow" />
                    </div>
                </div>
                <Divider />
            </template>
            <template #content>
                <div class="total-score-columns-content flex flex-col gap-4 items-stretch">
                    <template v-if="!!modelValue && !!modelValue.length">
                        <div class="total-score-columns-row flex gap-4 items-center justify-stretch"
                            v-for="(_, index) in modelValue" :key="index">
                            <MultiSelect name="subject" :options="getFilteredSubjects(index)"
                                :optionLabel="(id) => subjectIdNameMap[id]" placeholder="选择科目"
                                v-model="modelValue[index].subjects" class="flex-1">
                                <template #header>
                                    <div class="p-2 border-b border-surface-200 dark:border-surface-700">
                                        <div class="relative w-full">
                                            <i
                                                class="pi pi-search absolute left-3 top-1/2 -translate-y-1/2 text-surface-400" />
                                            <input type="text" v-model="rowFilters[index].filterQuery"
                                                placeholder="输入科目或拼音进行过滤..."
                                                class="w-full pl-9 pr-3 py-1.5 text-sm bg-transparent border-none outline-none focus:outline-none focus:ring-0 shadow-none focus:shadow-none"
                                                @click.stop />
                                        </div>
                                    </div>
                                </template>

                                <template #empty>
                                    <div class="p-3 text-center text-gray-400 text-sm">
                                        {{ subjectIds.length ? '没有找到匹配的科目' : '无可选科目, 请确认科目-列对应关系是否添加成功' }}
                                    </div>
                                </template>
                            </MultiSelect>

                            <div class="flex items-center gap-2 shrink-0">
                                <Checkbox v-model="modelValue[index].rank" :inputId="`rank-${index}`" name="rank"
                                    binary />
                                <label :for="`rank-${index}`" class="select-none text-sm">生成排名</label>
                            </div>

                            <Message severity="error" variant="outlined" v-if="errors.get(index)">{{ errors.get(index)
                                }}
                            </Message>
                            <Button icon="pi pi-trash" severity="danger" variant="text" rounded
                                @click="removeRow(index)" aria-label="删除总分/排名列" />
                        </div>
                    </template>
                    <div v-else class="text-center text-gray-4 text-sm p-y-2">
                        请添加总分/排名列
                    </div>
                </div>
            </template>
        </Card>
    </div>
</template>

<script setup lang="ts">
import { ref, watch } from 'vue';
import { Button, Card, Checkbox, Divider, MultiSelect, Message } from 'primevue';
import { SUBJECT_ID, SUBJECT_ID_META, type SubjectId, type TotalScoreColumn } from '../core/constants';

const props = withDefaults(defineProps<{
    title?: string
    modelValue: TotalScoreColumn[],
    subjectIds: Exclude<SubjectId, 'none'>[]
}>(), {
    title: '新增或删除总分/排名列'
})

const errors = ref<Map<number, string>>(new Map())
const subjectIdNameMap = Object.fromEntries(Object.values(SUBJECT_ID).map((subjectId) => {
    if (subjectId === SUBJECT_ID.NONE) return []

    const meta = SUBJECT_ID_META[subjectId as Exclude<SubjectId, 'none'>]
    return [subjectId, meta.display]
}))

interface RowFilter {
    filterQuery: string;
}
const rowFilters = ref<RowFilter[]>([]);

watch(() => props.modelValue, (newVal) => {
    if (!newVal) return;
    while (rowFilters.value.length < newVal.length) {
        rowFilters.value.push({ filterQuery: '' });
    }

    if (rowFilters.value.length > newVal.length) {
        rowFilters.value.length = newVal.length;
    }
    totalColumnsErrors()
}, { immediate: true, deep: true });


const getFilteredSubjects = (index: number) => {
    const currentSubjects = props.subjectIds;

    const filterState = rowFilters.value[index];
    if (!filterState) return currentSubjects;

    const query = filterState.filterQuery.trim().toLowerCase();
    if (!query) return currentSubjects;

    return currentSubjects.filter((id) => {
        const displayName = subjectIdNameMap[id]?.toLowerCase() || '';

        if (displayName.includes(query)) return true;

        const meta = SUBJECT_ID_META[id];
        if (meta && meta.matchers) {
            return meta.matchers.some(matcher => matcher.toLowerCase().includes(query));
        }

        return false;
    });
};

function addRow() {
    props.modelValue.push({
        subjects: [],
        rank: true
    })
    rowFilters.value.push({ filterQuery: '' })
}

function removeRow(index: number) {
    props.modelValue.splice(index, 1)
    rowFilters.value.splice(index, 1)
}

function totalColumnsErrors() {
    const subjects = new Set<string>();

    errors.value.clear()
    props.modelValue.map((val, index) => {
        if (!val.subjects || !val.subjects.length) {
            errors.value.set(index, '请至少选择两门科目, 没有选择科目')
            return false;
        } else if (val.subjects.length == 1 && props.subjectIds.length > 1) {  // 支持仅在一门课程时生成排名
            errors.value.set(index, '请至少选择两门科目, 仅选择了一门')
            return false;
        }

        const id = [...val.subjects].sort().join(',');

        if (subjects.has(id)) {
            errors.value.set(index, '已存在 ' + val.subjects.map((val) => subjectIdNameMap[val]).join(', ') + ' 的总分')
            return false;
        };
        subjects.add(id);
    });
}

defineExpose({
    totalColumnsErrors,
    errors
})
</script>