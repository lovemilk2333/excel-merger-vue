<template>
    <div class="flex flex-col p-4 gap-4">
        <Card class="sticky top-4 z-100">
            <template #content>
                <div class="flex flex-col gap-2">
                    <div class="flex items-center gap-2">
                        <Button label="reset" @click="reset"></Button>
                        <Button label="refresh subjects" @click="totalColumns?.refreshSubjects()"></Button>
                        <Button label="merge" @click="merge" :disabled="!!totalColumns?.errors.size"></Button>
                    </div>
                    <code>selectedColumns: {{ selectedColumns }}</code>
                    <code>totalScoreColumns: {{ totalScoreColumns }}</code>
                </div>
            </template>
        </Card>
        <FileUploader ref="uploader" @upload="onUpload" summary="upload test .xls/.xlsx file"></FileUploader>
        <ColumnsSelector :columns="columns" v-model="selectedColumns"></ColumnsSelector>
        <div class="flex flex-col mt-4 gap-4" v-if="!!subjects">
            <div class="flex flex-col w-full gap-2" v-for="[id, subjectMap] in subjects.entries()" :key="id">
                <code>{{ id }}</code>
                <code>{{ JSON.stringify(Object.fromEntries(subjectMap)) }}</code>
                <SubjectSelector :title="id" :model-value="subjectMap"
                    @update:model-value="(nextMap) => subjects?.set(id, nextMap)" />
            </div>
        </div>
        <TotalScoreColumnsManager v-model="totalScoreColumns" :ea="ea" ref="totalColumns"></TotalScoreColumnsManager>
    </div>
</template>
<script setup lang="ts">
import ColumnsSelector from '../components/ColumnsSelector.vue';
import { ref, useTemplateRef } from 'vue';
import FileUploader from '../components/FileUploader.vue';
import { ExcelAdapter } from '../core/core.ts';
import type { ColumnName, ExcelSheetId, SubjectId, TotalScoreColumn } from '../core/constants.ts';
import { Button, Card } from 'primevue';
import SubjectSelector from '../components/SubjectSelector.vue';
import TotalScoreColumnsManager from '../components/TotalScoreColumnsManager.vue';
import { utils, writeFile } from 'xlsx';

const ea = new ExcelAdapter()
const columns = ref<ColumnName[]>([])
const selectedColumns = ref<number[]>([])
const uploader = useTemplateRef('uploader')
const subjects = ref<Map<ExcelSheetId, Map<ColumnName, SubjectId>> | undefined>()
const totalScoreColumns = ref<TotalScoreColumn[]>([])
// const showTotalScoreColumns = ref(false)
const totalColumns = useTemplateRef('totalColumns')

async function onUpload(files: File[]) {
    await ea.read(files)
    console.debug('getPrimaryColumns', ea.getSharedColumns())
    // setTimeout(() => {
    //     uploader.value?.reset()
    //     alert('reset')
    // }, 1000)
    const [sharedColumns, _selectedColumns] = ea.getSharedColumns()
    columns.value.push(...sharedColumns)
    selectedColumns.value.push(..._selectedColumns)
    subjects.value = ea.getSubjectColumns()
    const subjectids = ea.getSubjectIds()
    totalScoreColumns.value.push({
        subjects: subjectids,
        rank: true
    })
    totalColumns.value?.refreshSubjects()
}

function reset() {
    uploader.value?.reset()
    columns.value.length = 0
    selectedColumns.value.length = 0
    ea.init()
    subjects.value = undefined
    totalScoreColumns.value.length = 0
    totalColumns.value?.refreshSubjects()
}

function merge() {
    if (!!totalColumns.value?.errors.size) return
    ea.setPrimaryColumns(ea.getSharedColumns()[0].filter((_, index) => selectedColumns.value.includes(index)))

    totalScoreColumns.value.forEach((val) => {
        ea.addTotalScoreColumn(val)
    })
    const [worksheet, data, errors] = ea.merge()
    // console.debug('data', data)
    console.debug('errors', errors)

    const workbook = utils.book_new();

    utils.book_append_sheet(workbook, worksheet, 'result');

    writeFile(workbook, 'output.xlsx');
}
</script>
