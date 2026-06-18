<template>
    <div class="flex flex-col items-center justify-center w-full h-full">
        <h2 class="text-4xl font-bold tracking-tight my-8 no-select">Excel Merger</h2>
        <Stepper v-model:value="step" linear>
            <StepList>
                <Step :value="1">选择 Excel 文件</Step>
                <Step :value="2">选择主键字段</Step>
                <Step :value="3">设置科目-列对应关系</Step>
                <Step :value="4">设置总分列</Step>
                <Step :value="5">合并并导出</Step>
            </StepList>
            <StepPanels>
                <StepPanel :value="1">
                    <div class="step-content">
                        <div class="flex flex-col step-content-inside">
                            <FileUploader ref="uploader"
                                summary="双击或拖拽到此处以选择 Excel (<code>.xls</code> 或 <code>.xlsx</code>) 文件"
                                @upload="setFiles">
                            </FileUploader>
                        </div>
                        <div class="flex pt-2 justify-end">
                            <Button label="下一步" icon="pi pi-arrow-right" iconPos="right" @click="step++"
                                :disabled="!uploaded" />
                        </div>
                    </div>
                </StepPanel>
                <StepPanel :value="2">
                    <div class="step-content">
                        <p class="mt-2 mb-4"><span class="text-xl">主键字段是一个或多个可以代表唯一学生的字段集合,
                                用于标识成绩的唯一性</span><br><span class="text-gray-300">例如, 选择了 "学生" 和 "班级", 那么代表当且仅当 "学生"
                                和 "班级" 字段值同时相等时, 才视为两者为同一学生的信息和成绩</span></p>
                        <div class="flex flex-col step-content-inside">
                            <ColumnsSelector v-if="!!sharedColumns && !!sharedColumns.length" class="mb-4"
                                :columns="sharedColumns" v-model="primaryColumnsIndex"></ColumnsSelector>
                            <Message v-else severity="error">没有找到不同表格内的同名字段, 请回退上一步重新选择文件</Message>
                        </div>
                        <div class="flex pt-2 justify-between">
                            <Button label="上一步" severity="secondary" icon="pi pi-arrow-left" @click="step--" />
                            <Button label="下一步" icon="pi pi-arrow-right" iconPos="right" @click="step++"
                                :disabled="!primaryColumnsIndex.length" />
                        </div>
                    </div>
                </StepPanel>
                <StepPanel :value="3">
                    <div class="step-content">
                        <p class="mt-2 mb-4"><span class="text-xl">通过设置科目-列对应关系, 可以将表格的某一列视为某科成绩, 并绑定到某个学生上,
                                以便合并数据</span></p>
                        <Button label="重新自动识别科目-列映射" severity="secondary" icon="pi pi-filter"
                            @click="mergeSubjects(ea.getSubjectColumns())"></Button>
                        <div class="flex flex-col step-content-inside" v-if="!!subjects">
                            <div class="flex flex-col w-full gap-2" v-for="[id, subjectMap] in subjects.entries()"
                                :key="id">
                                <SubjectSelector
                                    :title='`${JSON.parse(id).filename || "-"}<span class="pi pi-angle-right"></span>${JSON.parse(id).sheetname || "-"}`'
                                    :model-value="subjectMap" />
                            </div>
                        </div>
                        <Message v-else severity="error">错误状态: 科目配置为空, 请回退上一步重新选择文件</Message>
                        <div class="flex pt-2 justify-between">
                            <Button label="上一步" severity="secondary" icon="pi pi-arrow-left" @click="step--" />
                            <Button label="下一步" icon="pi pi-arrow-right" iconPos="right" @click="step++"
                                :disabled="!subjects?.size" />
                        </div>
                    </div>
                </StepPanel>
                <StepPanel :value="4">
                    <div class="step-content">
                        <p class="mt-2 mb-4"><span class="text-xl">添加总分列可以新增总分和排名信息, 以便用户查看</span></p>
                        <div class="flex flex-col step-content-inside">
                            <TotalScoreColumnsManager v-model="totalScoreColumns" :subject-ids="subjectIds"
                                ref="totalColumns"></TotalScoreColumnsManager>
                        </div>
                        <div class="flex pt-2 justify-between">
                            <Button label="上一步" severity="secondary" icon="pi pi-arrow-left" @click="step--" />
                            <Button label="下一步" icon="pi pi-arrow-right" iconPos="right" @click="step++"
                                :disabled="!!totalColumns?.errors.size" />
                        </div>
                    </div>
                </StepPanel>
                <StepPanel :value="5">
                    <div class="step-content">
                        <p class="mt-2 mb-4"><span class="text-xl">合并数据</span></p>
                        <div class="flex flex-col">
                            <Button class="mb-4" :label="exportError ? '重试导出' : '合并并导出'"
                                :icon="`pi ${exportError ? 'pi-refresh' : 'pi-file-export'}`"
                                :variant="exportError ? 'outlined' : ''" :severity="exportError ? 'danger' : ''"
                                iconPos="right" @click="mergeAndExport" :loading="exportLoading" />
                            <Message v-if="!!mergeErrors && !!mergeErrors.length" severity="error">合并出现错误,
                                请将日志发送给开发者:<span v-for="error in mergeErrors"><br><code>{{ error }}</code></span>
                            </Message>
                            <Message class="mb-4" v-else-if="exported" severity="success">导出成功! 请查看浏览器下载菜单</Message>
                        </div>
                        <div class="flex pt-2 justify-between">
                            <Button label="上一步" severity="secondary" icon="pi pi-arrow-left" @click="step--" />
                            <Button label="完成" severity="secondary" icon="pi pi-check-circle" iconPos="right" @click="step = 1" :disabled="!exported && !exportError"/>
                        </div>
                    </div>
                </StepPanel>
            </StepPanels>
        </Stepper>

        <!-- <div v-else class="text-red-500"> -->
        <!-- INVALID FILE STATE, PLEASE REFRESH PAGE OR CONTACT DEVELOPERS !!! -->
        <!-- </div> -->
    </div>
</template>
<style lang="css" scoped>
.step-content {
    display: flex;
    flex-direction: column;
    justify-content: center;
    box-sizing: border-box;
}

.step-content>*,
.step-content-inside>* {
    flex-shrink: 0;
}

.step-content-inside {
    overflow-y: scroll;
    max-width: 95dvw;
    max-height: 70dvh;
    box-sizing: border-box;
}
</style>
<script setup lang="ts">
import FileUploader from '../components/FileUploader.vue';
import Stepper from 'primevue/stepper';
import StepList from 'primevue/steplist';
import StepPanels from 'primevue/steppanels';
import Step from 'primevue/step';
import StepPanel from 'primevue/steppanel';
import Button from 'primevue/button';
import { ref, useTemplateRef, watch } from 'vue';
import ColumnsSelector from '../components/ColumnsSelector.vue';
import { ExcelAdapter } from '../core/core.ts';
import { SUBJECT_ID, type ColumnName, type JSONSheetRow, type SubjectColumnsType, type SubjectId, type TotalScoreColumn } from '../core/constants.ts';
import SubjectSelector from '../components/SubjectSelector.vue';
import TotalScoreColumnsManager from '../components/TotalScoreColumnsManager.vue';
import { Message } from 'primevue';

const ea = new ExcelAdapter()

const uploaded = ref(false)
const uploader = useTemplateRef('uploader')
const step = ref(1)
const sharedColumns = ref<ColumnName[]>([])
const primaryColumnsIndex = ref<number[]>([])
const subjects = ref<SubjectColumnsType | undefined>()
const subjectIds = ref<Exclude<SubjectId, 'none'>[]>([])
const totalScoreColumns = ref<TotalScoreColumn[]>([])
const totalColumns = useTemplateRef('totalColumns')
const mergeErrors = ref<{ row?: JSONSheetRow, error: Error }[]>([])

const exportError = ref(false)
const exportLoading = ref(false)
const exported = ref(false)

function setFiles(uploadedFiles: File[]) {
    ea.read(uploadedFiles)
    uploaded.value = true
}

function resetFiles() {
    uploaded.value = false
    ea.init()
    uploader.value?.reset();
    subjects.value = void 0
    primaryColumnsIndex.value.length = 0
    sharedColumns.value.length = 0
}

function filterInvalidSubjects() {
    if (!subjects.value) return;

    for (const map of subjects.value.values()) {

        for (const [columnName, subjectId] of map.entries()) {
            const isInvalidColumn = !columnName || String(columnName).length === 0;
            const isInvalidSubject = !subjectId || String(subjectId).length === 0 || subjectId === SUBJECT_ID.NONE;

            if (isInvalidColumn || isInvalidSubject) {
                map.set(columnName, SUBJECT_ID.NONE)
            }
        }
    }
}

function mergeAndExport() {
    exportLoading.value = true
    mergeErrors.value.length = 0
    let worksheet
    try {
        const r = ea.merge()
        worksheet = r[0]
        const errors = r[2]
        if (errors.length) {
            mergeErrors.value.push(...errors)
            exportLoading.value = false
            exportError.value = true
            return
        }
    } catch (e) {
        mergeErrors.value.push({
            row: void 0,
            error: e as Error
        })
        exportLoading.value = false
        exportError.value = true
        return
    }

    ea.export(worksheet)

    exportLoading.value = false
    exportError.value = false
    exported.value = true
}

function mergeSubjects(newSubjects: SubjectColumnsType) {
    if (!subjects.value) {
        subjects.value = newSubjects
        return
    }

    for (const [id, map] of newSubjects.entries()) {
        if (!subjects.value.has(id)) {
            subjects.value.set(id, map)
            continue
        }
        const parent = subjects.value.get(id)

        for (const [columnName, subjectId] of map.entries()) {
            parent?.set(columnName, subjectId)
        }
    }
}

watch(step, (val, oldVal) => {
    if (val === 1) {
        resetFiles()
    } else if (val === 2) {
        const [_sharedColumns, _selectedColumns] = ea.getSharedColumns()

        sharedColumns.value.length = 0
        sharedColumns.value.push(..._sharedColumns)
        primaryColumnsIndex.value.length = 0
        primaryColumnsIndex.value.push(..._selectedColumns)
    } else if (val === 3) {
        const primaryColumns = sharedColumns.value.filter((_, index) => primaryColumnsIndex.value.includes(index))
        ea.setPrimaryColumns(primaryColumns)
        if (!subjects.value) {
            mergeSubjects(ea.getSubjectColumns())
        }
    } else if (val === 4) {
        ea.setSubjectColumns(subjects.value!)
        filterInvalidSubjects()
        ea.clearTotalScoreColumn()
        subjectIds.value = ea.getSubjectIds()
        totalScoreColumns.value.length = 0
        if (!!subjectIds.value.length) {
            totalScoreColumns.value.push({
                subjects: subjectIds.value,
                rank: true
            })
        }
    } else if (val === 5) {
        for (const col of totalScoreColumns.value) {
            if (!col.subjects || !col.subjects.length) continue

            ea.addTotalScoreColumn(col)
        }

    }
})
</script>

<style lang="css" scoped>
.step-content {
    padding: 1rem;
}
</style>
