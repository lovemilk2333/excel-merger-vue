<template>
    <div ref="homePageEl" class="home-page flex flex-col items-center justify-center w-full h-full min-h-0 px-4">
        <div ref="stepperWrapEl" class="stepper-wrap relative w-full max-w-5xl flex flex-col items-stretch">
            <h2 class="home-title shrink-0 mb-6 text-4xl font-bold tracking-tight no-select text-center">
                Excel Merger
            </h2>
            <Stepper v-model:value="step" linear class="stepper-root w-full">
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
                            <p class="mt-2 mb-4">
                                <span class="text-xl">主键字段是一个或多个可以代表唯一学生的字段集合, 用于标识成绩的唯一性</span><br>
                                <span class="text-gray-300">例如, 选择了 "学生" 和 "班级", 那么代表当且仅当 "学生" 和 "班级" 字段值同时相等时,
                                    才视为两者为同一学生的信息和成绩</span>
                            </p>
                            <div class="flex flex-col step-content-inside gap-4">
                                <ColumnsSelector v-if="!!sharedColumns && !!sharedColumns.length"
                                    :columns="sharedColumns" v-model="primaryColumnsIndex"></ColumnsSelector>
                                <Message v-else severity="error">没有找到不同表格内的同名字段, 请回退上一步重新选择文件</Message>
                                <FixedSheetSelector v-if="sheetOptions.length" v-model="fixedSheetId"
                                    :sheet-ids="sheetOptions" />
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
                                <Button label="完成" severity="secondary" icon="pi pi-check-circle" iconPos="right"
                                    @click="resetAll" :disabled="!exported && !exportError" />
                            </div>
                        </div>
                    </StepPanel>
                </StepPanels>
            </Stepper>
        </div>

        <!-- <div v-else class="text-red-500"> -->
        <!-- INVALID FILE STATE, PLEASE REFRESH PAGE OR CONTACT DEVELOPERS !!! -->
        <!-- </div> -->
    </div>
</template>
<style lang="css" scoped>
.home-page {
    overscroll-behavior: contain;
    overflow: hidden;
}

.stepper-wrap {
    /* 唯一默认偏移：改这里即可（% 相对本元素内容高度） */
    --stepper-shift: -7.5%;

    max-width: min(64rem, 100%);
    width: 100%;
    height: auto;
    max-height: calc(100dvh - 1.5rem);
    overflow: hidden;
    flex: 0 0 auto;
    transform: translateY(var(--stepper-shift));
}

.stepper-root {
    /* 不要 flex-1，避免被拉满视口 */
    width: 100%;
    height: auto;
    max-height: none;
    display: block;
}

.step-content {
    display: flex;
    flex-direction: column;
    justify-content: flex-start;
    box-sizing: border-box;
    padding: 1rem;
    max-width: 100%;
}

.step-content>*:not(.step-content-inside) {
    flex-shrink: 0;
}

.step-content-inside {
    /* 内容过高时在此滚动；上限保证整块（title+steps+本区）可被 max-height 约束 */
    max-width: 95dvw;
    max-height: min(50dvh, 28rem);
    overflow-x: hidden;
    overflow-y: auto;
    overscroll-behavior: contain;
    -webkit-overflow-scrolling: touch;
    box-sizing: border-box;
}

.step-content-inside>* {
    max-width: 100%;
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
import { nextTick, onBeforeUnmount, onMounted, ref, useTemplateRef, watch } from 'vue';
import ColumnsSelector from '../components/ColumnsSelector.vue';
import FixedSheetSelector from '../components/FixedSheetSelector.vue';
import { ExcelAdapter } from '../core/core.ts';
import { SUBJECT_ID, type ColumnName, type ExcelSheetId, type JSONSheetRow, type SubjectColumnsType, type SubjectId, type TotalScoreColumn } from '../core/constants.ts';
import SubjectSelector from '../components/SubjectSelector.vue';
import TotalScoreColumnsManager from '../components/TotalScoreColumnsManager.vue';
import { Message } from 'primevue';

const ea = new ExcelAdapter()

const VIEWPORT_EDGE_PAD = 12
/** 与 CSS `--stepper-shift` 默认值一致 */
const DEFAULT_SHIFT_CSS = '-7.5vh'

const homePageEl = ref<HTMLElement | null>(null)
const stepperWrapEl = ref<HTMLElement | null>(null)

function applyShift(value: string) {
    const wrap = stepperWrapEl.value
    if (!wrap) return
    wrap.style.setProperty('--stepper-shift', value)
}

function updateStepperShift() {
    const wrap = stepperWrapEl.value
    const page = homePageEl.value
    if (!wrap || !page) return

    // 1) 先按默认布局并测量
    applyShift(DEFAULT_SHIFT_CSS)
    void wrap.offsetHeight

    const pageRect = page.getBoundingClientRect()
    const wrapRect = wrap.getBoundingClientRect()
    const minTop = pageRect.top + VIEWPORT_EDGE_PAD

    if (wrapRect.top >= minTop - 0.5) {
        // 默认安全，保持
        applyShift(DEFAULT_SHIFT_CSS)
        return
    }

    // 2) 顶出：从无偏移位置算到刚好贴顶
    applyShift('0px')
    void wrap.offsetHeight
    const top0 = wrap.getBoundingClientRect().top
    const dy = minTop - top0
    // dy 通常 ≤ 0；若内容过高，贴顶即可
    applyShift(`${dy}px`)
}

let shiftRaf = 0
let updatingShift = false
function scheduleShiftUpdate() {
    if (updatingShift) return
    cancelAnimationFrame(shiftRaf)
    shiftRaf = requestAnimationFrame(() => {
        nextTick(() => {
            updatingShift = true
            try {
                updateStepperShift()
            } finally {
                requestAnimationFrame(() => {
                    updatingShift = false
                })
            }
        })
    })
}

let resizeObserver: ResizeObserver | null = null

onMounted(() => {
    applyShift(DEFAULT_SHIFT_CSS)
    scheduleShiftUpdate()
    window.addEventListener('resize', scheduleShiftUpdate)
    resizeObserver = new ResizeObserver(() => scheduleShiftUpdate())
    if (stepperWrapEl.value) resizeObserver.observe(stepperWrapEl.value)
    if (homePageEl.value) resizeObserver.observe(homePageEl.value)
})

onBeforeUnmount(() => {
    window.removeEventListener('resize', scheduleShiftUpdate)
    cancelAnimationFrame(shiftRaf)
    resizeObserver?.disconnect()
})

const uploaded = ref(false)
const uploader = useTemplateRef('uploader')
const step = ref(1)
const sharedColumns = ref<ColumnName[]>([])
const primaryColumnsIndex = ref<number[]>([])
/** 已加载 sheet 的 id 列表（步骤 2 锚定选择） */
const sheetOptions = ref<ExcelSheetId[]>([])
/** 锚定 sheetId；null = 不锚定 */
const fixedSheetId = ref<ExcelSheetId | null>(null)
const subjects = ref<SubjectColumnsType | undefined>()
const subjectIds = ref<Exclude<SubjectId, 'none'>[]>([])
const totalScoreColumns = ref<TotalScoreColumn[]>([])
const totalColumns = useTemplateRef('totalColumns')
const mergeErrors = ref<{ row?: JSONSheetRow, error: Error }[]>([])

const exportError = ref(false)
const exportLoading = ref(false)
const exported = ref(false)

function resetAll() {
    uploaded.value = false
    step.value = 1
    sharedColumns.value.length = 0
    primaryColumnsIndex.value.length = 0
    sheetOptions.value.length = 0
    fixedSheetId.value = null
    subjects.value = undefined
    subjectIds.value.length = 0
    totalScoreColumns.value.length = 0
    mergeErrors.value.length = 0

    exportError.value = false
    exportLoading.value = false
    exported.value = false
}

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
    sheetOptions.value.length = 0
    fixedSheetId.value = null
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

watch(step, (val) => {
    if (val === 1) {
        resetFiles()
    } else if (val === 2) {
        const [_sharedColumns, _selectedColumns] = ea.getSharedColumns()

        sharedColumns.value.length = 0
        sharedColumns.value.push(..._sharedColumns)
        primaryColumnsIndex.value.length = 0
        primaryColumnsIndex.value.push(..._selectedColumns)

        sheetOptions.value = ea.getSheetIds()
        // 已选锚定若不在当前列表中则清空
        if (fixedSheetId.value && !sheetOptions.value.includes(fixedSheetId.value)) {
            fixedSheetId.value = null
        }
    } else if (val === 3) {
        const primaryColumns = sharedColumns.value.filter((_, index) => primaryColumnsIndex.value.includes(index))
        ea.setPrimaryColumns(primaryColumns)
        ea.setFixedSource(fixedSheetId.value)
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
                rank: true,
                skipAtOrBelow: null,
            })
        }
    } else if (val === 5) {
        // 再次同步锚定（防止中途改过）
        ea.setFixedSource(fixedSheetId.value)
        for (const col of totalScoreColumns.value) {
            if (!col.subjects || !col.subjects.length) continue

            ea.addTotalScoreColumn(col)
        }

    }

    scheduleShiftUpdate()
})

watch(subjects, () => scheduleShiftUpdate(), { deep: true })
</script>
