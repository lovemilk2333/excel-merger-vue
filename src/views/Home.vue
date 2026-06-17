<template>
    <div class="flex flex-col items-center justify-center w-full h-full">
        <h2 class="text-4xl font-bold tracking-tight my-8 no-select">Excel Merger</h2>

        <Stepper v-model:value="step" linear>
            <StepList>
                <Step :value="1">选择 Excel 文件</Step>
                <Step :value="2">选择主键字段</Step>
                <Step :value="3">Header III</Step>
            </StepList>
            <StepPanels>
                <StepPanel v-slot="{ activateCallback }" :value="1">
                    <div class="step-content">
                        <div class="flex flex-col">
                            <FileUploader ref="uploader"
                                summary="双击或拖拽到此处以选择 Excel (<code>.xls</code> 或 <code>.xlsx</code>) 文件"
                                @upload="setFiles">
                            </FileUploader>
                        </div>
                        <div class="flex pt-2 justify-end">
                            <Button label="下一步" icon="pi pi-arrow-right" iconPos="right" @click="activateCallback(2);"
                                :disabled="!uploaded" />
                        </div>
                    </div>
                </StepPanel>
                <StepPanel v-slot="{ activateCallback }" :value="2">
                    <div class="step-content">
                        <div class="flex flex-col">
                            <p class="mt-2 mb-4"><span class="text-xl">主键字段是一个或多个可以代表唯一学生的字段集合, 用于标识成绩的唯一性</span><br><span class="text-gray-300">例如, 选择了 "学生" 和 "班级", 那么代表当且仅当 "学生" 和 "班级" 字段值同时相等时, 才视为两者为同一学生的信息和成绩</span></p>
                            <ColumnsSelector class="mb-4" :columns="sharedColumns" v-model="primaryColumnsIndex"></ColumnsSelector>
                        </div>
                        <div class="flex pt-2 justify-between">
                            <Button label="上一步" severity="secondary" icon="pi pi-arrow-left"
                                @click="resetFiles(); activateCallback(1)" />
                            <Button label="下一步" icon="pi pi-arrow-right" iconPos="right" @click="activateCallback(3)" />
                        </div>
                    </div>
                </StepPanel>
                <StepPanel v-slot="{ activateCallback }" :value="3">
                    <div class="step-content">
                        <div class="flex flex-col">
                            <div
                                class="border-2 border-dashed border-surface-200 dark:border-surface-700 rounded bg-surface-50 dark:bg-surface-950 flex-auto flex justify-center items-center font-medium">
                                Content III</div>
                        </div>
                        <div class="pt-2">
                            <Button label="Back" severity="secondary" icon="pi pi-arrow-left"
                                @click="activateCallback(2)" />
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

<script setup lang="ts">
import FileUploader from '../components/FileUploader.vue';
import Stepper from 'primevue/stepper';
import StepList from 'primevue/steplist';
import StepPanels from 'primevue/steppanels';
import StepItem from 'primevue/stepitem';
import Step from 'primevue/step';
import StepPanel from 'primevue/steppanel';
import Button from 'primevue/button';
import { computed, onMounted, reactive, ref, useTemplateRef, watch } from 'vue';
import ColumnsSelector from '../components/ColumnsSelector.vue';
import { ExcelAdapter } from '../core/core.ts';
import type { ColumnName } from '../core/constants.ts';

const ea = new ExcelAdapter()

const uploaded = ref(false)
const uploader = useTemplateRef('uploader')
const step = ref(1)
const sharedColumns = ref<ColumnName[]>([])
const primaryColumnsIndex = ref<number[]>([])

function setFiles(uploadedFiles: File[]) {
    ea.read(uploadedFiles)
    uploaded.value = true
}

function resetFiles() {
    uploaded.value = false
    ea.init()
    uploader.value?.reset();
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
    }
})
</script>

<style lang="css" scoped>
.step-content {
    padding: 1rem;
}
</style>
