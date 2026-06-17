<template>
    <FileUpload name="files[]" @select="$emit('select', $event)" :multiple="true" :disabled="uploaded"
        accept="application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/wps-office.xlsx,application/wps-office.xls,.xls,.xlsx"
        :pt="{
            header: {
                style: { display: 'none' },
            },
            content: {
                style: {
                    padding: '0'
                }
            }
        }">
        <template #header="{ chooseCallback, clearCallback }">
            <button @click="chooseCallback()" class="hidden" ref="chooseFilesButton"></button>
            <button @click="clearCallback()" class="hidden" ref="clearFilesButton"></button>
        </template>
        <template #empty>
            <div v-if="!uploaded" class="flex items-center justify-center flex-col no-select cursor-pointer p-4.5"
                @dblclick="chooseFilesButton?.click()">
                <i class="pi pi-cloud-upload border-2! rounded-full! p-8! text-4xl! text-muted-color!" />
                <p class="mt-6 mb-0" v-if="!!$props.summary && !!$props.summary.length"><span
                        v-html="$props.summary"></span></p>
            </div>
            <div v-else class="flex items-center justify-center flex-col no-select cursor-default p-4.5"
                @dblclick="chooseFilesButton?.click()">
                <i class="pi pi-check-circle border-2! rounded-full! p-8! text-4xl! text-muted-color!" />
                <p class="mt-6 mb-0" v-if="!!$props.uploadedSummary && !!$props.uploadedSummary.length"><span
                        v-html="$props.uploadedSummary"></span>
                </p>
            </div>
        </template>
        <template #content="{ files, removeFileCallback }">
            <div class="flex flex-col gap-4 pt-4 ml-4 mr-4 mb-4" @dblclick="chooseFilesButton?.click()"
                v-if="files.length > 0">
                <div class="flex">
                    <Button @click="onUpload(files)" label="确认" icon="pi pi-check-circle" rounded variant="outlined"
                        severity="success" :disabled="!files || files.length === 0"></Button>
                </div>

                <div class="flex flex-wrap gap-4">
                    <div v-for="(file, index) of files" :key="file.name + file.type + file.size"
                        class="p-8 rounded-border flex flex-col border rounded-xl border-surface items-center gap-4"
                        @dblclick.stop>
                        <span class="font-semibold text-ellipsis max-w-60 whitespace-nowrap overflow-hidden">{{
                            file.name }}</span>
                        <div class="flex justify-center w-full gap-2">
                            <p>{{ formatSize(file.size) }}</p>
                            <Badge value="等待确认" severity="warn" />
                        </div>
                        <Button icon="pi pi-times" @click="removeFileCallback(index)" variant="outlined" rounded
                            severity="danger" />
                    </div>
                </div>
            </div>
        </template>
    </FileUpload>
</template>

<script setup lang="ts">
import { Badge, Button } from 'primevue';
import FileUpload, { type FileUploadSelectEvent } from 'primevue/fileupload';
import { onMounted, ref, useTemplateRef } from 'vue';
import { formatSize } from '../core/utils';

const chooseFilesButton = useTemplateRef('chooseFilesButton')
const clearFilesButton = useTemplateRef('clearFilesButton')
onMounted(() => {
    chooseFilesButton.value?.remove()
    clearFilesButton.value?.remove()
})

const uploaded = ref(false)

defineProps({
    summary: { type: String, default: null },
    uploadedSummary: { type: String, default: '成功' },
})
const emit = defineEmits<{
    select: [event: FileUploadSelectEvent],
    upload: [event: File[]]
}>()

function onUpload(files: File[]) {
    clearFilesButton.value?.click()
    emit('upload', files)
    uploaded.value = true
}

function reset() {
    clearFilesButton.value?.click()
    uploaded.value = false
}

defineExpose({
    reset
})
</script>
