<template>
    <FileUpload name="files[]" @select="$emit('select', $event)" :multiple="true" :disabled="uploaded"
        accept="application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/wps-office.xlsx,application/wps-office.xls,.xls,.xlsx"
        class="file-uploader w-full" :pt="{
            root: {
                class: 'file-uploader-root w-full',
            },
            header: {
                style: { display: 'none' },
            },
            content: {
                class: 'file-uploader-content',
                style: {
                    padding: '0',
                    overflow: 'visible',
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
            <div class="file-list-panel flex flex-col gap-3 p-4" v-if="files.length > 0"
                @dblclick="chooseFilesButton?.click()">
                <div class="file-list-actions shrink-0 flex items-center gap-2">
                    <Button @click.stop="onUpload(files)" label="确认" icon="pi pi-check-circle" rounded
                        variant="outlined" severity="success" :disabled="!files || files.length === 0"></Button>
                    <span class="text-sm text-muted-color">已选 {{ files.length }} 个文件</span>
                </div>
                <Divider class="mt-0! mb-2!" />

                <div class="file-list-scroll flex flex-wrap gap-4 content-start" @dblclick.stop>
                    <div v-for="(file, index) of files" :key="file.name + file.type + file.size"
                        class="file-card relative p-6 rounded-border flex flex-col border rounded-xl border-surface items-center gap-3"
                        @mouseenter="openNameTip(fileKey(file, index))"
                        @mouseleave="scheduleCloseNameTip(fileKey(file, index))"
                        @focusin="openNameTip(fileKey(file, index))"
                        @focusout="scheduleCloseNameTip(fileKey(file, index))" @dblclick.stop>
                        <span
                            class="font-semibold text-ellipsis max-w-60 whitespace-nowrap overflow-hidden select-none">{{
                                file.name }}</span>
                        <div class="flex justify-center w-full gap-2 items-center">
                            <p class="m-0">{{ formatSize(file.size) }}</p>
                            <Badge value="等待确认" severity="warn" />
                        </div>
                        <Button icon="pi pi-times" @click.stop="removeFileCallback(index)" variant="outlined" rounded
                            severity="danger" />

                        <div v-show="activeNameTip === fileKey(file, index)"
                            class="file-name-tip absolute left-1/2 z-20 -translate-x-1/2 bottom-[calc(100%+0.35rem)] w-max max-w-80"
                            role="tooltip" @mouseenter="openNameTip(fileKey(file, index))"
                            @mouseleave="scheduleCloseNameTip(fileKey(file, index))" @mousedown.stop @click.stop>
                            <div class="file-name-tip-inner flex items-start gap-2 rounded-lg shadow-lg px-3 py-2">
                                <span class="file-name-tip-text text-sm break-all select-text cursor-text"
                                    tabindex="0">{{ file.name }}</span>
                                <!-- <Button type="button" icon="pi pi-copy" size="small" text rounded
                                    :severity="copiedKey === fileKey(file, index) ? 'success' : 'secondary'"
                                    :aria-label="`复制文件名 ${file.name}`"
                                    v-tooltip.top="copiedKey === fileKey(file, index) ? '已复制' : '复制文件名'"
                                    @click.stop="copyFileName(file.name, fileKey(file, index))" /> -->
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </template>
    </FileUpload>
</template>

<script setup lang="ts">
import { Badge, Button } from 'primevue';
import FileUpload, { type FileUploadSelectEvent } from 'primevue/fileupload';
import { onBeforeUnmount, onMounted, ref, useTemplateRef } from 'vue';
import { formatSize } from '../core/utils';

const chooseFilesButton = useTemplateRef('chooseFilesButton')
const clearFilesButton = useTemplateRef('clearFilesButton')
onMounted(() => {
    chooseFilesButton.value?.remove()
    clearFilesButton.value?.remove()
})

const uploaded = ref(false)
/** 当前展示完整文件名浮层的卡片 key */
const activeNameTip = ref<string | null>(null)
const copiedKey = ref<string | null>(null)
let closeTipTimer: ReturnType<typeof setTimeout> | null = null
let copiedClearTimer: ReturnType<typeof setTimeout> | null = null

function fileKey(file: File, index: number) {
    return `${index}:${file.name}:${file.size}`
}

function openNameTip(key: string) {
    if (closeTipTimer) {
        clearTimeout(closeTipTimer)
        closeTipTimer = null
    }
    activeNameTip.value = key
}

function scheduleCloseNameTip(key: string) {
    if (closeTipTimer) clearTimeout(closeTipTimer)
    // 短暂延迟，便于鼠标移入浮层选中/复制
    closeTipTimer = setTimeout(() => {
        if (activeNameTip.value === key) {
            activeNameTip.value = null
        }
        closeTipTimer = null
    }, 180)
}

// async function copyFileName(name: string, key: string) {
//     try {
//         if (navigator.clipboard?.writeText) {
//             await navigator.clipboard.writeText(name)
//         } else {
//             const ta = document.createElement('textarea')
//             ta.value = name
//             ta.setAttribute('readonly', '')
//             ta.style.position = 'fixed'
//             ta.style.left = '-9999px'
//             document.body.appendChild(ta)
//             ta.select()
//             document.execCommand('copy')
//             document.body.removeChild(ta)
//         }
//         copiedKey.value = key
//         if (copiedClearTimer) clearTimeout(copiedClearTimer)
//         copiedClearTimer = setTimeout(() => {
//             if (copiedKey.value === key) copiedKey.value = null
//             copiedClearTimer = null
//         }, 1500)
//     } catch (e) {
//         console.warn('copy file name failed', e)
//     }
// }

onBeforeUnmount(() => {
    if (closeTipTimer) clearTimeout(closeTipTimer)
    if (copiedClearTimer) clearTimeout(copiedClearTimer)
})

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
    activeNameTip.value = null
    copiedKey.value = null
}

defineExpose({
    reset
})
</script>

<style scoped>
.file-list-scroll {
    max-height: 45dvh;
    overflow-x: hidden;
    overflow-y: auto;
    overscroll-behavior: contain;
    -webkit-overflow-scrolling: touch;
    padding-right: 0.25rem;
    padding-bottom: 0.25rem;
}

.file-list-actions {
    position: sticky;
    top: 0;
    z-index: 2;
    background: var(--p-content-background, var(--p-surface-0, #fff));
}

:global(.p-dark) .file-list-actions,
:global(html.app-dark) .file-list-actions,
:global(.dark) .file-list-actions {
    background: var(--p-content-background, var(--p-surface-900, #18181b));
}

.file-name-tip {
    pointer-events: auto;
}

.file-name-tip-inner {
    pointer-events: auto;
    /* PrimeVue Aura content 背景（不依赖可能未生成的 tailwind surface 类） */
    background-color: var(--p-content-background, var(--p-surface-0, #ffffff));
    color: var(--p-text-color, var(--p-surface-700, #334155));
    border: 1px solid var(--p-content-border-color, var(--p-surface-200, #e2e8f0));
    box-shadow:
        0 10px 15px -3px rgb(0 0 0 / 0.12),
        0 4px 6px -4px rgb(0 0 0 / 0.1);
}

@media (prefers-color-scheme: dark) {
    .file-name-tip-inner {
        background-color: var(--p-content-background, var(--p-surface-900, #18181b));
        color: var(--p-text-color, #ffffff);
        border-color: var(--p-content-border-color, var(--p-surface-700, #3f3f46));
    }
}

/* 若应用显式 dark 类，优先于系统偏好 */
:global(.p-dark) .file-name-tip-inner,
:global(html.app-dark) .file-name-tip-inner,
:global(.dark) .file-name-tip-inner {
    background-color: var(--p-content-background, var(--p-surface-900, #18181b));
    color: var(--p-text-color, #ffffff);
    border-color: var(--p-content-border-color, var(--p-surface-700, #3f3f46));
}

.file-name-tip-text {
    user-select: text;
    -webkit-user-select: text;
    max-width: 16rem;
    line-height: 1.35;
    color: inherit;
}
</style>
