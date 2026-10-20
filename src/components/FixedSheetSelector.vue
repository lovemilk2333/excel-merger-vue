<template>
    <div class="fixed-sheet-selector flex flex-col gap-2 no-select">
        <div class="flex flex-wrap items-center gap-3">
            <Checkbox v-model="enabled" inputId="fixed-sheet-enabled" binary @update:model-value="onEnabledChange" />
            <label for="fixed-sheet-enabled" class="select-none text-sm cursor-pointer">
                <p class="mt-2 mb-2">
                    <span class="text-base">锚定 Excel 表</span><br>
                    <span class="text-gray-300 text-sm">选定锚定表后, 合并后的 Excel 会按锚定表内顺序输出被锚定的表中包含的学生</span>
                </p>
            </label>
        </div>
        <div class="flex flex-wrap items-center gap-2 mb-4" v-if="enabled">
            <Select inputId="fixed-sheet-select" :options="options" optionLabel="label" optionValue="value" v-model="selected" placeholder="选择一个 Sheet" class="min-w-72 flex-1" fluid :disabled="!options.length" />
        </div>
    </div>
</template>

<script setup lang="ts">
import { computed, ref, watch } from 'vue'
import Checkbox from 'primevue/checkbox'
import Select from 'primevue/select'
import type { ExcelSheetId, ExcelSheetIdParsed } from '../core/constants'

export interface FixedSheetOption {
    value: ExcelSheetId
    label: string
    filename: string
    sheetname: string
}

const props = defineProps<{
    /** 可选 sheet 列表（sheetId） */
    sheetIds: ExcelSheetId[]
    /** 当前锚定 sheetId；null 表示不锚定 */
    modelValue: ExcelSheetId | null
}>()

const emit = defineEmits<{
    (e: 'update:modelValue', value: ExcelSheetId | null): void
}>()

function parseSheetId(id: ExcelSheetId): ExcelSheetIdParsed {
    try {
        return JSON.parse(id) as ExcelSheetIdParsed
    } catch {
        return { filename: id, sheetname: '' }
    }
}

const options = computed<FixedSheetOption[]>(() =>
    props.sheetIds.map((id) => {
        const { filename, sheetname } = parseSheetId(id)
        return {
            value: id,
            filename,
            sheetname,
            label: sheetname ? `${filename} › ${sheetname}` : filename,
        }
    })
)

const enabled = ref(props.modelValue != null)

watch(
    () => props.modelValue,
    (val) => {
        enabled.value = val != null
    }
)

const selected = computed<ExcelSheetId | null>({
    get() {
        return props.modelValue
    },
    set(val) {
        emit('update:modelValue', val)
    },
})

function onEnabledChange(on: boolean | undefined) {
    const enabledNow = !!on
    enabled.value = enabledNow
    if (!enabledNow) {
        emit('update:modelValue', null)
        return
    }
    // 开启：保留已有选择，否则默认第一项
    if (props.modelValue != null) return
    const first = options.value[0]?.value ?? null
    emit('update:modelValue', first)
}

// sheet 列表变化时，若当前值已不在列表中则清空或回落到第一项
watch(
    () => props.sheetIds,
    (ids) => {
        if (props.modelValue == null) return
        if (!ids.includes(props.modelValue)) {
            if (enabled.value && ids.length) {
                emit('update:modelValue', ids[0])
            } else {
                emit('update:modelValue', null)
                enabled.value = false
            }
        }
    }
)
</script>
