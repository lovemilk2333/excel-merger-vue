<template>
    <div class="columns-selector flex items-center gap-4 no-select">
        <div class="flex items-center gap-1" v-for="(columnName, index) in columns" :key="index">
            <Checkbox v-model="selected" :inputId="`column-${index}`" :name="columnName" :value="index" />
            <label :for="`column-${index}`" class="cursor-pointer"> {{ columnName }} </label>
        </div>
    </div>
</template>

<script setup lang="ts">
import { computed } from 'vue'
import type { ColumnName, JSONSheetRows } from '../core/constants';
import Checkbox from 'primevue/checkbox';

const props = defineProps<{
    columns: ColumnName[],
    modelValue: number[],
    sheet?: JSONSheetRows,
}>()

const emit = defineEmits<{
    (e: 'update:modelValue', value: number[]): void
}>()

const selected = computed<number[]>({
    get() {
        return props.modelValue || [];
    },
    set(newVal) {
        emit('update:modelValue', newVal);
    }
});
</script>
