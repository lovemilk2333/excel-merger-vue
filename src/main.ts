import { createApp } from 'vue'
import './style.css'
import './tw.css'
import 'primeicons/primeicons.css'
import App from './App.vue'
import VueExcelEditor from 'vue3-excel-editor'
import PrimeVue from 'primevue/config';
import Aura from '@primeuix/themes/aura';
import { router } from './routers.ts';

const app = createApp(App)
app.use(VueExcelEditor)
app.use(PrimeVue, {
    theme: {
        preset: Aura
    }
})
app.use(router)
app.mount('#app')
