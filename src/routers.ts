import { createWebHashHistory, createRouter } from 'vue-router'
import Home from './views/Home.vue'
import Dev from './views/Dev.vue'

const isDev = import.meta.env.DEV
const routes = [
  {
    path: '/', component: Home
  },
]

if (isDev) {
  routes.push({
    path: '/dev', component: Dev
  })
}

export const router = createRouter({
  history: createWebHashHistory(),
  routes,
})
