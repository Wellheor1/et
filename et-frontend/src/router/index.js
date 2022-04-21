import Vue from 'vue'
import VueRouter from 'vue-router'
import timeSheet from '@/components/time-Sheet'

Vue.use(VueRouter)

const routes = [
  {
    path: '/timesheet',
    name: 'timesheet',
    component: timeSheet
  }
]

const router = new VueRouter({
  routes
})

export default router
