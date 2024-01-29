## 如何用vite搭建vue3项目

* 终端命令进入到项目文件夹，输入以下命令
* npm create vite@latest

## vue中引入echarts

* 安装echsrts
* npm install echarts -S

* 在父组件App.vue中通过依赖注入引入echarts

```js
//父组件App.vue
<script setup>
  import { provide } from 'vue'
  import * as echarts from 'echarts'
  provide('echarts', echarts)
</script>
```

* 在子组件中获取父组件的全局传值
```js
//子组件
<script setup>
import { onMounted, inject } from 'vue'
const $echarts = inject('echarts')
onMounted(() => {
    // 基于准备好的dom，初始化echarts实例
    var myChart = $echarts.init(document.getElementById('main'));
})
</script>
```
