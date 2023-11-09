<template>
  <div class="app">
      <div class="app-radio">
        <a-radio-group v-model:value="methodChecked">
          <a-radio :style="radioStyle" :value="1">
            全局去重
          </a-radio>
          <a-radio :style="radioStyle" :value="2">
            单列去重
          </a-radio>
        </a-radio-group>
      </div>
      <div v-if="methodChecked === 2" class="app-cols-input">
        <span>请输入单列去重标题</span>
        <a-input v-model:value="colsName" :maxlength="20" :style="{ width: '150px' }"></a-input>
      </div>
      <a-upload
          v-model:file-list="fileList"
          name="file"
          @beforeUpload="beforeUpload"
          @change="handleFileUpload"
      >
          <a-button>
              <UploadOutlined />
              点击上传 EXCEL 文件
          </a-button>
      </a-upload>
      <p class="app-tips">
        Tips: <br />
        1.EXCEL文件每列需要表格标题做唯一标识 <br />
        2.文件必须为xlsx格式<br />
        3.导出文件名格式为源文件名_时间_已去重 <br />
      </p>
  </div>
</template>

<script setup>
import { ref } from 'vue'
import * as XLSX from 'xlsx'
import { UploadOutlined } from '@ant-design/icons-vue'
import dayjs from 'dayjs'

const fileList = ref([])
const methodChecked = ref(1)
const colsName = ref('')

// 全局去重
function handleDuplicates (excelData) {
  try {
    console.log(excelData, 'excelData')
    // 创建一个 Map 用于存储数据
    const dataMap = new Map()

    // 遍历 Excel 数据
    for (const row of excelData) {
        // 遍历每一行的属性
        for (const key in row) {
            if (row.hasOwnProperty(key)) {
                const value = row[key]
                if (!dataMap.has(key)) {
                    dataMap.set(key, [value])
                } else if (!dataMap.get(key).includes(value)) {
                    dataMap.get(key).push(value)
                }
            }
        }
    }

    // 将 Map 转换为所需的数组格式
    const result = Array.from(dataMap, ([key, values]) => {
        const dataObj = { [key]: values }
        return dataObj
    })

    console.log(result, 'init result')

    const map = {}
    const lenMap = {}
    const minLenMap = {}

    result.forEach(ele => {
        const key = Object.keys(ele)[0]
        lenMap[key] = ele[key].length
        ele[key].forEach(item => {
            if (map[item]) {
                map[item] += 1
            } else {
                map[item] = 1
            }
        })
    })

    console.log(map, '各数据出现次数Map')
    const newMap = {}
    Object.keys(map).forEach(key => {
        const value = map[key]
        if (value >= 3) {
            newMap[key] = value
        }
    })

    result.forEach(ele => {
        const key = Object.keys(ele)[0]
        ele[key].forEach(item => {
            if (minLenMap[item]) {
              minLenMap[item].push(key)
            } else {
              minLenMap[item] = [key]
            }
        })
    })

    console.log(lenMap, '各列长度Map')

    console.log(newMap, '大于3次出现次数Map')

    console.log(minLenMap, '各数据出现Map')

    if (methodChecked.value === 2 && !Object.keys(lenMap).includes(colsName.value)) {
      alert('当前输入单列去重标题不存在')
      return
    }

    // 全局去重
    if (methodChecked.value === 1) {
      // const entries = Object.entries(lenMap)
      // const minValuePropertys = entries.sort((a, b) => a[1] - b[1]).slice(0, 2).map(ele => ele[0])
      result.forEach(ele => {
          console.log(ele, 'result item')
          Object.keys(ele).forEach(key => {
              const value = ele[key]
              value.forEach((valueItem, i) => {
                  // console.log(valueItem, `第${i + 1}列数据`)
                  if (Object.keys(newMap).includes(String(valueItem))) {
                      const minValuePropertys = minLenMap[valueItem].sort((a, b) => lenMap[a] - lenMap[b]).slice(0, 2)
                      if (!minValuePropertys.includes(key)) {
                          // 如果有两列最小数据长度不包含当前列，则剔除当前属性
                          console.log(valueItem, key, minValuePropertys, 'duplicate item')
                          delete ele[key][i]
                      }
                  }
              })
              ele[key] = ele[key].filter(item => item)
          })
      })
      console.log(result, 'before result')
      const res = [] // 结果数组
      result.forEach((ele, i) => {
          const key = Object.keys(ele)[0]
          ele[key].forEach((item, itemIndex) => {
              if (!res[itemIndex]) res[itemIndex] = {}
              res[itemIndex][key] = item
          })
      })
      console.log(res, 'result')
      exportExcel(res)
    }

    // 单列去重
    if (methodChecked.value === 2) {
      const targetArr = result.find(ele => {
        return ele.hasOwnProperty(colsName.value)
      })[colsName.value]
      console.log(targetArr, 'targetArr')
      result.forEach(ele => {
          console.log(ele, 'result item')
          Object.keys(ele).forEach(key => {
              const value = ele[key]
              value.forEach((valueItem, i) => {
                  // console.log(valueItem, `第${i + 1}列数据`)
                  if (targetArr.includes(valueItem)) {
                      if (key !== colsName.value) {
                          // 如果有两列最小数据长度不包含当前列，则剔除当前属性
                          console.log(valueItem, key, 'duplicate item')
                          delete ele[key][i]
                      }
                  }
              })
              ele[key] = ele[key].filter(item => item)
          })
      })
    }
    let res = []
    result.forEach((ele, i) => {
          const key = Object.keys(ele)[0]
          ele[key].forEach((item, itemIndex) => {
              if (!res[itemIndex]) res[itemIndex] = {}
              res[itemIndex][key] = item
          })
      })
      console.log(res, 'result')
      exportExcel(res)
  } catch (error) {
    console.error(error)
  }
}

// 导出表格
function exportExcel (res) {
  let name = fileList.value[0].name
  name = name.replace('.xlsx', '')
  const currentDate = dayjs().format('YYYY-MM-DD');
  const defaultFileName = `${name}_${currentDate}_已去重.xlsx`

  // 创建一个工作簿对象
  const wb = XLSX.utils.book_new()

  // 创建一个工作表对象
  const ws = XLSX.utils.json_to_sheet(res)

  // 将工作表添加到工作簿
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1')

  // 导出
  XLSX.writeFile(wb, defaultFileName)

  fileList.value = []; // 清空文件列表
}

function beforeUpload(file, _fileList) {
  fileList.value = [file];
  return false
}

function handleFileUpload () {
  const file = fileList.value[0];
  if (file) {
    // 处理文件逻辑
    readFile(file.originFileObj)
  }
}

// 读取文件
function readFile (file) {
  const reader = new FileReader()

  reader.onload = (e) => {
      const fileContent = e.target.result
      processExcelData(fileContent)
  }

  reader.readAsArrayBuffer(file)
}
// 处理文件
function processExcelData (fileContent) {
  const data = new Uint8Array(fileContent)
  const workbook = XLSX.read(data, { type: 'array' })

  const firstSheetName = workbook.SheetNames[0]
  const worksheet = workbook.Sheets[firstSheetName]

  const excelData = XLSX.utils.sheet_to_json(worksheet)
  console.log(excelData)
  console.log(JSON.stringify(excelData))

  handleDuplicates(excelData) 
}

// 处理文件(单列)
</script>

<style scoped>
.logo {
  height: 6em;
  padding: 1.5em;
  will-change: filter;
  transition: filter 300ms;
}

.logo:hover {
  filter: drop-shadow(0 0 2em #646cffaa);
}

.logo.vue:hover {
  filter: drop-shadow(0 0 2em #42b883aa);
}

.app-radio {
  margin: 24px 0;
}

.app-cols-input {
  margin: 24px 0;
  font-size: 14px;
  > span {
    margin: 0 12px 0 0;
  }
}

.app-tips {
  margin: 24px 0;
  line-height: 30px;
  text-align: left;
}
</style>
