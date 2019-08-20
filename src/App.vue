<template>
  <div id="app">
    <audio
      :src="error"
      ref="error"
    ></audio>
    <audio
      :src="success"
      ref="successMP3"
    ></audio>
    <div class="file-content">
      <input
        type="file"
        id="excel-file"
        accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      />
    </div>
    <el-card
      shadow="never"
      style="margin: 10px 0;"
    >
      <el-form
        ref="form"
        :model="form"
        :inline="true"
        label-width="80px"
      >
        <el-form-item label="运单号">
          <el-input
            v-model="form.num"
            @keyup.enter.native="onSubmit"
          ></el-input>
          <el-input style="display: none;"></el-input>
        </el-form-item>
        <el-form-item>
          <el-button
            type="primary"
            @click="onSubmit"
          >查询</el-button>
          <el-button
            type="primary"
            @click="exportExcel"
          >导出</el-button>
        </el-form-item>
        <el-form-item>
          {{ tip }}
        </el-form-item>
      </el-form>
    </el-card>
    <el-row>
      <el-col :span="11">
        <el-table
          :data="tableData"
          height="450"
          border
          style="width: 100%"
        >
          <el-table-column
            prop="客户"
            label="客户"
          >
          </el-table-column>
          <el-table-column
            prop="批次号"
            label="批次号"
          >
          </el-table-column>
          <el-table-column
            prop="运单号"
            label="运单号"
          >
          </el-table-column>
          <el-table-column
            prop="品名"
            label="品名"
          >
          </el-table-column>
          <el-table-column
            prop="箱规"
            label="箱规"
          >
          </el-table-column>
          <el-table-column
            prop="备注"
            label="备注"
          >
          </el-table-column>
        </el-table>
      </el-col>
      <el-col :span="2"></el-col>
      <el-col
        :span="11"
        id="exportExcel"
      >
        <el-table
          :data="selectData"
          style="width: 100%"
          height="450"
        >
          <el-table-column label="扫码数据统计">
            <el-table-column :label="`导入数据：${tableLengthData.tableLength}票;实际扫码：${tableLengthData.actual}票;数据内：${tableLengthData.inData};不在数据内：${tableLengthData.noInData};重复：${tableLengthData.repeat}`">
              <el-table-column label="未扫码运单号">
                <el-table-column
                  prop="客户"
                  label="客户"
                >
                </el-table-column>
                <el-table-column
                  prop="批次号"
                  label="批次号"
                >
                </el-table-column>
                <el-table-column
                  prop="运单号"
                  label="运单号"
                >
                </el-table-column>
                <el-table-column
                  prop="品名"
                  label="品名"
                >
                </el-table-column>
                <el-table-column
                  prop="箱规"
                  label="箱规"
                >
                </el-table-column>
                <el-table-column
                  prop="备注"
                  label="备注"
                >
                </el-table-column>
              </el-table-column>
            </el-table-column>
          </el-table-column>
        </el-table>
        <el-table
          :data="noSelectData"
          style="width: 100%"
          height="450"
        >
          <el-table-column label="不在数据内分运单">
            <el-table-column
              prop="客户"
              label="客户"
            >
            </el-table-column>
            <el-table-column
              prop="批次号"
              label="批次号"
            >
            </el-table-column>
            <el-table-column
              prop="运单号"
              label="运单号"
            >
            </el-table-column>
            <el-table-column
              prop="品名"
              label="品名"
            >
            </el-table-column>
            <el-table-column
              prop="箱规"
              label="箱规"
            >
            </el-table-column>
            <el-table-column
              prop="备注"
              label="备注"
            >
            </el-table-column>
          </el-table-column>
        </el-table>
      </el-col>
    </el-row>
  </div>
</template>

<script>
import { table2excel } from '@/utils/exportExcel'
export default {
  name: 'App',
  data () {
    return {
      error: require('./assets/error.mp3'),
      success: require('./assets/success.mp3'),
      tableData: [],
      selectData: [],
      noSelectData: [],
      searchData: [],
      tableLengthData: {
        tableLength: 0, // 总票数
        actual: 0, // 实际扫码
        inData: 0, // 数据内
        noInData: 0, // 不在数据内
        repeat: 0 // 重复
      },
      form: {
        num: ''
      },
      tip: ''
    }
  },
  mounted () {
    this.excelFileChange()
  },
  methods: {
    exportExcel () {
      if (!this.tableData.length) return this.$message.error('请导入表格')
      this.$confirm('是否确认导出?', '提示', {
        confirmButtonText: '确定',
        cancelButtonText: '取消',
        type: 'warning'
      }).then(() => {
        table2excel('exportExcel')
      }).catch(() => {
      })
    },
    onSubmit () {
      if (!this.tableData.length) return this.$message.error('请导入表格')
      this.tableLengthData.actual = this.tableLengthData.actual + 1
      for (let i = 0; i < this.tableData.length; i++) {
        // 重复
        const index = this.searchData.findIndex(item => {
          return item['运单号'] === this.form.num
        })
        if (index !== -1) {
          this.$refs['error'].play()
          this.$message.warning('运单号重复')
          this.tableLengthData.repeat = this.tableLengthData.repeat + 1
          this.form.num = ''
          return
        }
        // 查找成功
        if (this.tableData[i]['运单号'] === this.form.num) {
          this.$refs['successMP3'].play()
          this.tableLengthData.inData = this.tableLengthData.inData + 1
          const index = this.selectData.findIndex(item => {
            return item['运单号'] === this.form.num
          })
          this.searchData.push(this.tableData[i])
          this.selectData.splice(index, 1)
          this.tip = this.tableData[i]
          this.$message.success('运单号查询成功')
          this.form.num = ''
          return
        }
      }
      // 查询不到
      this.$refs['error'].play()
      this.noSelectData.push({ '运单号': this.form.num })
      this.tableLengthData.noInData = this.tableLengthData.noInData + 1
      this.$message.warning('查询不到运单号')
      this.form.num = ''
    },
    excelFileChange () {
      document.getElementById('excel-file').onchange = (e) => {
        var files = e.target.files
        var fileReader = new FileReader()
        fileReader.onload = (ev) => {
          try {
            var data = ev.target.result
            // eslint-disable-next-line no-undef
            var workbook = XLSX.read(data, {
              type: 'binary'
            }) // 以二进制流方式读取得到整份excel表格对象
            var persons = [] // 存储获取到的数据
            console.log(workbook)
          } catch (e) {
            console.log('文件类型不正确')
            return
          }
          // 表格的表格范围，可用于判断表头是否数量是否正确
          var fromTo = ''
          // 遍历每张表读取
          for (var sheet in workbook.Sheets) {
            if (workbook.Sheets.hasOwnProperty(sheet)) {
              fromTo = workbook.Sheets[sheet]['!ref']
              console.log(fromTo)
              // eslint-disable-next-line no-undef
              persons = persons.concat(XLSX.utils.sheet_to_json(workbook.Sheets[sheet]))
              // break; // 如果只取第一张表，就取消注释这行
            }
          }
          const length = persons.length
          this.$set(this.tableLengthData, 'tableLength', length)
          this.$set(this, 'tableData', [...persons])
          this.$set(this, 'selectData', [...persons])
        }
        // 以二进制方式打开文件
        fileReader.readAsBinaryString(files[0])
      }
    }
  }
}
</script>

<style lang="scss">
#app {
  font-family: "Avenir", Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  color: #2c3e50;
}
#nav {
  padding: 30px;
  a {
    font-weight: bold;
    color: #2c3e50;
    &.router-link-exact-active {
      color: #42b983;
    }
  }
}
</style>
