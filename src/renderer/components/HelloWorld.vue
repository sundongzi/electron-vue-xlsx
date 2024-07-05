<script setup>
import { ref } from 'vue';
import Excel from 'exceljs'
import fileSaver from 'file-saver'

const CELL_STYLE = {
  "font": {
      "size": 10,
      "color": {
          "theme": 1
      },
      "name": "宋体",
      "family": 3,
      "charset": 134,
      "scheme": "minor"
  },
  "border": {
      "left": {
          "style": "thin"
      },
      "right": {
          "style": "thin"
      },
      "top": {
          "style": "thin"
      },
      "bottom": {
          "style": "thin"
      }
  },
  "fill": {
      "type": "pattern",
      "pattern": "none"
  },
  "alignment": {
      "horizontal": "center",
      "vertical": "middle",
      "wrapText": true
  }
}

// 通过index下标映射英文字母
const excelColumnName = (index) => {
    let columnName = '';
    while (index > 0) {
        let modulo = (index - 1) % 26;
        columnName = String.fromCharCode(65 + modulo) + columnName;
        index = (index - modulo) / 26 | 0;
    }
    return columnName || undefined;
}

/**
 * 读取信息表：获取项目名称、项目编号、包号、设备名称、投标公司名称
 */
const defaultInfo = {
  // 项目名称
  projectName: '',
  // 项目编号(拼接包号)
  projectCode: '',
  // 包名以及关联的公司名称
  packageData: {}
}
let info = ref({
  ...defaultInfo
})
const uploadFile = async (e) => {
  // 每次重新上传文件时重置
  info = ref({
    ...defaultInfo
  })
  if (!e.target.files[0]) return
  const data = await e.target.files[0].arrayBuffer()
  const workbook = new Excel.Workbook()
  await workbook.xlsx.load(data)
  const worksheet = workbook.getWorksheet('投标报名信息')
  worksheet?.eachRow({ includeEmpty: true }, function(row, rowNumber) {
    if (rowNumber === 1) {
      const { projectName, projectCode } = getProjectInfo(row.values[1])
      info.value.projectName = projectName
      info.value.projectCode = projectCode
    }
    
    
    if (rowNumber > 2) {
      const packageNumber = row.values[1];
      if (!info.value.packageData[packageNumber]) {
        info.value.packageData[packageNumber] = {
          deviceName: row.values[2],
          corporateNameList: [row.values[4]]
        }
      } else {
        info.value.packageData[row.values[1]].corporateNameList.push(row.values[4])
      }
      
    }
  });
  if (templateFileData.value) {
    uploadScoringFile()
  }
}
const getProjectInfo = (value) => {
  const nameRegex = /.*(?=（[^（]*$)/
  const codeRegex = /（([^（）]*)）[^（）]*$/
  const projectName = value.match(nameRegex)[0]
  const projectCode = value.match(codeRegex)[1]
  return {
    projectName,
    projectCode
  }
}
// 读取评分表模板
const templateFileData = ref(null)
const uploadScoringFile = async (e) => {
  if (!templateFileData.value) {
    templateFileData.value = await e.target.files[0].arrayBuffer()
  }
  const workbook = new Excel.Workbook()
  await workbook.xlsx.load(templateFileData.value)
  const worksheet = workbook.getWorksheet('模板')
  for (let [key, value] of Object.entries(info.value.packageData)) {
    const newWorksheet = workbook.addWorksheet(key);
    worksheet?.eachRow({ includeEmpty: true }, function(row, rowNumber) {
      // 先复制模板
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          const newCell = newWorksheet.getCell(rowNumber, colNumber);
          newCell.value = cell.value;
          if (cell.style) {
              newCell.style = { ...cell.style };
          }
          if ([2, 3].includes(colNumber)) {
            newCell._column.width = 22
          }
      });

      newWorksheet?.eachRow({ includeEmpty: true }, function(newRow, newRowNumber) {
        const { number } = newWorksheet.lastRow 
        // 最后两行不需要处理
        if ([number, number - 1].includes(rowNumber)) {
          return
        }
        if (newRowNumber === 2) {
          newRow.height = 30
          newRow.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true  }
          const cell2 = newRow.getCell(1)
          cell2.value = `项目名称：${info.value.projectName}\n项目编号：${info.value.projectCode}-${key}`
        }
        handleRow({ 
          worksheet: newWorksheet, 
          row: newRow, 
          rowIndex: newRowNumber, 
          deviceName: value.deviceName, 
          companyList: value.corporateNameList
        })
      })
    });
  }
  
  workbook.removeWorksheet('模板')
  // 文件下载
  downloadFile(workbook)
  
}
// 处理行以及单元格
const handleRow = ({ worksheet, row, rowIndex = 1, companyList, deviceName } = {} ) => {
  // 设置设备名称
  if (rowIndex === 1) {
    row.getCell(1).value = `${deviceName} 技术评分表`
  }
  // 第三（序号）行追加的列需要设置内容
  const isNumberRow = rowIndex === 3
  // 第十四行（合计）行
  const isTotalRow = rowIndex === 14
  // 求和开始列
  const START_SUM_ROW = 4
  // 获取当前列数
  let columnCount = row.cellCount;
  const lastCell = row.getCell(columnCount);
  companyList?.forEach((item, index) => {
    let newCell = row.getCell(columnCount + 1 + index);
    if (isTotalRow) {
      const columnName = newCell.address[0]
      newCell.value = {
        formula: `SUM(${columnName}${START_SUM_ROW}:${columnName}${rowIndex - 1})`
      }
      worksheet.unMergeCells('A14')
      worksheet.mergeCells('A14:B14')
      worksheet.unMergeCells('C14')
      worksheet.mergeCells('C14:D14')
    } else {
      if (isNumberRow) {
        newCell.value = item
      } else {
        newCell.value = null
        const grade = typeof lastCell.value === 'number' ?  lastCell.value : lastCell.value?.split('-')[1]
        // 单元格分数校验
        newCell.dataValidation = {
          allowBlank: true,
          formulae: [Number(grade)],
          operator: "lessThanOrEqual",
          showErrorMessage: true,
          showInputMessage: true,
          type: 'whole'
        }
      }
    }
    // 设置单元格样式
    newCell.style = CELL_STYLE
    // // 设置单元格宽度
    newCell._column.width = 20

    if (rowIndex === 5) {
      worksheet.unMergeCells('A4')
      worksheet.mergeCells('A4:A5')
      worksheet.unMergeCells('B4')
      worksheet.mergeCells('B4:B5')
      worksheet.mergeCells(rowIndex, columnCount + 1 + index , rowIndex - 1, columnCount + 1 + index)
    }
  })
  if (rowIndex === 3) {
    worksheet.mergeCells(rowIndex, 3 ,rowIndex, 2)
  }
  // 处理第一行与第二行需要所有单元格合并
  const isMergeRow = [1, 2].includes(rowIndex)
  if (!isMergeRow) {
    return
  }
  // 只有第一行与第二行需要整体单元格合并
  const newColumnCount = row.cellCount
  worksheet.unMergeCells(`A${rowIndex}`)
  worksheet.mergeCells(rowIndex, newColumnCount ,rowIndex, 1)
}
// 文件导出
const downloadFile = async (workbook) => {
  const buffer = await workbook.xlsx.writeBuffer()
  fileSaver(new Blob([buffer], {
    type: 'application/octet-stream'
  }), '评分表.xlsx')
}
</script>

<template>
<div>
  <div>
    <label class="label">
      请选择信息表：
    </label>
    <input type="file" @change="uploadFile">
  </div>
  <div>
    <label class="label">
      请选择技术评分表：
    </label>
    <input type="file" @change="uploadScoringFile">
  </div>
</div>
</template>
<style scoped>
.label {
  display: inline-block;
  width: 160px;
  text-align:right;
  margin-bottom: 8px;
}
</style>