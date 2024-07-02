<script setup>
import { ref } from 'vue';
import Excel from 'exceljs'
import fileSaver from 'file-saver'

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

const selectedFile = ref(null);

  const handleFileChange = (file) => {
    selectedFile.value = file.raw;
  };

  const handleUpload = () => {
    const reader = new FileReader();
    reader.onload = (event) => {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      worksheet?.eachRow({ includeEmpty: true }, function(row, rowNumber) {
      console.log('Row ' + rowNumber + ' = ' + JSON.stringify(row.values));
        row.eachCell({ includeEmpty: true }, function(cell, colNumber) {
          console.log('Cell ' + colNumber + ' = ' + cell.value);
        });
      });
      // 将工作表内容转换为JSON
      // const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      // // 删除第一列
      // const modifiedData = jsonData.map(row => row.slice(1));

      // // 创建一个新的工作簿和工作表
      // const newWorkbook = XLSX.utils.book_new();
      // const newWorksheet = XLSX.utils.aoa_to_sheet(modifiedData);
      // XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Sheet1');

      // // 生成新的.xlsx文件
      // const newExcelBuffer = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' });
      // const blob = new Blob([newExcelBuffer], { type: 'application/octet-stream' });
      // const link = document.createElement('a');
      // link.href = URL.createObjectURL(blob);
      // link.download = 'modified.xlsx';
      // link.click();
    };
    reader.readAsArrayBuffer(selectedFile.value);
  }
  /**
   * 读取信息表：获取项目名称、项目编号、包号、设备名称、投标公司名称
   */
  const uploadFile = async (e) => {
    const data = await e.target.files[0].arrayBuffer()
    console.log('eeeee', data)
    const workbook = new Excel.Workbook()
    console.log('workbook', workbook)
    await workbook.xlsx.load(data)
    const worksheet = workbook.getWorksheet('投标报名信息')
    console.log('worksheetworksheet', worksheet.getRow(1))
    worksheet?.eachRow({ includeEmpty: true }, function(row, rowNumber) {
    console.log('Row ' + rowNumber + ' = ' + JSON.stringify(row.values), row);
      row.eachCell({ includeEmpty: true }, function(cell, colNumber) {
        console.log('Cell ' + colNumber + ' = ' + cell.value, cell);
      });
    });
  }
  // 读取评分表
  const uploadScoringFile = async (e) => {
    const data = await e.target.files[0].arrayBuffer()
    const workbook = new Excel.Workbook()
    await workbook.xlsx.load(data)
    const worksheet = workbook.getWorksheet('1')
    worksheet?.eachRow({ includeEmpty: true }, function(row, rowNumber) {
      const { number } = worksheet.lastRow 
      // 最后两行不需要处理
      if ([number, number - 1].includes(rowNumber)) {
        return
       }
      handleRow(worksheet, row, rowNumber)

      console.log('Row ' + rowNumber + ' = ' + JSON.stringify(row.values), row);
      row.eachCell({ includeEmpty: true }, function(cell, colNumber) {
        console.log('Cell ' + colNumber + ' = ' + cell.value, cell);
      });
    });
    const buffer = await workbook.xlsx.writeBuffer()
    
    fileSaver(new Blob([buffer], {
      type: 'application/octet-stream'
    }), 'aaa.xlsx')
    
  }
  // 处理第一行与第二行
  const handleRow = (worksheet, row, rowIndex = 1) => {
    // 处理第一行与第二行需要所有单元格合并
    const isMergeRow = [1, 2].includes(rowIndex)
    // 第三（序号）行追加的列需要设置内容
    const isNumberRow = rowIndex === 3
    // 第十四行（合计）行
    const isTotalRow = rowIndex === 14
    // 求和开始列
    const START_SUM_ROW = 4
    // 获取当前列数
    let columnCount = row.cellCount;
    [1,2,3,4,5].forEach((item, index) => {
      let newCell = row.getCell(columnCount + 1 + index);
      if (isTotalRow) {
        const lastCell = row.getCell(columnCount)
        console.log('lastCell', lastCell, newCell)
        // newCell.formula = `SUM()`
        const columnName = newCell.address[0]
        newCell.value = {
          formula: `SUM(${columnName}${START_SUM_ROW}:${columnName}${rowIndex - 1})`,
        }
        // 求和校验，不超过100
        newCell.dataValidation = {
          allowBlank: true,
          formulae: [6],
          operator: "lessThanOrEqual",
          showErrorMessage: true,
          showInputMessage: true,
          type: 'whole'
        }
      } else {
        newCell.value = isNumberRow ? item : ''
      }
      newCell.alignment = {
        horizontal: 'center',
        vertical: 'middle'
      }
      
    })
    if (!isMergeRow) {
      return
    }
    // 只有第一行与第二行需要整体单元格合并
    const newColumnCount = row.cellCount
    worksheet.unMergeCells(`A${rowIndex}`)
    worksheet.mergeCells(rowIndex, newColumnCount ,rowIndex, 1)
  }
  // 处理第三行
  const handleThreeRow = (worksheet) => {
    
  }
  const downloadFile = async () => {
    const workbook = new Excel.Workbook()
    const worksheet = workbook.addWorksheet('1')
    const buffer = await workbook.xlsx.writeBuffer()
    
    fileSaver(new Blob([buffer], {
      type: 'application/octet-stream'
    }), 'aaa.xlsx')
  }
  const submitUpload = () => {
    uploadRef.value.submit();
  };

  const uploadRef = ref(null);
</script>

<template>
<div>
  请选择信息表：<input type="file" @change="uploadFile">
  请选择技术评分表：<input type="file" @change="uploadScoringFile">
  <!-- <el-upload
    ref="uploadRef"
    action=""
    :http-request="handleUpload"
    :on-change="handleFileChange"
  >
    <el-button slot="trigger" size="small" type="primary">选取文件</el-button>
    <el-button
      style="margin-left: 10px;"
      type="success"
      @click="submitUpload"
    >
      上传并处理文件
    </el-button>
  </el-upload> -->
</div>
</template>
