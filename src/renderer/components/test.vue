<script setup lang="ts">
import Excel from 'exceljs'
import { saveAs } from 'file-saver'

const uploadFile = async (e:any) => {
  const data = e.target.files[0].arrayBuffer
  const workbook = new Excel.Workbook()
  await workbook.xlsx.load(data)
  const worksheet = workbook.getWorksheet(1)
  worksheet?.eachRow({ includeEmpty: true }, function(row, rowNumber) {
  console.log('Row ' + rowNumber + ' = ' + JSON.stringify(row.values));
    row.eachCell({ includeEmpty: true }, function(cell, colNumber) {
      console.log('Cell ' + colNumber + ' = ' + cell.value);
    });
  });
}

const downFile = () => {

}
</script>

<template>
 <input type="file" @change="uploadFile">
 <button @click="downFile"></button>
</template>
