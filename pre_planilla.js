import ExcelJS from 'exceljs'
// import users from './users.json' assert { type: 'json' }

const createPlanilla = async () => {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile('planillaP1.xlsx')
  const worksheet = workbook.getWorksheet('notas')

  workbook.properties.date1904 = true

  await workbook.xlsx.writeFile('planilla.xlsx')

  console.log('File created')
}

createPlanilla()
