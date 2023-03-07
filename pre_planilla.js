import ExcelJS from 'exceljs'
import users from './users.json' assert { type: 'json' }
// import estudiantes from './estudiantes.json' assert { type: 'json' }

const createPlanilla = async data => {
  const workbook = new ExcelJS.Workbook()
  const worksheet = workbook.addWorksheet('notas', {
    properties: { tabColor: { argb: 'FFC0000' } },
    views: [{ showGridLines: false }],
  })

  workbook.properties.date1904 = true

  worksheet.columns = [
    { header: 'Id', key: 'id', width: 10 },
    { header: 'First name', key: 'first_name', width: 20 },
    { header: 'Last name', key: 'last_name', width: 20 },
  ]
  worksheet.addRows(data)

  worksheet.autoFilter = {
    from: 'A1',
    to: 'C1',
  }

  await workbook.xlsx.writeFile('planilla.xlsx')

  console.log('File created')
}

createPlanilla(users)
