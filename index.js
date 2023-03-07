import ExcelJS from 'exceljs'
import users from './users.json' assert { type: 'json' }
// await import('./users.json', { assert: { type: 'json' } })

const createExcel = async data => {
  // Crear un archivo en excel
  const workbook = new ExcelJS.Workbook()
  // Crear una hoja dentro del archivo excel
  const worksheet = workbook.addWorksheet('Users')

  // Definir las columnas de la hora del excel
  worksheet.columns = [
    { header: 'Id', key: 'id', width: 10 },
    { header: 'First name', key: 'first_name', width: 20 },
    { header: 'Last name', key: 'last_name', width: 20 },
  ]

  // Agregar la data a la hoja del excel
  worksheet.addRows(data)

  // Guardar el excel
  await workbook.xlsx.writeFile('users.xlsx')

  console.log('File created')

  // Cargar un excel
  const newWorkbook = new ExcelJS.Workbook()
  await newWorkbook.xlsx.readFile('users.xlsx')

  // Obtener una hoja del excel cargado
  const newWorksheet = newWorkbook.getWorksheet('Users')
  // Definir las columnas del nuevo archivo
  newWorksheet.columns = [
    { header: 'Id', key: 'id', width: 10 },
    { header: 'First name', key: 'first_name', width: 20 },
    { header: 'Last name', key: 'last_name', width: 20 },
  ]

  // Agregar un nuevo row al excel
  newWorksheet.addRow({ id: 3, first_name: 'user 3', last_name: 'last name 3' })

  await newWorkbook.xlsx.writeFile('users2.xlsx')

  console.log('File created')
}

createExcel(users)
