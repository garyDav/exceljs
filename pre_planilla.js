import ExcelJS from 'exceljs'
import users from './users.json' assert { type: 'json' }
// import estudiantes from './estudiantes.json' assert { type: 'json' }

const createPlanilla = async data => {
  const workbook = new ExcelJS.Workbook()
  const worksheet = workbook.addWorksheet('notas', {
    headerFooter: { firstHeader: 'Hello Exceljs', firstFooter: 'Hello World' },
  })

  workbook.properties.date1904 = true
  const logo = workbook.addImage({
    filename: './logo.png',
    extension: 'png',
  })

  // worksheet.addBackgroundImage(logo)
  // worksheet.addImage(logo, 'E1:E3')
  // worksheet.getCell('B2').alignment = { textRotation: 'vertical' }
  /*worksheet.columns = [
    { header: 'Id', key: 'id', width: 10 },
    { header: 'First name', key: 'first_name', width: 20 },
    { header: 'Last name', key: 'last_name', width: 20 },
  ]
  worksheet.addRows(data)

  worksheet.autoFilter = {
    from: 'A1',
    to: 'C1',
  }*/
  worksheet.addImage(logo, {
    tl: { col: 1.2, row: 0.5 },
    ext: { width: 63, height: 75 },
  })

  worksheet.mergeCells('C2:J3')
  worksheet.getCell('C2').alignment = {
    vertical: 'middle',
    horizontal: 'center',
  }
  worksheet.getCell('B6').alignment = { horizontal: 'right' }
  worksheet.getCell('B7').alignment = { horizontal: 'right' }
  worksheet.getCell('B8').alignment = { horizontal: 'right' }
  worksheet.getCell('G6').alignment = { horizontal: 'right' }
  worksheet.getCell('G7').alignment = { horizontal: 'right' }
  worksheet.getCell('G8').alignment = { horizontal: 'right' }
  worksheet.getCell('C2').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' },
  }
  worksheet.getRow(11).height = 15
  worksheet.getColumn('A').width = 5
  worksheet.getColumn('B').width = 25

  worksheet.getCell('C2').value = {
    richText: [
      {
        font: {
          bold: true,
          size: 20,
          color: { theme: 1 },
          name: 'Calibri',
          family: 2,
          scheme: 'minor',
        },
        text: 'PLANILLA PRIMER TRIMESTRE (PARTE 1)',
      },
    ],
  }
  worksheet.getCell('B6').value = {
    richText: [
      {
        font: {
          bold: true,
        },
        text: 'UNIDAD EDUCATIVA: ',
      },
    ],
  }
  worksheet.getCell('B7').value = {
    richText: [
      {
        font: {
          bold: true,
        },
        text: 'AÑO DE ESCOLARIDAD: ',
      },
    ],
  }
  worksheet.getCell('B8').value = {
    richText: [
      {
        font: {
          bold: true,
        },
        text: 'MAESTRO (A): ',
      },
    ],
  }
  worksheet.getCell('G6').value = {
    richText: [
      {
        font: {
          bold: true,
        },
        text: 'ÁREA: ',
      },
    ],
  }
  worksheet.getCell('G7').value = {
    richText: [
      {
        font: {
          bold: true,
        },
        text: 'MATERIA: ',
      },
    ],
  }
  worksheet.getCell('G8').value = {
    richText: [
      {
        font: {
          bold: true,
        },
        text: 'GESTIÓN: ',
      },
    ],
  }

  await workbook.xlsx.writeFile('planilla.xlsx')

  console.log('File created', data)
}

createPlanilla(users)
