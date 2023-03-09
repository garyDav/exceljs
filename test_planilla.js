import ExcelJS from 'exceljs'
// import users from './users.json' assert { type: 'json' }
// import estudiantes from './estudiantes.json' assert { type: 'json' }

const createPlanilla = async () => {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile('planillaP1.xlsx')
  const worksheet = workbook.getWorksheet('notas')

  workbook.properties.date1904 = true
  /*const logo = workbook.addImage({
    filename: './logo.png',
    extension: 'png',
  })*/

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
  /*// Add Image
  worksheet.addImage(logo, {
    tl: { col: 0.6, row: 2.1 },
    ext: { width: 65, height: 78 },
  })

  // Merge Cells horizontal
  worksheet.mergeCells('A1:L2') // Título
  worksheet.mergeCells('A3:B3') // Unidad Educativa
  worksheet.mergeCells('C3:L3')
  worksheet.mergeCells('A4:B4') // Año de escolaridad
  worksheet.mergeCells('C4:L4')
  worksheet.mergeCells('A5:B5') // Maestro(a)
  worksheet.mergeCells('C5:L5')
  worksheet.mergeCells('A6:B6') // Área
  worksheet.mergeCells('C6:L6')
  worksheet.mergeCells('A7:B7') // Materia
  worksheet.mergeCells('C7:L7')
  worksheet.mergeCells('A8:B8') // Gestión
  worksheet.mergeCells('C8:L8')

  worksheet.mergeCells('C9:K9') // Evaluación Maestra(o)
  worksheet.mergeCells('C10:K10') // Dimensiones
  worksheet.mergeCells('C11:G11') // Saber 35pt
  worksheet.mergeCells('H11:K11') // Hacer 35pt

  // Merge Cells vertical
  worksheet.mergeCells('A9:A13') // Número de Lista
  worksheet.mergeCells('B9:B13') // Apellidos y nombres
  worksheet.mergeCells('L9:L13') // Total trimestral

  worksheet.getCell('C12').value = {
    richText: [
      {
        font: {
          bold: true,
        },
        text: 'VAR. EVAL.',
      },
    ],
  }*/

  /*worksheet.getCell('C12').fill = {
    type: 'gradient',
    gradient: 'angle',
    degree: 90,
  }*/

  /*worksheet.getCell('C2').alignment = {
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
  }*/

  await workbook.xlsx.writeFile('planilla.xlsx')

  console.log('File created')
  // console.log(workbook._worksheets[1]._columns[2]._worksheet._rows[11])
}

createPlanilla()
