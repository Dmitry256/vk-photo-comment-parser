import ExcelJS from 'exceljs'
import fs from 'fs'
import path from 'path'

async function processExcel(
  filePath: string,
  outputDir: string
): Promise<void> {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(filePath)

  const worksheet = workbook.worksheets[0]
  const groups: {[name: string]: ExcelJS.Row[]} = {}

  // Собираем заголовки
  const headers = worksheet.getRow(1).values as ExcelJS.CellValue[]

  // Группируем строки по имени
  worksheet.eachRow({includeEmpty: false}, (row, rowNumber) => {
    if (rowNumber === 1) return

    const name = row.getCell(2).text.trim()
    if (!name) return

    if (!groups[name]) groups[name] = []
    groups[name].push(row)
  })

  // Создаем папку для результатов
  if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, {recursive: true})

  // Создаем файлы для каждой группы
  for (const [name, rows] of Object.entries(groups)) {
    const newWorkbook = new ExcelJS.Workbook()
    const newSheet = newWorkbook.addWorksheet('Данные')

    // Добавляем заголовки
    newSheet.columns = [
      {width: 5},
      {width: 21},
      {width: 30},
      {width: 8},
      {width: 8},
    ]
    newSheet.addRow(headers).font = {bold: true}

    // Добавляем строки и считаем сумму
    let total = 0
    rows.forEach((row) => {
      const values = row.values
      const currentRow = newSheet.addRow(values)
      currentRow.getCell('E').value = values[5].result // Костыль, нужно переделать, чтобы

      // Суммируем столбец E (+20%)
      const value = row.getCell(5).result
      if (typeof value === 'number') total += value
    })

    // const formula = {formula: 'СУММ(СМЕЩ(E1;;;СТРОКА()-1;1))'}
    const formula = {formula: 'SUM(OFFSET(E1,0,0,ROW()-1,1))'}

    // Добавляем итоговую строку
    newSheet.addRow(['', '', '', 'Итого:', formula, total]).font = {bold: true}

    // Сохраняем файл
    const safeName = name.replace(/[\\/*?:[\]]/g, '_')
    await newWorkbook.xlsx.writeFile(path.join(outputDir, `${safeName}.xlsx`))
  }
}

// Запуск обработки
processExcel('output/test.xls', 'output/files')
  .then(() => console.log('Файлы успешно созданы!'))
  .catch(console.error)
