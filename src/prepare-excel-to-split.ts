import ExcelJS from 'exceljs'
import {UserComment, userCommentSchema} from './schemas'

const backupFirstWorksheet = (workbook: ExcelJS.Workbook): void => {
  const hasBackup = workbook.worksheets.some((sheet) =>
    sheet.name.startsWith('Backup_')
  )
  if (hasBackup) {
    console.info('Backup already exists')
    return
  }
  const originalSheet = workbook.worksheets[0]

  const originalSheetModel = originalSheet.model

  const backupSheet = workbook.addWorksheet()

  backupSheet.model = {
    ...originalSheetModel,
    name: `Backup_${originalSheet.name}`,
  }
}

const remImgFirstWorksheet = (workbook: ExcelJS.Workbook): void => {
  const originalSheet = workbook.worksheets[0]

  const originalSheetModel = originalSheet.model

  originalSheet.model = {
    ...originalSheetModel,
    media: [], // Очищаем изображения
  }
}

const prepareToSplit = async (filePath: string): Promise<void> => {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(filePath)

  // Делаем копию первого листа
  backupFirstWorksheet(workbook)

  // Удаляем media (img) из рабочего (первого) листа
  remImgFirstWorksheet(workbook)

  // Удаляем столбцы EFG (фото и описание)
  const worksheet = workbook.worksheets[0] // TODO use getWorksheet(MAIN_WORKSHEET_NAME)
  worksheet.spliceColumns(5, 3)

  // Sort by name and by photo number

  const rows = worksheet.getRows(2, worksheet.rowCount)

  let comments: UserComment[] = []
  rows?.forEach((row) => {
    if (Array.isArray(row.values)) {
      const [, hyperlink, userName, text, price, numberInAlbum] = row.values // Пропускаем 0-ой элемент (он везде undefined)
      const comment = userCommentSchema.parse({
        hyperlink,
        userName,
        text,
        price,
        numberInAlbum,
      })
      comments.push(comment)
    }
  })

  comments = comments
    .filter((comment: UserComment) => {
      return (
        comment.userName !== undefined && comment.numberInAlbum !== undefined
      )
    })
    .sort((a, b) => {
      const compareByName = String(a.userName).localeCompare(String(b.userName))
      if (compareByName !== 0) return compareByName
      return (a.numberInAlbum ?? 0) - (b.numberInAlbum ?? 0)
    })

  console.log('comments :', comments)

  for (let i = worksheet.rowCount; i >= 2; i--) {
    worksheet.spliceRows(i, 1)
  }

  // 4. Записываем отсортированные данные

  let userCommentCount = 0
  comments.forEach((comment: UserComment, index: number, comments) => {
    const currentRow = worksheet.addRow([...Object.values(comment)])
    userCommentCount += 1
    if (comments[index + 1]?.userName !== comment.userName) {
      worksheet.addRow([
        '',
        '',
        '',
        '',
        '',
        {
          formula: `SUM(E${currentRow.number - userCommentCount}:E${currentRow.number}`,
        },
      ])
      userCommentCount = 0
    }
  })

  // 5. Заменяем колонку с номером фото на цену +20%

  worksheet.getColumn('E').eachCell((cell, rowNumber) => {
    if (rowNumber === 1) {
      cell.value = '+20%'
    } else {
      cell.value = {formula: `D${rowNumber}*1.2`}
    }
  })

  // Сохраняем файл
  await workbook.xlsx.writeFile(filePath)
}

// Запуск обработки
prepareToSplit('output/test.xls')
  .then(() => console.log('Файлы успешно созданы!'))
  .catch(console.error)
