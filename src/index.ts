import axios, {AxiosInstance} from 'axios'
import ExcelJS from 'exceljs'
import * as dotenv from 'dotenv'
import fs from 'fs'

// https://vk.com/album-225190306_307435564
// https://vk.com/album-225190306_307783090
// https://vk.com/album-225190306_307826349
// https://vk.com/album-225190306_307858232

const OWNER_ID = -225190306
const ALBUM_ID = 307858232
const IMAGES_FOLDER = '.temp/images/'

dotenv.config()

// настройка instance of axios
const instance = axios.create({
  baseURL: process.env.BASE_URL,
  headers: {
    Authorization: `Bearer ${process.env.ACCESS_TOKEN}`,
  },
  params: {
    v: '5.199',
  },
})

const getAllAlbumComments = async (axios: AxiosInstance) => {
  return await axios
    .get('execute', {
      params: {
        code: `
        var items = [];
        var currentResultLength = 100;
        var offset = 0;
        while (currentResultLength == 100) {
          var result = API.photos.getAllComments({
            owner_id: ${OWNER_ID},
            album_id: ${ALBUM_ID},
            count: 100,
            offset: offset
          });
          offset = offset + 100;
          currentResultLength = result.items.length;
          items = items + result.items;
        };
        return items;`,
      },
    })
    .then((response) => {
      if (response.data.error) throw Error(response.data.error.error_msg)
      return response.data.response
    })
}

const getUsersByIds = async (usersIds: number[], axios: AxiosInstance) => {
  return await axios
    .get('users.get', {
      params: {
        user_ids: usersIds.join(','),
      },
    })
    .then((response) => {
      return response.data.response
    })
    .catch((error) => {
      console.log('error :', error)
    })
}

const getPhotos = async (photosIds: string[], axios: AxiosInstance) => {
  const photos = photosIds.join(', ')
  return await axios
    .get('photos.getById', {
      params: {
        photos: photos,
      },
    })
    .then((response) => {
      return response.data.response
    })
    .catch((error) => {
      console.log('error :', error)
    })
}

const getAlbumPhotosOrder = async (axios: AxiosInstance) => {
  return await axios
    .get('photos.get', {
      params: {
        owner_id: OWNER_ID,
        album_id: ALBUM_ID,
        count: 500,
      },
    })
    .then((response) => {
      const albumPhotos = response.data.response.items
      const albumPhotosOrder = albumPhotos.map((photo) => {
        return photo.id
      })
      return albumPhotosOrder
    })
    .catch((error) => {
      console.log('error :', error)
    })
}

const calculateLines = (text: string, charsPerLine: number = 80): number => {
  return text
    .split('\n')
    .reduce((acc: number, currentLine: string) => {
      acc += Math.ceil(currentLine.length / charsPerLine)
      return acc
    }, 0)
}

const exportToExcel = async (
  comments: any[],
  photos: any[],
  filename: string
): Promise<void> => {
  const workbook = new ExcelJS.Workbook()
  const worksheet = workbook.addWorksheet('Comments')

  worksheet.columns = [
    {header: 'Фото', key: 'photo_url', width: 5},
    {
      header: 'Имя',
      key: 'user_name',
      width: 21,
      style: {
        alignment: {
          wrapText: true,
        },
      },
    },
    {
      header: 'Комментарий',
      key: 'text',
      width: 30,
      style: {
        alignment: {
          wrapText: true,
        },
      },
    },
    {header: 'Цена', key: 'price', width: 5},
    {
      header: `Описание`,
      width: 27,
      style: {
        alignment: {
          wrapText: true,
        },
      },
    },
    {width: 11},
    {width: 11},
  ]

  const albumPhotosOrder = await getAlbumPhotosOrder(instance)

  comments.forEach((comment, index, comments) => {
    if (index === 0 || comment.pid !== comments[index - 1]?.pid) {
      const row = worksheet.addRow({})
      const rowNumber = row.number
      const imageId = workbook.addImage({
        filename: `${IMAGES_FOLDER}${comment.pid}.jpg`,
        extension: 'jpeg',
      })
      worksheet.addImage(imageId, {
        tl: {col: 5, row: rowNumber - 1},
        ext: {width: 100, height: 100},
      })
      const imageRow = worksheet.getRow(rowNumber)
      const descriptionCell = worksheet.getCell(`E${rowNumber}`)
      const descriptionText = photos.find(
        (photo) => photo.id === comment.pid
      ).text
      descriptionCell.value = descriptionText
      imageRow.height = Math.max(77, calculateLines(descriptionText, 26) * 20)
    }

    const row = worksheet.addRow({
      photo_url: {
        text: albumPhotosOrder.indexOf(comment.pid) + 1,
        hyperlink: comment.photo_url,
      },
      user_name: comment.from
        ? `${comment.from.first_name} ${comment.from.last_name}`
        : 'Я',
      text: comment.text,
    })

    row.height = calculateLines(comment.text) * 17 // Определяем высоту строки по количеству текста комментария

    if (comment.attachments) {
      const commentPhotos = comment.attachments
        .filter((attachment) => attachment.type === 'photo')
        .map((attachment) => attachment.photo)

      for (let i = 0; i < commentPhotos.length; i++) {
        const photo = commentPhotos[i]
        const imageId = workbook.addImage({
          filename: `${IMAGES_FOLDER}${photo.id}.jpg`,
          extension: 'jpeg',
        })
        worksheet.addImage(imageId, {
          tl: {col: 4 + i, row: row.number - 1},
          ext: {width: 100, height: 100},
        })
      }
      row.height = 77
    }
  })

  await workbook.xlsx.writeFile(filename)
  console.log(`Exported to ${filename}`)
}

const downloadImage = async (imageUrl: string, fileName: string) => {
  const response = await axios.get(imageUrl, {responseType: 'arraybuffer'})
  fs.writeFileSync(fileName, Buffer.from(response.data, 'binary'))
}

const allComments: any[] = await getAllAlbumComments(instance) // Get all comments

// sort comments
allComments.reverse().sort((a, b) => a.pid - b.pid)

const photosIds = new Set<string>()
const usersIds = new Set<number>()
const photosFromComments = new Set<any>()
allComments.forEach((comment) => {
  usersIds.add(comment.from_id)
  photosIds.add(`${OWNER_ID}_${comment.pid}`)
  comment.attachments
    ?.filter((attachment) => attachment.type === 'photo')
    .forEach((attachment) => {
      photosFromComments.add(attachment.photo)
    })
})

const albumPhotos = await getPhotos(Array.from(photosIds), instance)
const photos = albumPhotos.concat(...photosFromComments)

for (const photo of photos) {
  if (!fs.existsSync(`${IMAGES_FOLDER}${photo}.jpg`)) {
    const pSizeUrl = photo.sizes.find((size) => size.type === 'm').url
    await downloadImage(pSizeUrl, `${IMAGES_FOLDER}${photo.id}.jpg`)
  }
}

const uniqUsers = await getUsersByIds(Array.from(usersIds), instance)

const commentsWithUsers = allComments.map((comment) => {
  const user = uniqUsers.find((user) => comment.from_id === user.id)
  comment.from = user
  comment.photo_url = `https://vk.com/photo${OWNER_ID}_${comment.pid}`
  return comment
})

exportToExcel(commentsWithUsers, photos, 'output/test.xls')
