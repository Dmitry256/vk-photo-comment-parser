import axios, {AxiosInstance} from 'axios'
import ExcelJS from 'exceljs'
import * as dotenv from 'dotenv'
import fs from 'fs'

// https://vk.com/album-225190306_307435564
// https://vk.com/album-225190306_307783090

const OWNER_ID = -225190306
const ALBUM_ID = 307783090
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

const exportToExcel = async (
  comments: any[],
  photos: any[],
  filename: string
): Promise<void> => {
  const workbook = new ExcelJS.Workbook()
  const worksheet = workbook.addWorksheet('Comments')

  worksheet.columns = [
    {header: 'Photo URL', key: 'photo_url', width: 30},
    {header: 'User Name', key: 'user_name', width: 25},
    {
      header: 'Comment',
      key: 'text',
      width: 30,
      style: {
        alignment: {
          wrapText: true,
        },
      },
    },
    {header: 'Date', key: 'date', width: 20},
    {width: 27},
    {width: 27},
  ]

  comments.forEach((comment, index, comments) => {
    const row = worksheet.addRow({
      photo_url: {
        text: `Просмотр фото №${comment.pid}`,
        hyperlink: comment.photo_url,
      },
      user_name: comment.from
        ? `${comment.from.first_name} ${comment.from.last_name}`
        : 'owner',
      text: comment.text,
      date: new Date(comment.date * 1000).toLocaleString(),
    })

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
          ext: {width: 250, height: 250},
        })
      }
      row.height = 200
    }

    if (comment.pid !== comments[index + 1]?.pid) {
      const rowNumber = row.number + 1
      const imageId = workbook.addImage({
        filename: `${IMAGES_FOLDER}${comment.pid}.jpg`,
        extension: 'jpeg',
      })
      worksheet.addImage(imageId, {
        tl: {col: 0, row: rowNumber - 1},
        ext: {width: 250, height: 250},
      })
      const imageRow = worksheet.getRow(rowNumber)
      const descriptionCell = worksheet.getCell(`C${rowNumber}`)
      descriptionCell.value = photos.find(
        (photo) => photo.id === comment.pid
      ).text
      imageRow.height = 200
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
  const pSizeUrl = photo.sizes.find((size) => size.type === 'p').url
  if (!fs.existsSync(`${IMAGES_FOLDER}${photo}.jpg`)) {
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
