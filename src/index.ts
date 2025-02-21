import axios, {AxiosInstance} from 'axios'
import ExcelJS from 'exceljs'
import * as dotenv from 'dotenv'
import fs from 'fs'

// https://vk.com/album-225190306_307435564

const OWNER_ID = -225190306
const ALBUM_ID = 307435564 //307654146

dotenv.config()

// настройка instance of axios
const instance = axios.create({
  baseURL: process.env.BASE_URL,
  // timeout: 1000,
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
    // return response.data.response
    .then((response) => {
      // console.log('response :', response)
      return response.data.response
    })
    .catch((error) => {
      console.log('error :', error)
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
      // console.log(response.data.response)
      return response.data.response
    })
    .catch((error) => {
      console.log('error :', error)
    })
}

const exportToExcel = async (
  comments: any[],
  filename: string
): Promise<void> => {
  const workbook = new ExcelJS.Workbook()
  const worksheet = workbook.addWorksheet('Comments')

  worksheet.columns = [
    {header: 'Photo URL', key: 'photo_url', width: 30},
    {header: 'User Name', key: 'user_name', width: 25},
    {header: 'Comment', key: 'text', width: 77},
    {header: 'Date', key: 'date', width: 20},
  ]

  comments.forEach((comment, index, comments) => {
    worksheet.addRow({
      photo_url: {
        text: `Просмотр фото №${comment.pid}`, // Текст ссылки (можно динамически генерировать)
        hyperlink: comment.photo_url, // URL
      },
      user_name: comment.from
        ? `${comment.from.first_name} ${comment.from.last_name}`
        : 'owner',
      text: comment.text,
      date: new Date(comment.date * 1000).toLocaleString(),
    })
    if (comment.pid !== comments[index + 1]?.pid) worksheet.addRow({}).number
  })

  await workbook.xlsx.writeFile(filename)
  console.log(`Exported to ${filename}`)
}

const downloadImage = async (imageUrl: string, fileName: string) => {
  const response = await axios.get(imageUrl, { responseType: 'arraybuffer' });
  fs.writeFileSync(fileName, Buffer.from(response.data, 'binary'));
};

const getAllComments = (axiosInstance: AxiosInstance) => {
  axiosInstance
    .get('photos.getAllComments', {
      params: {
        owner_id: -225190306,
        album_id: 307654146,
        count: 5,
      },
    })
    .then(function (response) {
      // handle success

      console.log(response.data.response.items)
      console.log('items count = ' + response.data.response.items.length)
    })
}

const getUsers = (axiosInstance: AxiosInstance) => {
  axiosInstance
    .get('users.get', {
      params: {
        user_ids: 743784474,
        fields: 'bdate',
      },
    })
    .then(function (response) {
      // handle success
      console.log(response.data)
    })
}

// Просто запрос (без предварительной настройки)
const getUsersWithoutInstance = () => {
  axios
    .get(
      'https://api.vk.com/method/users.get?user_ids=743784474&fields=bdate&v=5.199 HTTP/1.1',
      {
        headers: {
          Authorization: `Bearer ${process.env.ACCESS_TOKEN}`,
        },
      }
    )
    .then(function (response) {
      // handle success
      // console.log(response.data)
    })
    .catch(function (error) {
      // handle error
      console.log(error)
    })
    .finally(function () {
      // always executed
    })
}

const allComments: any[] = await getAllAlbumComments(instance) //Получаем все комменты
console.log('allComments.length :', allComments.length)
// console.log('allComments[0] :', allComments[28], allComments[55])

// сортируем комментарии
allComments.reverse().sort((a, b) => a.pid - b.pid) // TODO посмотреть , может есть возможность в запросе отсортировать

const usersIds = new Set<number>() //получаем массив уникальных пользователей
allComments.forEach((comment) => {
  usersIds.add(comment.from_id)
})

const uniqUsers = await getUsersByIds(Array.from(usersIds), instance)
// костыль добавляющий владельца группы
uniqUsers.push({
  id: -225190306,
  first_name: 'Владычица',
  last_name: 'группы!!!',
})

const commentsWithUsers = allComments.map((comment) => {
  const user = uniqUsers.find((user) => comment.from_id === user.id)
  comment.from = user
  comment.photo_url = `https://vk.com/photo${OWNER_ID}_${comment.pid}`
  return comment
})

console.log('commentsWithUsers :', commentsWithUsers)

exportToExcel(commentsWithUsers, 'test.xls')

downloadImage('https://sun9-40.userapi.com/impg/HVMXzXhOC4OhtX0jx3XWKSgR6jkpitutGCzG2g/BOrjAmGiorQ.jpg?size=800x800&quality=95&sign=a0cb4f940c6382a8af329c7b4a658836&type=album', 'testPhoto.jpg')
