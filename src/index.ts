import axios, {AxiosInstance} from 'axios'
import * as dotenv from 'dotenv'

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
            owner_id: -225190306,
            album_id: 307654146, //307435564  307654146
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

const allComments = await getAllAlbumComments(instance)  //Получаем все комменты
console.log('allComments.length :', allComments.length)
// console.log('allComments[0] :', allComments[28], allComments[55])

const usersIds = new Set<number>()  //получаем массив уникальных пользователей
allComments.forEach((comment) => {
  usersIds.add(comment.from_id)
})

const uniqUsers = await getUsersByIds(Array.from(usersIds), instance)
uniqUsers.push(
  {
    id: -225190306,
    first_name: 'GroupOwner',
  }
)


const commentsWithUsers = allComments.map((comment) => {
  const user = uniqUsers.find((user) => comment.from_id === user.id)
  comment.from = user
  return comment
})

console.log('commentsWithUsers :', commentsWithUsers)