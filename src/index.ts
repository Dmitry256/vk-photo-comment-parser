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

const executeExample = async (axios: AxiosInstance) => {
  const result = await axios
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
    return result.data.response
    // .then((response) => {
    //   // handle success
    //   // console.log(response.data)
    //   console.log(response.data)
    //   // console.log(response.data.response.items[2].attachments[0].photo)
    //   // console.log(response.data.response.count)
    //   return response
    // })
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

const allComments = await executeExample(instance)
console.log('allComments.length :', allComments.length)
