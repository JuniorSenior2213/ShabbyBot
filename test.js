const axios = require('axios');

// Ваш API ключ от Pushbullet
const apiKey = 'o.o3adbq8icgTSNCVjd48yuzlIZSZqKWuV';

// URL для Pushbullet API
const url = 'https://api.pushbullet.com/v2/pushes';

// Данные для отправки
const data = {
  type: 'note',
  title: 'ПРИВЕТ!',
  body: 'ПРИВЕТ!',
  sound: 'boop'
};

// Заголовки для авторизации
const headers = {
  'Access-Token': apiKey
};

// Отправка POST запроса
axios.post(url, data, { headers })
  .then(response => {
    console.log('Сообщение отправлено:', response.data);
  })
  .catch(error => {
    console.error('Ошибка при отправке сообщения:', error);
  });
