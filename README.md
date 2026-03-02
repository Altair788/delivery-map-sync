# 🗺️ Delivery Map Sync

Автоматизированная система визуализации логистических данных на карте (на примере г. Москвы).

## ✨ Возможности

- Импорт данных из Google Sheets (адреса + статусы)
- Автоматический геокодинг через Nominatim (OpenStreetMap)
- Автообновление: каждые 60 сек (карта) + 5 мин (синхронизация)
- Цветовая дифференциация по 5 статусам
- ☑Управление видимостью слоёв через чекбоксы
- Экспорт в KML для Google Earth / GIS-систем
- Кэширование геокодов для экономии лимитов API
- Обработка ошибок: логирование в отдельный лист, graceful degradation

## 🛠 Технологический стек

| Компонент | Технология |
|-----------|-----------|
| Backend | Google Apps Script (JavaScript ES6) |
| Frontend | Leaflet.js + OpenStreetMap |
| Геокодинг | Nominatim API |
| Хранение | Google Sheets (GeoCache, GeocodingErrors) |
| Деплой | Google Apps Script Web App |

## 🚀 Быстрый старт

1. Откройте вашу таблицу в Google Sheets
2. Откройте [Apps Script](https://script.google.com/) → Новый проект
3. Вставьте код из `apps-script/Code.js` и `MapPage.html` в соответствующие файлы в Apps Sript (Code.gs, MapPage.html)
4. Настройте `CONFIG` в начале `Code.js`
5. Запустите `create5MinTrigger()` для автообновления
6. Разверните веб-приложение: **Развернуть** → **Новое развертывание** → **Веб-приложение**

## 📈 Метрики

| Показатель | Значение |
|-----------|----------|
| Обработка адресов | ~47 за ~45 сек (с кэшем) |
| Точность геокодинга | >95%  |
| Время обновления | 1-5 минут после изменения в таблице |
| Лимиты API | Соблюдение Nominatim (1 req/sec), Apps Script (6 min/exec) |

## 🤝 Contributing

1. Fork репозиторий
2. Создайте feature-ветку (`git checkout -b feature/amazing-feature`)
3. Закоммитьте изменения (`git commit -m 'Add amazing feature'`)
4. Запушьте (`git push origin feature/amazing-feature`)
5. Откройте Pull Request

---

**Автор:** Эдуард Слободяник  
**Контакты:** [Telegram](https://t.me/zermatt788) | [Email](mailto:slobodyanik.eduard@yandex.ru) | [GitHub](https://github.com/Altair788)
