const XLSX = require('xlsx')
const path = require('path')
const filePath = path.join(__dirname, 'list.xlsx')
// Чтение файла
const workbook = XLSX.readFile(filePath)

// Получение данных с листа
const sheetNames = workbook.SheetNames
const firstSheetName = sheetNames[0]
const worksheet = workbook.Sheets[firstSheetName]
const dataXlsx = XLSX.utils.sheet_to_json(worksheet, { header: 'A' })
const columnH = 'H'
const columnI = 'I'
const columnJ = 'J'
let startRow = 2
let currentIndex = 1
console.log(dataXlsx.length)
function fetchData() {
    if (dataXlsx.length === 1) {
        console.log('Нет данных')
        return
    }
    // Делаем первый запрос что бы получить requestId и токены
    if (currentIndex < dataXlsx.length ) {  // Проверка дабы не вызвать бесконеный цикл рекурсии
        const currentItem = dataXlsx[currentIndex];
        const dateObj = (item) => {
            // Преобразую дату из документа в нормальный вид
            const obj = XLSX.SSF.parse_date_code(item)
            if (obj.d && obj.m && obj.y) {
                return obj.d + '.' + obj.m + '.' + obj.y
            }
        }
        const dataBird = dateObj(currentItem.D)
        const dataDoc = dateObj(currentItem.G)
        const data = `c=find&captcha=&captchaToken=&fam=${currentItem.A}&nam=${currentItem.B}&otch=${currentItem.C}&bdate=${dataBird}&doctype=${currentItem.E}&docno=${String(currentItem.F).split(' ').join('+')}&docdt=${dataDoc}` // склеваю строку для запроса
        fetch("https://service.nalog.ru/inn-new-proc.json", {
            "headers": {
                "accept": "application/json, text/javascript, */*; q=0.01",
                "accept-language": "ru,en;q=0.9",
                "content-type": "application/x-www-form-urlencoded; charset=UTF-8",
                "sec-ch-ua": "\"Not_A Brand\";v=\"8\", \"Chromium\";v=\"120\", \"YaBrowser\";v=\"24.1\", \"Yowser\";v=\"2.5\"",
                "sec-ch-ua-mobile": "?0",
                "sec-ch-ua-platform": "\"Windows\"",
                "sec-fetch-dest": "empty",
                "sec-fetch-mode": "cors",
                "sec-fetch-site": "same-origin",
                "x-requested-with": "XMLHttpRequest",
                "cookie": "JSESSIONID=8149006E4556A6C6BEB40B32054ECC92; uniI18nLang=RUS; _ym_uid=1708084357154048989; _ym_d=1708084357; _ym_isad=2; _ym_visorc=b; upd_inn=2AF84B5AEC7DECCFCB8B5BF3431C21815BBBB3DA1C11941AC1080001FE2F1F2D9D2D4475FD9D9C329D7195A22F86F694",
                "Referer": "https://service.nalog.ru/inn.do",
                "Referrer-Policy": "strict-origin-when-cross-origin"
            },
            "body": `${data}`,
            "method": "POST"
        })
            .then(r => r.json())
            .then(response => {
                if (response.requestId !== '') {
                    // делаем 2 запрос что бы уже получить сам результат
                    const dataGet = `c=get&requestId=${response.requestId}` // формируем строку запроса их полученных доступов
                    fetch("https://service.nalog.ru/inn-new-proc.json", {
                        "headers": {
                            "accept": "application/json, text/javascript, */*; q=0.01",
                            "accept-language": "ru,en;q=0.9",
                            "content-type": "application/x-www-form-urlencoded; charset=UTF-8",
                            "sec-ch-ua": "\"Not_A Brand\";v=\"8\", \"Chromium\";v=\"120\", \"YaBrowser\";v=\"24.1\", \"Yowser\";v=\"2.5\"",
                            "sec-ch-ua-mobile": "?0",
                            "sec-ch-ua-platform": "\"Windows\"",
                            "sec-fetch-dest": "empty",
                            "sec-fetch-mode": "cors",
                            "sec-fetch-site": "same-origin",
                            "x-requested-with": "XMLHttpRequest",
                            "cookie": "JSESSIONID=8149006E4556A6C6BEB40B32054ECC92; uniI18nLang=RUS; _ym_uid=1708084357154048989; _ym_d=1708084357; _ym_isad=2; _ym_visorc=b; upd_inn=2AF84B5AEC7DECCFCB8B5BF3431C21815BBBB3DA1C11941AC1080001FE2F1F2D9D2D4475FD9D9C329D7195A22F86F694",
                            "Referer": "https://service.nalog.ru/inn.do",
                            "Referrer-Policy": "strict-origin-when-cross-origin"
                        },
                        "body": `${dataGet}`,
                        "method": "POST"
                    })
                        .then(responseGet => responseGet.json())
                        .then(result => {
                            const cellAddressH = `${columnH}${startRow}`;  // Получаем колонку с ИНН
                            const cellAddressI = `${columnI}${startRow}`; // Получаем колонку с Паспортом
                            if (result && result.inn !== undefined) {
                                worksheet[cellAddressH] = { t: 's', v: result.inn }; // ПОдготовливаем строку для записи
                                worksheet[cellAddressI] = { t: 's', v: String('ИНН:' + result.inn + ' ' + 'Паспорт:' + currentItem.F) }; // ПОдготовливаем строку для записи
                                XLSX.writeFile(workbook, filePath); // Записываем в файл
                                console.log('Данные успешно записаны в файл.', currentIndex, result.result, currentItem.F);
                                startRow += 1; // Увиличиваем шаг на 1 что бы писать в след строку
                            } else {
                                worksheet[cellAddressH] = { t: 's', v: 'Не найдено' };
                                worksheet[cellAddressI] = { t: 's', v: String('ИНН:' + 'Не найдено' + ' ' + 'Паспорт:' + currentItem.F) };
                                XLSX.writeFile(workbook, filePath);
                                console.log('Данные успешно записаны в файл.', currentIndex, 'Не найдено', currentItem.F);
                                startRow += 1;
                            }
                            currentIndex += 1 // Увеличиваем элемент который получаем для отправки на удачи
                            setTimeout(fetchData, 3000); // Вызываем функцию ещё раз с таймаутом в 3с
                        })
                        .catch(ErrGet => {
                            const cellAddressJ = `${columnJ}${startRow}`;
                            worksheet[cellAddressJ] = { t: 's', v: `Ошибка на 2 уровне : ${ErrGet}` };
                            startRow += 1;
                            currentIndex += 1 // Увеличиваем элемент который получаем для отправки при не удаче
                            setTimeout(fetchData, 3000);
                        })
                }
            })
            .catch(err => {
                const cellAddressJ = `${columnJ}${startRow}`;
                worksheet[cellAddressJ] = { t: 's', v: `Ошибка на 1 уровне : ${err}` };
                startRow += 1;
                currentIndex += 1
                setTimeout(fetchData, 3000);
            })
    }
}



fetchData()

