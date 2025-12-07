// kredensial
const spreadsheetId      = '1JRXIjxCDYmTWbHO9nfn4N3vedAISCEphIWfDAjrJQmE'
const dataOrderSheetName = 'Data VM'
const logSheetName       = 'Log'

const botHandle      = '@VoiceMemberUB3Bot'
const botToken       = '8396135443:AAEzQ-kjDaYVkonNvpEQCAYH3EOvmo2h8yw'
const appsScriptUrl  = 'https://script.google.com/macros/s/AKfycbxJw1fnrka0hVI6z15-JwqwNx7nGMWAnuB_GSLsMkQ7IxGQNIUmiYO7dEmXtrqR51QpZg/exec'
const telegramApiUrl = `https://api.telegram.org/bot${botToken}`


function log(logMessage = '') {
  // akses sheet
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId)
  const sheet       = spreadsheet.getSheetByName(logSheetName)
  const lastRow     = sheet.getLastRow()
  const row         = lastRow + 1

  // inisiasi nilai
  const today = new Date

  // insert row kosong
  sheet.insertRowAfter(lastRow)

  // insert data
  sheet.getRange(`A${row}`).setValue(today)
  sheet.getRange(`B${row}`).setValue(logMessage)
}


function formatDate(date) {
  const monthIndoList = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Ags', 'Sep', 'Okt', 'Nov', 'Des']

  const dateIndo  = date.getDate()
  const monthIndo = monthIndoList[date.getMonth()]
  const yearIndo  = date.getFullYear()

  const result = `${dateIndo} ${monthIndo} ${yearIndo}`

  return result
}


function sendTelegramMessage(chatId, replyToMessageId, textMessage) {
  // url kirim pesan
  const url = `${telegramApiUrl}/sendMessage`;
  
  // payload
  const data = {
    parse_mode              : 'HTML',
    chat_id                 : chatId,
    reply_to_message_id     : replyToMessageId,
    text                    : textMessage,
    disable_web_page_preview: true,
  }
  
  const options = {
    method     : 'post',
    contentType: 'application/json',
    payload    : JSON.stringify(data)
  }

  const response = UrlFetchApp.fetch(url, options).getContentText()
  return response;
}


function parseMessage(message = '') {
  // pisahkan berdasarkan karakter enter
  const splitted = message.split('\n')

  // inisiasi variabel
  let nama       = ''
  let proses     = ''
  let line       = ''
  let noreg      = ''
  let voice      = ''

  // parsing pesan untuk mencari nilai variabel
  splitted.forEach(el => {
    nama = el.includes('Nama:') ? el.split(':')[1].trim().replaceAll('\n', ' ') : nama;
    proses = el.includes('Proses:') ? el.split(':')[1].trim().replaceAll('\n', ' ') : proses;
    line = el.includes('Line:') ? el.split(':')[1].trim().replaceAll('\n', ' ') : line;
    noreg = el.includes('Noreg:') ? el.split(':')[1].trim().replaceAll('\n', ' ') : noreg;
    voice = el.includes('Voice Member:') ? el.split(':')[1].trim().replaceAll('\n', ' ') : voice;
  })

  // kumpulkan hasil
  const result = {
    nama      : nama,
    proses    : proses,
    line      : line,
    noreg     : noreg,
    voice     : voice,
  }

  // jika data kosong
  const isEmpty = (nama === '' && proses === '' && line === '' && noreg === '' && voice === '')

  return isEmpty ? false : result
}


function inputDataOrder(data) {
  try {
    // akses sheet
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId)
    const sheet = spreadsheet.getSheetByName(dataOrderSheetName)
    const lastRow = sheet.getLastRow()
    const row = lastRow + 1

    // inisiasi nilai
    const number  = lastRow
    const idOrder = `VM-${number}`
    const today   = new Date

    // insert row kosong
    sheet.insertRowAfter(lastRow)

    // insert data
    sheet.getRange(`A${row}`).setValue(number)
    sheet.getRange(`B${row}`).setValue(idOrder)
    sheet.getRange(`C${row}`).setValue(today)
    sheet.getRange(`D${row}`).setValue(data['nama'])
    sheet.getRange(`E${row}`).setValue(data['proses'])
    sheet.getRange(`F${row}`).setValue(data['line'])
    sheet.getRange(`G${row}`).setValue(data['noreg'])
    sheet.getRange(`H${row}`).setValue(data['voice'])
    sheet.getRange(`I${row}`).setValue('Sedang diproses')
    sheet.getRange(`J${row}`).setValue(data['chatId'])

    // jika berhasil, return idOrder
    return idOrder

  } catch(err) {
    return false
  }
}


function cekNoreg(noreg = null) {
  // cegah noreg kosong
  if (!noreg) {
    return 'Format pencarian noreg tidak valid.'
  }

  // akses sheet
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId)
  const sheet       = spreadsheet.getSheetByName(dataOrderSheetName)
  const lastRow     = sheet.getLastRow()

  // ambil data
  const range    = `A2:J${lastRow}`
  const dataList = sheet.getRange(range).getValues()

  // filter data
  const dataListFiltered = dataList.filter(el => el[6].toString().toLowerCase() === noreg.toString().toLowerCase())

  // cek jika noreg ditemukan  
  const isResiFound = dataListFiltered.length > 0

  // variabel balasan
  let messageReply = ''

  // jika ditemukan
  if (isResiFound) {
		// jika ada no resi yang sama, yang diambil yang paling atas
    const data = dataListFiltered[0]

    messageReply = `Info noreg <b>${noreg}</b>

ID Order: ${data[1]}
Tanggal Order: ${formatDate(data[2])}
Nama: ${data[3]}
Proses: ${data[4]}
Line: ${data[5]}
Voice Member: ${data[7]}
Status Penanggulangan: <b>${data[8]}</b>`
  
  // jika tidak
  } else {
    messageReply = `Noreg ${noreg} tidak ditemukan.`
  }

  return messageReply
}


function handleUpdateDeliveryStatus(e) {
  // ambil info sheet dan row yang baru diedit
	const row       = e.range.getRow()
  const column    = e.range.getA1Notation().replace(/[^a-zA-Z]/g, '')
  const sheetName = e.range.getSheet().getSheetName()

	// jika perubahan bukan pada sheet data order kolom H
  if (sheetName !== dataOrderSheetName || column !== 'I') {
    return false
  }

  // akses sheet
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId)
  const sheet       = spreadsheet.getSheetByName(dataOrderSheetName)
  const today       = new Date

  // ambil data
  const range = `A${row}:J${row}`
  const data  = sheet.getRange(range).getValues()

  // isi konstanta
  const idOrder          = data[0][1]
  const tanggalOrder     = data[0][2]
  const nama             = data[0][3]
  const proses           = data[0][4]
  const line             = data[0][5]
  const noreg            = data[0][6]
  const voice            = data[0][7]
  const statusPenanggulangan = data[0][8]
  const chatId           = data[0][9].toString()

  const textMessage = `Update Info Noreg <b>${noreg}</b>

ID Order: ${idOrder}
Tanggal Order: ${formatDate(tanggalOrder)}
Nama: ${nama}
Proses: ${proses}
Line: ${line}
Voice Member: ${voice}
Status Penanggulangan: <b>${statusPenanggulangan}</b>

<i>Data per-${formatDate(today)}</i>`

  // kirim pesan
  sendTelegramMessage(chatId, null, textMessage)
}


function doPost(e) {
  try {
    // urai pesan masuk
    const contents            = JSON.parse(e.postData.contents)
    const chatId              = contents.message.chat.id
    const receivedTextMessage = contents.message.text.replace(botHandle, '').trim() // hapus botHandle jika pesan berasal dari grup
    const messageId           = contents.message.message_id

    let messageReply = ''

    // 1. jika pesan /start
    if (receivedTextMessage.toLowerCase() === '/start') {
      // tulis pesan balasan
      messageReply = `Halo! Status bot dalam keadaan aktif.`

    // 2. jika pesan diawali dengan /input
    } else if (receivedTextMessage.split('\n')[0].toLowerCase() === '/input') {
      const parsedMessage = parseMessage(receivedTextMessage)

      // 2a.jika ada data
      if (parsedMessage) {
        const data = {
          nama      : parsedMessage['nama'],
          proses    : parsedMessage['proses'],
          line      : parsedMessage['line'],
          noreg     : parsedMessage['noreg'],
          voice     : parsedMessage['voice'],
          chatId    : chatId
        }

        // insert data ke sheet
        const idOrder = inputDataOrder(data)

        // tulis pesan balasan
        messageReply = idOrder ? `Data berhasil disimpan dengan ID Order <b>${idOrder}</b>` : 'Data gagal disimpan'

      // 2b. jika tidak ada data
      } else {
        messageReply = 'Data kosong dan tidak dapat disimpan'
      }

    // 3. cek noreg
    } else if (receivedTextMessage.split(' ')[0].toLowerCase() === '/noreg') {
      // ambil noreg
      const noreg = receivedTextMessage.split(' ')[1]

      // ambil info
      messageReply = cekNoreg(noreg)

    // 4. format
    } else if (receivedTextMessage.toLowerCase() === '/format') {
      messageReply = `Untuk <b>input data order</b> gunakan format:

<pre>/input
Nama: 
Proses: 
Line: 
Noreg:
Voice Member: </pre>

Untuk <b>cek Data</b> gunakan format:

<pre>/noreg [noreg TMMIN]</pre>
(Tanpa tanda kurung siku)`

    // 5. format salah
    } else {
      messageReply = `Pesan yang Anda kirim tidak sesuai format.

Kirim perintah /format untuk melihat daftar format pesan yang tersedia.`
    }

    // kirim pesan balasan
    sendTelegramMessage(chatId, messageId, messageReply)

  } catch(err) {
    log(err)
  }
}


function setWebhook() {
  // akses api
  const url      = `${telegramApiUrl}/setwebhook?url=${appsScriptUrl}`
  const response = UrlFetchApp.fetch(url).getContentText()
  
  Logger.log(response)
}
