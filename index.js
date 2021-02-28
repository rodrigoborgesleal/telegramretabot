//npm i convert-excel-to-json
const excelToJson = require('convert-excel-to-json')
//npm i node-telegram-bot-api
const TelegramBot = require('node-telegram-bot-api')
const token = '1636470598:AAEnRXFFYY6LxOEAvWnItsK9DRuY2meM_SY'
const testebot_TOKEN = '1501819929:AAH9r3lFXcRXqy5RkflDA8PHV91RQRVsBXI'
const bot = new TelegramBot(testebot_TOKEN, {polling: true})

async function getSheet() {
  const { sheet } = excelToJson({
    sourceFile: './estoq.xlsx',
    range: 'A2:H39813',
    header:{
        rows: 1
    },
    columnToKey: {
        A: 'code',
        B: 'desc',
        C: 'section',
        D: 'folder',
        E: 'ean1',
        F: 'ean2',
        G: 'ean3',
        H: 'embqt'
    }
  })
  return sheet //RETORNA UM ARRAY
}


async function filterSheet(arg) {
  const sheetData = await getSheet()
  const sheetFiltered = [] 
  await sheetData.forEach((p)=>{
    if(p.code == arg || p.desc.includes(arg) || p.section == arg || p.folder == arg || p.ean1 == arg || p.ean2 == arg || p.ean3 == arg || p.embqt == arg) {
      sheetFiltered.push(p)
    }
  })
  var dataFormated = ''
  await sheetFiltered.forEach((i)=>{
    dataFormated += `Código: ${i.code}\r\nEAN: ${i.ean1}\r\nSeção: ${i.section}\r\nPasta: ${i.folder}\r\nQuant.Emb: ${i.embqt}\r\n${i.desc}\r\n\r\n`
  })
  return dataFormated
}


bot.on('message', async msg => {
  const chatId = msg.chat.id
  const usersData = msg.chat
  const { id, first_name, last_name, username } = usersData
  console.log(`Usuário (${username}) : ${first_name} ${last_name} Enviou: ${msg.text}`) 
  var response
    response = await filterSheet(msg.text.toLocaleUpperCase())
    bot.sendMessage(chatId, response)
  
})

