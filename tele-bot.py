from os import getenv, path, remove
from utils import path_project, permitionUser
from shutil import rmtree
from dotenv import load_dotenv
from uvloop import install
from pyrogram import Client, filters
from pyrogram.types import Message
from process_excel import Main

load_dotenv()
install()
TOKEN = "7450111423:AAHYEXagpxOesrI4n6uU4g_qFXygPxTWqaU"

app = Client(
  'route_bot',
  api_id=getenv('TELEGRAM_API_ID'),
  api_hash=getenv('TELEGRAM_API_HASH'),
  bot_token=getenv('TELEGRAM_BOT_TOKEN')
)

def log_info_user(message):
  from datetime import datetime
  data = datetime.now().strftime('%d/%m/%Y %H:%M:%S')

  user = message.from_user

  try:
    print(f'''==============INFO USER===============
Data: {data}
ID: {user.id}
Usuario: {user.username}
Telefone: {user.phone_number}
======================================
    ''')
  except:
    pass


@app.on_message(filters.command('/agrupar'))
async def group(client: Client, message: Message):
  await message.reply('Envie o endereço principal')

  response = await app.listen(message.chat.id, filters=filters.text, timeout=60)

  print(response)


@app.on_message(filters.document)
async def document(client: Client, message: Message):
  try:
    log_info_user(message)
    user_id = message.from_user.id

    # permitionUser(user_id)

    
    file_name = message.document.file_name
    file_extension = path.splitext(file_name)[1]

    if file_extension != '.xlsx':
      raise Exception('**Apenas arquivos .xlsx são permitidos!**')

    await message.reply('**Aguarde**', quote=True)
    document_path = await message.download()

    processed_file = Main(document_path)
    if processed_file:
      await client.send_document(chat_id=message.chat.id, document=processed_file)

      path_to_delete = f'{path_project()}/downloads/'    
      rmtree(path_to_delete)
      remove(processed_file)
      
    else:
      app.reply_to(message, '**Aguarde...**', quote=True)

  except Exception as error:
    print(error)
    await message.reply(error, quote=True)



@app.on_message()
async def messages(client: Client, message: Message):
  log_info_user(message)
  
  await message.reply('Somente arquivos excel: XLSX')

print('bot no ar!!!')
app.run()