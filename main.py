import telebot
from main_01 import reestr
import os
# from telebot import types
from zipfile import ZipFile


bot = telebot.TeleBot("5970164869:AAH0ccv1c5pG5YpWVtPiJTglKmxK1WcVB_0")

@bot.message_handler(func=lambda message: True)

def handler_all(message):

    def hello():
        bot.reply_to(message, "Привет!\n"
                              "Я бот-помощник сотрудникам SR")

    def bot_superpower():
        bot.reply_to(message, "Сегодня я могу помочь вам с :\n"
                    "- Созданием фин. реестров НП\n\n"
                    "Выберите в Меню:\n"
                     "- Приветствие - что бы повторно вывести данное сообщение\n"
                     "- Приступить - что бы начать делать реестры НП\n"
                     "- Завершить - что бы очистить кэш для следующего раза")

    def clean():
        directory_name = 'пул'
        for file_name in os.listdir(directory_name):
            file_path = os.path.join(directory_name, file_name)
            if os.path.isfile(file_path):
                os.remove(file_path)
        directory_name = 'res'
        for file_name in os.listdir(directory_name):
            file_path = os.path.join(directory_name, file_name)
            if os.path.isfile(file_path):
                os.remove(file_path)
        directory_name = 'zip'
        for file_name in os.listdir(directory_name):
            file_path = os.path.join(directory_name, file_name)
            if os.path.isfile(file_path):
                os.remove(file_path)
        # bot.reply_to(message, "Кэш очищен.")
        file = open("yyyaaa.txt", "w")
        file.truncate()
        file = open("yyyaaasf.txt", "w")
        file.truncate()

    def uploadd():
        for folder_name in ['zip']:
            for file_name in os.listdir(folder_name):
                file_path = folder_name + "/" + file_name
                bot.send_document(message.chat.id, open(file_path, 'rb'))

    def zipp():
        path = "пул"
        zip_name = "zip/archive.zip"
        zip = ZipFile(zip_name, "w")

        for folder, subfolders, files in os.walk(path):
            for file in files:
                file_path = folder + "/" + file
                zip.write(file_path, os.path.relpath(file_path, path))

    if message.text.lower() == "/command1":
        clean()
        hello()
        bot_superpower()

    elif "ya " in message.text.lower():
        with open("yyyaaa.txt", "w") as file:
            file.write(message.text.lower())
        bot.reply_to(message, 'Принято Яндекс СР: %s' %(message.text.lower()))

    elif "ysf " in message.text.lower():
        with open("yyyaaasf.txt", "w") as file:
            file.write(message.text.lower())
        bot.reply_to(message, 'Принято Яндекс СЭЙФ: %s' %(message.text.lower()))


    elif message.text.lower() == "/start":
        clean()
        hello()
        bot_superpower()

    elif message.text.lower() == "/command2":
        clean()
        bot.reply_to(message, "- Если будут отчеты Яндекс, пришли мне платежки яндекса в формате: ya 3 123 456 798\n"
                              "(где ya (для сэйф ysf) обязательное ключевое слово, далле цифрой количество платёжек, и далее каждая платежка (всё через пробел), "
                              "отдельными сообщениями по СЭЙФ и СР\n\n"
                              '- Пришли мне файлы реестров ТК (по одному или все сразу) и я преобразую их в реестры SR\n\n'
                              "**Тебе необходимо строго соблюдать правила именования реестров ТК , смотри справку: https://disk.yandex.ru/i/7wwmGz1bB01wJA \n\n")

    elif message.text.lower() == "/command3":
        bot.reply_to(message, "Процесс завершен, кэш очищен")
        clean()

    elif message.text.lower() == "гоу":
        r = reestr()
        rr = "\n".join(r)
        bot.reply_to(message, 'Функция "Реестры" результат выполнения: \nБыли созданы файлы, Проверь контрольные суммы фалов! \n\n')
        bot.reply_to(message, rr)
        bot.reply_to(message, "напиши мне zip что бы получить архив с этими отчетами, или выбери в меню 'прервать' что бы всё очистить")

    elif message.text.lower() == "zip":
        zipp()
        uploadd()
        bot.reply_to(message, "Вот твой файл , обязательно нажми в меню Завершить - что бы очистить кэш для следующего раза, файл из чата не удалится")

    else:
        bot_superpower()

@bot.message_handler(content_types=['document', 'video', 'audio', 'voice', 'sticker'])
def handle_file(message):
    file_info = bot.get_file(message.document.file_id)
    downloaded_file = bot.download_file(file_info.file_path)
    save_path = message.document.file_name  # сохраняем файл с его исходным именем
    with open(f'пул\{save_path}', 'wb') as new_file:
        new_file.write(downloaded_file)
    bot.reply_to(message, 'Файл %s принят к обработке. \nПродолжай присылать файлы, если закончил напиши Гоу' % (message.document.file_name))

if __name__ == "__main__":
    bot.polling(none_stop=True)
