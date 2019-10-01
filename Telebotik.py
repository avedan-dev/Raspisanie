import telebot
import excel_print
import time_my

print(telebot.__file__)

bot = telebot.TeleBot('792575208:AAFaPrc_J7L_mMLEtPDuskornONkCeVyqB8')
keyboard_faculties = telebot.types.ReplyKeyboardMarkup(one_time_keyboard = True)
keyboard_faculties.row('Горный', 'ГРФ')
keyboard_faculties.row('Нефтегаз', 'ФПМС', 'Строительный')
keyboard_faculties.row('Эконом', 'ЭМФ')

keyboard_courses = telebot.types.ReplyKeyboardMarkup(one_time_keyboard  = True)
keyboard_courses.row('1 курс', '2 курс')
keyboard_courses.row('3 курс', '4 курс')
keyboard_courses.row('5 курс', '6 курс')

keyboard_func = telebot.types.ReplyKeyboardMarkup()
keyboard_func.row('Следующая пара')
keyboard_func.row('Пары на завтра')
keyboard_func.row('Пары на сегодня')
keyboard_func.row('Какая идет неделя?')
keyboard_func.row('Сменить группу')

groups=[]
fac=''

def create_keyboard(fac, cours):
    global groups
    groups=excel_print.groups(fac,cours)
    keyboard_groups = telebot.types.ReplyKeyboardMarkup(one_time_keyboard  = True)
    for i in range(0,len(groups)-len(groups)%3,3):
        keyboard_groups.add(groups[i], groups[i+1], groups[i+2])
    if len(groups)%3!=0:
        if len(groups)%3==2:
            keyboard_groups.add(groups[i+3], groups[i+4])
        else:
            keyboard_groups.add(groups[i+3])
    return(keyboard_groups)

@bot.message_handler(commands=['start'])
def start_message(message):
    bot.send_message(message.chat.id, 'Привет, горняк, выбери свой факультет на клавиатуре внизу', reply_markup=keyboard_faculties)

@bot.message_handler(func=lambda mess: 'Сменить группу' == mess.text)
def start_message(message):
    bot.send_message(message.chat.id, 'Привет, горняк, выбери свой факультет на клавиатуре внизу', reply_markup=keyboard_faculties)

@bot.message_handler(func=lambda mess: 'Сменить группу' == mess.text)
def start_message(message):
    bot.send_message(message.chat.id, 'Привет, горняк, выбери свой факультет на клавиатуре внизу', reply_markup=keyboard_faculties)

@bot.message_handler(func=lambda mess: 'эконом' == str(mess.text).lower() or 'горный' == str(mess.text).lower() or
                     'грф' == str(mess.text).lower() or 'фпмс' == str(mess.text).lower() or 'строительный' == str(mess.text).lower() or
                     'эмф' == str(mess.text).lower() or 'нефтегаз' == str(mess.text).lower(), content_types=['text'])
def send_welcom(message):
    global fac
    if message.text.lower() == 'эконом':
        bot.send_message(message.chat.id, 'Теперь укажи курс на котором обучаешься', reply_markup=keyboard_courses)
        fac='эконом'
        ids=message.chat.id
    elif message.text.lower() == 'нефтегаз':
        bot.send_message(message.chat.id, 'Теперь укажи курс на котором обучаешься', reply_markup=keyboard_courses)
        fac='нефтегаз'
        ids=message.chat.id
    elif message.text.lower() == 'строительный':
        bot.send_message(message.chat.id, 'Теперь укажи курс на котором обучаешься', reply_markup=keyboard_courses)
        fac='строительный'
        ids=message.chat.id
    elif message.text.lower() == 'горный':
        bot.send_message(message.chat.id, 'Теперь укажи курс на котором обучаешься', reply_markup=keyboard_courses)
        fac='горный'
        ids=message.chat.id
        excel_print.izm(ids,fac,'','')
    elif message.text.lower() == 'фпмс':
        bot.send_message(message.chat.id, 'Теперь укажи курс на котором обучаешься', reply_markup=keyboard_courses)
        fac='фпмс'
        ids=message.chat.id
    elif message.text.lower() == 'эмф':
        bot.send_message(message.chat.id, 'Теперь укажи курс на котором обучаешься', reply_markup=keyboard_courses)
        fac='эмф'
        ids=message.chat.id
    elif message.text.lower() == 'грф':
        bot.send_message(message.chat.id, 'Теперь укажи курс на котором обучаешься', reply_markup=keyboard_courses)
        fac='грф'
        ids=message.chat.id

@bot.message_handler(func=lambda mess: '1 курс' == mess.text or '2 курс' == mess.text or
                     '3 курс' == mess.text or '4 курс' == mess.text or '5 курс' == mess.text or
                     '6 курс' == mess.text, content_types=['text'])
def send_course(message):
    global cours
    if message.text == '1 курс':
        cours=message.text[0]
        keyboard_groups=create_keyboard(fac,cours)
        ids=message.chat.id
        bot.send_message(message.chat.id, 'Выбери свою группу', reply_markup=keyboard_groups)   
    elif message.text == '2 курс':
        cours=message.text[0]
        keyboard_groups=create_keyboard(fac,cours)
        ids=message.chat.id
        bot.send_message(message.chat.id, 'Выбери свою группу', reply_markup=keyboard_groups)
    elif message.text == '3 курс':
        cours=message.text[0]
        keyboard_groups=create_keyboard(fac,cours)
        ids=message.chat.id
        bot.send_message(message.chat.id, 'Выбери свою группу', reply_markup=keyboard_groups)
    elif message.text == '4 курс':
        cours=message.text[0]
        keyboard_groups=create_keyboard(fac,cours)
        ids=message.chat.id
        bot.send_message(message.chat.id, 'Выбери свою группу', reply_markup=keyboard_groups)
    elif message.text == '5 курс':
        cours=message.text[0]
        if (fac=='эконом'):
            bot.send_message(message.chat.id, 'Ты выбрал не тот курс или факультет, выбери новый курс, чтобы выбрать другой факультет введи /start', reply_markup=keyboard_courses)
        else:
            keyboard_groups=create_keyboard(fac,cours)
            ids=message.chat.id
            bot.send_message(message.chat.id, 'Выбери свою группу', reply_markup=keyboard_groups)
    elif message.text == '6 курс':
        cours=message.text[0]
        if (fac=='нефтегаз' or fac=='грф' or fac=='эконом'):
            bot.send_message(message.chat.id, 'Ты выбрал не тот курс или факультет, выбери новый курс, чтобы выбрать другой факультет введи /start', reply_markup=keyboard_courses)
        else:
            keyboard_groups=create_keyboard(fac,cours)
            ids=message.chat.id
            bot.send_message(message.chat.id, 'Выбери свою группу', reply_markup=keyboard_groups)

@bot.message_handler(func=lambda mess: groups.count(mess.text)==1, content_types=['text'])
def send_test(message):
    group=message.text
    bot.send_message(message.chat.id, 'Давай проверим твои данные, твой факультет: '+fac +'\nтвой курс: '+cours+'\nтвоя группа: '+ group)
    bot.send_message(message.chat.id, 'Если что-то неверно, то пройди процедуру заново, для этого введи /start')
    bot.send_message(message.chat.id, 'Если все верное, то можешь посмотреть свою расписание на завтра или на следующий день с помощью кнопок внизу', reply_markup=keyboard_func)
    ids=message.chat.id
    excel_print.izm(ids,fac,cours,group)


@bot.message_handler(func=lambda mess: mess.text=='Следующая пара')
def send_next_les(message):
    n=excel_print.output(str(message.chat.id))
    s=excel_print.next_les(n[0], n[1], n[2])
    bot.send_message(message.chat.id, s)

@bot.message_handler(func=lambda mess: mess.text=='Пары на завтра')
def send_next_day(message):
    n=excel_print.output(str(message.chat.id))
    s=excel_print.next_day(n[0], n[1], n[2])
    bot.send_message(message.chat.id, s)

@bot.message_handler(func=lambda mess: mess.text=='Пары на сегодня')
def send_next_day(message):
    n=excel_print.output(str(message.chat.id))
    s=excel_print.to_day(n[0], n[1], n[2])
    bot.send_message(message.chat.id, s)

@bot.message_handler(func=lambda mess: mess.text=='Какая идет неделя?')
def send_next_day(message):
    n=int(time_my.week_number())
    if n%2==1:
        bot.send_message(message.chat.id, 'Нечетная')
    else:
        bot.send_message(message.chat.id, 'Четная')        


bot.polling()
