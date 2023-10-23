import os
import argh
import datetime
import pandas as pd
import requests
import json
import random


def timestamp():
    return str(datetime.datetime.now()) + ': '


# исправиляем названия колонк, если они съехали
def fix_header(df):
    if 'Time Started' in df.iloc[0].values:
        df.columns = df.iloc[0]
        df.drop([0], inplace=True)

    if 'Column1' in df.columns:
        with open('col_names_new.txt', 'r', encoding='utf-8') as f:
            colnames = f.read().split('\n')
        df.columns = colnames + [""]
    df.drop([0, 1, 2, 3], inplace=True)
    return df


# функция для склеивания колонок с текстовым ответом
# def combine_text(nps: int, text1: str, text2: str) -> str:
#     if nps < 9:
#         return text1
#     else:
#         return text2


# функция для создания нового названия файла
def make_filename(filename: str) -> str:
    filename = filename.replace(".xlsx", "")
    start, end = [d.split('.') for d in filename.split()[1].split('-')]

    year = ""
    month = ""
    day_s = start[0]
    day_e = end[0]

    if len(start) == 2:
        month = start[1]
    else:
        month = end[1]

    if len(end[2]) == 2:
        year = '20' + end[2]
    else:
        year = end[2]

    return f'auto_{year}.{month}.{day_s}-{day_e}_AFFB.xlsx'


def fix_scoreOnBoard(score):
    if score == "Плохо":
        return 1
    elif score == "Так себе":
        return 2
    elif score == "Могло быть и лучше":
        return 3
    elif score == "Хорошо":
        return 4
    elif score == "Замечательно":
        return 5
    else:
        return 0

# функция для создания списка словарей из json формата колонки "Параметры"
def parse_json(df):
    list_of_params = list(df.iloc[:, 3])
    for i in range(len(list_of_params)):
        list_of_params[i] = json.loads(list_of_params[i])

    return list_of_params

def create_columns(df, list_of_params):
    df["Status"], df["Longitude"], df["Latitude"], df["Country"], \
    df["City"], df["URL Variable: pax"], df["URL Variable: anx_list"], df['who'], \
    df["Другая цель (уточните, какая именно?): Нам очень помогут детали вашего путешествия    Основная цель вашей " \
       "поездки была..."], \
    df["Да, в аэропорту:Вы обращались за помощью к сотрудникам авиакомпании?Для того, чтобы снять выделение с ответа, "\
       "нажмите еще раз на выделенную галочку  "], \
    df["Да, по телефону:Вы обращались за помощью к сотрудникам авиакомпании?Для того, чтобы снять выделение с ответа," \
       " нажмите еще раз на выделенную галочку"], \
    df["Да, в чате:Вы обращались за помощью к сотрудникам авиакомпании?Для того, чтобы снять выделение с ответа," \
       " нажмите еще раз на выделенную галочку"], \
    df["Нет:Вы обращались за помощью к сотрудникам авиакомпании?Для того, чтобы снять выделение с ответа, нажмите еще" \
       " раз на выделенную галочку"] = \
        None, None, None, None, None, None, None, None, None, None, None, None, None

    df['Language'] = 'Russian'
    df['Date Submitted'] = df['Time Started'].iloc[:]


    # заполняем ячейки колонок из списка "names", если есть значение по ключу из "keys" , в ином случае оставляем поле пустым
    names = ['URL Variable: airport_destination', 'URL Variable: airport_origin', 'URL Variable: ap1',
             'URL Variable: ap2', 'URL Variable: baggage', 'URL Variable: cabin', 'URL Variable: delay',
             'URL Variable: flightdate', 'URL Variable: flightno', 'URL Variable: paxtype', 'URL Variable: pnr',
             'URL Variable: pnr_create_date', 'URL Variable: regchannel', 'URL Variable: spnr', 'URL Variable: transfer',
             'URL Variable: autoreg', 'URL Variable: earlyflight', 'URL Variable: emd', 'URL Variable: extraseat',
             'URL Variable: gatefood', 'URL Variable: pet', 'URL Variable: refund', 'URL Variable: s7boost',
             'URL Variable: sales_channel', 'URL Variable: sales_interaction_type', 'URL Variable: sms']

    keys = ['AIRPORT_DESTINATION', 'AIRPORT_ORIGIN', 'ap1', 'ap2', 'baggage', 'Cabin', 'delay', 'FlightDate',
            'FlightNo', 'paxtype', 'pnr', 'PNR_CREATE_DATE', 'regchannel', 'spnr', 'transfer', 'autoreg', 'earlyflight',
            'EMD', 'extraseat', 'gatefood', 'pet', 'refund', 's7boost', 'sales_channel', 'sales_interaction_type', 'sms']

    for i in range(len(names)):
        df[names[i]] = [(a_dict[keys[i]] if keys[i] in a_dict else None) for a_dict in list_of_params]

    return df

def process_file(filename):
    # читаем файл
    print(timestamp(), f'reading file {filename}')
    df = pd.read_excel(filename)
    print(timestamp(), 'checking and fixing header')
    df = fix_header(df)

    # переименовываем колонки, которые будем изменять
    rename_fileds = {df.columns[0]: "xResponse ID", df.columns[1]: "IP Address", df.columns[4]: "User Agent",
                     df.columns[5]: "Time Started", df.columns[10]: "Вы покупали билет самостоятельно?",
                     df.columns[11]: "Насколько легко было купить билет на рейс S7 Airlines?",
                     df.columns[12]: "Любопытно. Помогите нам провести работу над ошибками! Расскажите, почему вы поставили такую оценку?",
                     df.columns[13]: "Оценка полёта", df.columns[14]: "NPS", df.columns[15]: "Обоснование оценок",
                     df.columns[16]: "Оценка WWW", df.columns[17]: "Как вам работа нашего приложения?",
                     df.columns[18]: "Оценка аэропорта вылета", df.columns[19]: "Оценка на борту",
                     df.columns[20]: "Оценка аэропорта прилёта", df.columns[21]: "Покупка билетов:Детали WWW",
                     df.columns[22]: "Оформление доп. услуг:Детали WWW", df.columns[23]: "Онлайн-регистрация:Детали WWW",
                     df.columns[24]: "Удобство личного кабинета:Детали WWW",
                     df.columns[25]: "Покупка билетов:Давайте поговорим подробнее про некоторые детали вашего путешествияКак вам наше приложение?",
                     df.columns[26]: "Оформление доп. услуг:Давайте поговорим подробнее про некоторые детали вашего путешествияКак вам наше приложение?",
                     df.columns[27]: "Онлайн-регистрация:Давайте поговорим подробнее про некоторые детали вашего путешествияКак вам наше приложение?",
                     df.columns[28]: "Удобство личного кабинета:Давайте поговорим подробнее про некоторые детали вашего путешествияКак вам наше приложение?",
                     df.columns[29]: "Сотрудники компании доброжелательные? :Детали аэропорта вылета",
                     df.columns[30]: "Как вам обслуживание на стойке регистрации?:Детали аэропорта вылета",
                     df.columns[31]: "Навигация в аэропорту понятная? :Детали аэропорта вылета",
                     df.columns[32]: "Что скажете про обслуживание при посадке и доставке на самолёт?:Детали аэропорта вылета",
                     df.columns[33]: "Как вам работа сотрудников в условиях задержки рейса?:Детали аэропорта вылета",
                     df.columns[34]: "Информация о процедурах стыковки ясная? :Детали аэропорта вылета",
                     df.columns[35]: "Расскажите, что вы думаете про аэропорт вылета {ap1}? Сотрудники компании доброжелательные?",
                     df.columns[36]: "Расскажите, что вы думаете про аэропорт вылета {ap1}? Навигация в аэропорту понятная?",
                     df.columns[37]: "Расскажите, что вы думаете про аэропорт вылета {ap1}? Что скажете про обслуживание при посадке и доставке на самолёт?",
                     df.columns[38]: "Расскажите, что вы думаете про аэропорт вылета {ap1}? Как вам обслуживание на стойке регистрации?",
                     df.columns[39]: "Расскажите, что вы думаете про аэропорт вылета {ap1}? Как вам работа сотрудников в условиях задержки рейса?",
                     df.columns[40]: "Расскажите, что вы думаете про аэропорт вылета {ap1}? Информация о процедурах стыковки ясная?",
                     df.columns[41]: "Вы сказали, что сотрудники авиакомпании были недостаточно доброжелательны. Запишите подробнее, что и где именно было не так",
                     df.columns[42]: "В аэропорту вылета вам предлагали приобрести услугу по выбору места у окна или прохода (место в салоне, место повышенной комфортности, повышение класса в бизнес)?",
                     df.columns[43]: "Теперь расскажите про обслуживание и комфорт в салоне самолета подробнееВам было уютно на борту? Как вам дорожные наборы?",
                     df.columns[44]: "Теперь расскажите про обслуживание и комфорт в салоне самолета подробнееВам было уютно на борту? Вам понравилось питание и напитки?",
                     df.columns[45]: "Теперь расскажите про обслуживание и комфорт в салоне самолета подробнееВам было уютно на борту? Бортпроводники доброжелательные?",
                     df.columns[46]: "Теперь расскажите про обслуживание и комфорт в салоне самолета подробнееВам было уютно на борту? Пилот понятно озвучивал информацию?",
                     df.columns[47]: "Теперь расскажите про обслуживание и комфорт в салоне самолета подробнееВам было уютно на борту? Достаточно места для ручной клади и личных вещей?",
                     df.columns[48]: "Теперь расскажите про обслуживание и комфорт в салоне самолета подробнееВам было уютно на борту? А кресла удобные?",
                     df.columns[49]: "Теперь расскажите про обслуживание и комфорт в салоне самолета подробнееВам было уютно на борту? Температура на борту комфортная?",
                     df.columns[50]: "Теперь расскажите про обслуживание и комфорт в салоне самолета подробнееВам было уютно на борту? В салоне самолета было чисто?",
                     df.columns[51]: "Теперь расскажите про обслуживание и комфорт в салоне самолета подробнееВам было уютно на борту? Как вам дорожные наборы?",
                     df.columns[52]: "Теперь расскажите про обслуживание и комфорт в салоне самолета подробнееВам было уютно на борту? Вам понравилось питание и напитки?",
                     df.columns[53]: "Теперь расскажите про обслуживание и комфорт в салоне самолета подробнееВам было уютно на борту? Бортпроводники доброжелательные?",
                     df.columns[54]: "Теперь расскажите про обслуживание и комфорт в салоне самолета подробнееВам было уютно на борту? Пилот понятно озвучивал информацию?",
                     df.columns[55]: "Теперь расскажите про обслуживание и комфорт в салоне самолета подробнееВам было уютно на борту? Достаточно места для ручной клади и личных вещей?",
                     df.columns[56]: "Теперь расскажите про обслуживание и комфорт в салоне самолета подробнееВам было уютно на борту? А кресла удобные?",
                     df.columns[57]: "Теперь расскажите про обслуживание и комфорт в салоне самолета подробнееВам было уютно на борту? Температура на борту комфортная?",
                     df.columns[58]: "Теперь расскажите про обслуживание и комфорт в салоне самолета подробнееВам было уютно на борту? В салоне самолета было чисто?",
                     df.columns[59]: "Расскажите про сервис в аэропорту прилета {ap2} Как оцениваете доставку из самолета в аэропорт?",
                     df.columns[60]: "Расскажите про сервис в аэропорту прилета {ap2} Сотрудники компании доброжелательные?",
                     df.columns[61]: "Расскажите про сервис в аэропорту прилета {ap2} Багаж получили быстро?",
                     df.columns[62]: "Давайте поговорим про дополнительные услуги. Как вам услуга «Дополнительный багаж»?",
                     df.columns[63]: "Давайте поговорим про дополнительные услуги. Как оцените услугу «Повышение класса обслуживания (Апгрейд)»?",
                     df.columns[64]: "Давайте поговорим про дополнительные услуги. Что скажете про услугу «Место повышенной комфортности»?",
                     df.columns[65]: "Давайте поговорим про дополнительные услуги. Что думаете про услугу «Выбор места в самолете»?",
                     df.columns[66]: "Давайте поговорим про дополнительные услуги. Как вам услуга «Платное питание»?",
                     df.columns[67]: "Давайте поговорим про дополнительные услуги. Как вы оцениваете услугу «Бизнес-зал»?",
                     df.columns[68]: "Давайте поговорим про дополнительные услуги. Что скажете про услугу «Провоз питомца»?",
                     df.columns[69]: "Давайте поговорим про дополнительные услуги. Как оцените услугу «Еда к гейту (доставка из ресторана)»?",
                     df.columns[70]: "Давайте поговорим про дополнительные услуги. Что скажете по поводу услуги «Пересадка на ранний рейс»?",
                     df.columns[71]: "Давайте поговорим про дополнительные услуги. Что думаете по поводу услуги «Полный возврат»?",
                     df.columns[72]: "Давайте поговорим про дополнительные услуги. Как вам услуга «Абонемент S7 Boost»?",
                     df.columns[73]: "Давайте поговорим про дополнительные услуги. Что думаете по поводу услуги «Дополнительное место рядом» (Extra Seat)?",
                     df.columns[74]: "Давайте поговорим про дополнительные услуги. Как вы оцениваете услугу «Авторегистрация»?",
                     df.columns[75]: "Давайте поговорим про дополнительные услуги. Как оцените услугу «Платные SMS-оповещения»?",
                     df.columns[76]: "Как вам услуга «Дополнительный багаж»? Удобное оформление?",
                     df.columns[77]: "Как вам услуга «Дополнительный багаж»? Как вам соотношение цены и качества?",
                     df.columns[78]: "Как вам Повышение класса обслуживания (апгрейд)? Удобное оформление?",
                     df.columns[79]: "Как вам Повышение класса обслуживания (апгрейд)? Как вам соотношение цены и качества?",
                     df.columns[80]: "Что скажете про Место повышенной комфортности? Удобное оформление?",
                     df.columns[81]: "Что скажете про Место повышенной комфортности? Как вам соотношение цены и качества?",
                     df.columns[82]: "Как вам услуга «Выбор места в самолёте»? Удобное оформление?",
                     df.columns[83]: "Как вам услуга «Выбор места в самолёте»? Как вам соотношение цены и качества?",
                     df.columns[84]: "Что думаете про услугу «Платное питание»? Удобное оформление?",
                     df.columns[85]: "Что думаете про услугу «Платное питание»?	Как вам соотношение цены и качества?",
                     df.columns[86]: "Что скажете про услугу &quot;Бизнес-зал&quot;? Удобное оформление?",
                     df.columns[87]: "Что скажете про услугу &quot;Бизнес-зал&quot;? Как вам соотношение цены и качеств",
                     df.columns[88]: "Что скажете про услугу &quot;Провоз питомца&quot;? Описание услуги понятное?",
                     df.columns[89]: "Что скажете про услугу &quot;Провоз питомца&quot;? Удобное оформление?",
                     df.columns[90]: "Что скажете про услугу &quot;Провоз питомца&quot;? Как вам соотношение цены и качеств",
                     df.columns[91]: "Как оцените услугу &quot;Заказ еды к гейту (доставка из ресторана)&quot;?	Удобное оформление?",
                     df.columns[92]: "Как оцените услугу &quot;Заказ еды к гейту (доставка из ресторана)&quot;?	Как вам соотношение цены и качеств",
                     df.columns[93]: "Что скажете про услугу &quot;Пересадка на ранний рейс&quot;? Удобное оформление?",
                     df.columns[94]: "Что скажете про услугу &quot;Пересадка на ранний рейс&quot;? Как вам соотношение цены и качеств",
                     df.columns[95]: "Что думаете по поводу услуги «Полный возврат»? Удобное оформление?",
                     df.columns[96]: "Что думаете по поводу услуги «Полный возврат»? Как вам соотношение цены и качеств",
                     df.columns[97]: "А что насчёт Абонемента &quot;S7 Boost&quot;? Удобное оформление?",
                     df.columns[98]: "А что насчёт Абонемента &quot;S7 Boost&quot;? Как вам соотношение цены и качеств",
                     df.columns[99]: "А что насчёт Абонемента &quot;S7 Boost&quot;? Описание услуги понятное?",
                     df.columns[100]: "Как вам услуга «Авторегистрация»?	Удобное оформление?",
                     df.columns[101]: "Как вам услуга «Авторегистрация»?	Как вам соотношение цены и качеств",
                     df.columns[102]: "Что думаете по поводу услуги «Дополнительное место рядом (Extra Seat)»? Удобное оформление?",
                     df.columns[103]: "Что думаете по поводу услуги «Дополнительное место рядом (Extra Seat)»? Как вам соотношение цены и качеств",
                     df.columns[104]: "Как оцените услугу “Платные SMS-оповещения?” Удобное оформление?",
                     df.columns[105]: "Как оцените услугу “Платные SMS-оповещения?” Как вам соотношение цены и качеств",
                     df.columns[106]: "Нам очень помогут некоторые детали вашего путешествия.Основная цель вашей поездки была...",
                     df.columns[107]: "Вы обращались за помощью к сотрудникам авиакомпании?",
                     df.columns[108]: "Вы довольны тем, как сотрудники решили ваш вопрос в аэропорту?",
                     df.columns[109]: "Вы довольны тем, как сотрудники решили ваш вопрос по телефону?",
                     df.columns[110]: "Вы довольны тем, как сотрудники решили ваш вопрос в чате?",
                     df.columns[111]: "Расскажите подробнее: что произошло? К кому вы обратились?Как быстро вам помогли? Что думаете по этому поводу?",
                     df.columns[112]: "Мы очень хотим, чтобы в следующий раз вы тоже выбрали S7 😊Как думаете, что нам нужно изменить в лучшую сторону? "}

    df.rename(columns=rename_fileds, inplace=True)

    # меняет тип колонок со строчного на числовой
    print(timestamp(), 'changing columns to numeric type')

    print(timestamp(), 'making new ID column')
    df['Response ID'] = df['xResponse ID'].copy()
    # убирает все символы кроме цифр их "xResponse ID", обрезает до 8 символов, перемешивает и
    # записывая во фрейм прибавляет в начало "1"

    for indexes in range(len(df['xResponse ID'].index)):
        string = ''
        for char in df['xResponse ID'].loc[df.index[indexes]]:
            if char.isdecimal():
                string += char
        arr = list(string[:8])
        random.shuffle(arr)
        res_str = ''.join(arr)

        df['Response ID'].loc[df.index[indexes]] = '1' + res_str

    int_fields = ['Response ID', 'NPS']
    for field in int_fields:
        df[field] = pd.to_numeric(df[field])

    # меняем слова на цифры в колонке "Оценка на борту"
    df['Оценка на борту'] = df['Оценка на борту'].apply(fix_scoreOnBoard)

    list_of_params = parse_json(df)
    create_columns(df, list_of_params)


    # склеиваем колонки
    # print(timestamp(), 'making new Обоснование оценок column')
    # df['Обоснование оценок'] = df.apply(
    #     lambda x: combine_text(
    #         x['NPS'],
    #         x['xОбоснование оценок'],
    #         x['Что вам больше всего понравилось в этот раз?']), axis=1
    # )

    # создаем новое название файла
    print(timestamp(), 'making new filename')
    new_filename = "source/" + make_filename(filename)



    # задаем новый порядок колонок
    new_order = ['Response ID', 'NPS', 'Обоснование оценок', 'xResponse ID', 'Time Started', 'Date Submitted', 'Status', 'Language', 'User Agent', 'IP Address', 'Longitude', 'Latitude', 'Country', 'City', 'URL Variable: airport_destination',
                 'URL Variable: airport_origin',
                 'URL Variable: ap1', 'URL Variable: ap2', 'URL Variable: baggage', 'URL Variable: cabin',
                 'URL Variable: delay', 'URL Variable: flightdate', 'URL Variable: flightno', 'URL Variable: pax',
                 'URL Variable: paxtype',
                 'URL Variable: pnr', 'URL Variable: pnr_create_date', 'URL Variable: regchannel', 'URL Variable: spnr',
                 'URL Variable: transfer', 'URL Variable: anx_list', 'URL Variable: autoreg', 'URL Variable: earlyflight', 'URL Variable: emd',
                 'URL Variable: extraseat', 'URL Variable: gatefood', 'URL Variable: pet', 'URL Variable: refund',
                 'URL Variable: s7boost', 'URL Variable: sales_channel', 'URL Variable: sales_interaction_type',
                 'URL Variable: sms', 'who'] + list(df.columns[11:-42])



    # сохраняем в файл
    print(timestamp(), "saving file")
    df[new_order].to_excel(new_filename, index=False, sheet_name='Worksheet')
    print(timestamp(), f"file saved at {new_filename}")
    print('Finished')

    return new_filename



def cmd(command):
    print(timestamp(), "$", command)
    code = os.system(f'cmd /c "{command}"')  #
    return code


@argh.arg('filename', help='name of the raw file from e-mail e.g. "AFFB 04-10.04.22.xlsx"')
@argh.arg('start', help='start date in YYYY-MM-DD format, eg. 2022-04-04')
@argh.arg('end', help='start  +1 day in YYYY-MM-DD format, eg. 2022-04-10')
def main(filename, start, end):
    # cmd("svn up")
    new_name = process_file("source/" + filename)
    # cmd(f"svn add {new_name}")
    # cmd("svn commit")

    # first = cmd(f"java -Xmx4g -jar dist/SurveyLoader.jar config_svm.json --dbl {new_name}")
    # if first not in [0, 1]:
    #     raise Exception(f"ran with exit code {first}")
    #
    # second = cmd(f"java -Xmx4g -jar dist/SurveyLoader.jar config_svm.json --clsx {start} {end}")
    # if second not in [0, 1]:
    #     raise Exception(f"ran with exit code {second}")
    #
    url = f"http://prod3.infoqubes.ru:58080/S7MKS/process?col=60223dafc73e852cdb038005&cat=5ef5f9ddc73e852cdb6fbb56&from={start}&to={end}"
    print(timestamp(), f"Sending request to {url} ")
    # res = requests.get(url)
    # print(timestamp(), res.json())
    print(timestamp(), 'check link http://prod3.infoqubes.ru:58080/S7MKS/tasks for progress')


if __name__ == '__main__':
    argh.dispatch_command(main)
    #main('AFFB 02-08.10.2023.xlsx', '2023-05-22', '2023-05-29')
