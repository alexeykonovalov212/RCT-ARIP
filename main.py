def main():
    import pandas as pd
    import openpyxl
    import html
    import numpy as np
    import requests
    from bs4 import BeautifulSoup
    from selenium import webdriver
    from selenium.webdriver.common.keys import Keys
    import time
    import re


    sleep = 4

    url = 'https://easuz.mosreg.ru/Arip/#/ObjectPurchaseSearch'

    driver = webdriver.Edge('C:\\Users\\KonovalovAlE\\Desktop\\jupyter notebook\\Вебдрайвер edge\\msedgedriver.exe')

    driver.get(url)
    time.sleep(sleep)

    login = driver.find_elements_by_id("loginBox.username")        # ищем id поля login, вводим данные, нажимаем click
    login[0].send_keys("KonovalovAlE@mosreg.ru")
    login[0].click()

    login = driver.find_elements_by_id("loginBox.password")        # ищем id поля password, вводим данные, нажимаем click
    login[0].send_keys("2Jr6m53xY")
    login[0].click()
    time.sleep(2)

    button = driver.find_elements_by_tag_name("button")        # ищем tag кнопки, нажимаем click, закрываем окно
    button[2].click()

    #button = driver.find_elements_by_tag_name("button")        # Закрываем всплывающее окно с информированием, возможно придётся отключить когда отключат информирование на сайте
    #button[0].click()


    list_kn = ['50:03:0070209:644', '50:08:0070354:678']

    # Помещение '50:30:0000000:18014', '50:38:0000000:5543', '50:30:0000000:18017'
    # 50:16:0103041:727 , 50:26:0110715:685 , 50:28:0090312:510 '50:09:0060210:659'

    def parse_arip(kad_num):
        time.sleep(2)
        kn = driver.find_elements_by_xpath('/html/body/div[2]/div/div[2]/div/div[3]/div[2]/div[4]/div[1]/input') # вставляем кадастровый номер в поле
        time.sleep(sleep)

        kn[0].send_keys(kad_num)
        kn[0].click()



        button = driver.find_elements_by_tag_name("button")     # Нажимаем кнопку Поиск
        time.sleep(sleep)
        click = ''
        for i in button:
            if 'Поиск' in i.get_attribute('innerHTML'):
                click = i
        click.click()
        time.sleep(sleep)

        #find_element_by_xpath("//*").get_attribute("outerHTML")
        pagesource = driver.page_source           # Ищем ссылку на объект торгов
        time.sleep(sleep)
        soup = BeautifulSoup(pagesource, features='lxml')
        post = soup.find_all('a')
        href = []
        for i in post:
            if '#/ObjectPurchase' in str(i):
                href.append(i.get('href'))
        link = 'https://easuz.mosreg.ru/Arip/' + href[2]
        driver.get(link)
        time.sleep(sleep)

        page = driver.find_element_by_xpath("//*").get_attribute("outerHTML")               # Получаем код страницы объекта

        soup = BeautifulSoup(page, features='lxml')

        # Описание объекта
        object_type = soup.find_all('div', class_='col-lg-4 col-xs-4 ng-binding')[0].text.strip()                   # Описание объекта


        # Местоположение
        try:
            place = soup.find('div', class_='col-lg-10 col-xs-10 ng-binding').text.split('(')[1].split(')')[0]      # Местоположение

        except:
            place = soup.find('div', class_='col-lg-10 col-xs-10 ng-binding').text.strip()


        # Кадастровый номер
        kn = soup.find('div', class_='col-lg-2 col-xs-2 ng-binding').text.strip()                                   # Кадастровый номер


        # Площадь
        square = soup.find('div', class_='col-lg-1 col-xs-1 ng-binding').text.strip().replace(' ','')                           # Площадь
        if 'Помещение' in object_type:
            square = soup.find_all('div', class_='col-lg-1 col-xs-1 ng-binding')[2].text.strip().replace(' ','')


        # Дата входящего обращения ОМСУ
        date_omsu = soup.find_all('div', class_='col-lg-4 col-xs-4 ng-binding')[-2].text.strip()[0:10]          # Дата входящего обращения ОМСУ
        if 'Помещение' in object_type:
            date_omsu = soup.find_all('div', class_='col-lg-4 col-xs-4 ng-binding')[-1].text.strip()[0:10]
            print(date_omsu)


        # Реестровый номер объекта АРИП
        number_arip = soup.find_all('div', class_='col-lg-4 col-xs-4 ng-binding')[3].text.strip()               # Реестровый номер объекта АРИП


        # Вопрос протокола МВК
        prot_mvk = soup.find_all('div', class_='col-lg-4 col-xs-4 ng-binding')[-1].text.strip()                 # Вопрос протокола МВК
        if 'Помещение' in object_type:
            prot_mvk = np.NaN

        # Номер входящего обращения ОМСУ
        omsu_num = soup.find_all('div', class_='col-lg-4 col-xs-4 wb ng-binding')[3].text.strip()               # Номер входящего обращения ОМСУ


        # Дата протокола МВК
        date_mvk = ''
        try:
            date_mvk = soup.find_all('div', class_='col-lg-4 col-xs-4 ng-binding ng-scope')[1].text.strip()[0:10]   # Дата протокола МВК

        except:
            date_mvk = ''

        # Номер протокола МВК
        mvk_num = ''
        try:
            mvk_num = soup.find_all('div', class_='col-lg-4 col-xs-4 wb ng-binding')[4].text.strip()                # Номер протокола МВК

        except:
            mvk_num = ''

        # Кол-во дней с МВК
        try:
            mvk_day = soup.find_all('div', class_='col-lg-4 col-xs-4 wb ng-binding')[5].text.strip()            # Кол-во дней с МВК
        except:
            mvk_day = np.NaN



        # Переходим в план торгов
        button = driver.find_elements_by_tag_name('a')
        for i in button:
            if 'PlanTorgObject' in str(i.get_attribute('outerHTML')):
                button = i
        button.click()
        time.sleep(sleep)

        # Получаем код страницы объекта в плане торгов
        page = driver.find_element_by_xpath("//*").get_attribute("outerHTML")
        soup = BeautifulSoup(page, features='lxml')

        # Вид торгов
        vid_torg = soup.find_all('div', class_='col-lg-4 col-xs-4 ng-binding')[1].text.strip().lower()            # Вид торгов


        # ВРИ/Функциональное назначение
        vri = soup.find_all('div', class_='col-lg-4 col-xs-4 ng-binding')[6].text.strip()                          # ВРИ/Функциональное назначение
        if '(2.2)' in vri:
            vri = vri.replace('(2.2)','').strip()

        # НМЦ
        nmc = soup.find_all('div', class_='col-lg-4 col-xs-4 ng-binding')[4].text.strip()                          # НМЦ


        ################################################################################################################
        # Только для аренды!!!!!!!!!!!!!!!!!!!!!!!!!!

        # Пробуем достать количество лет аренды, иначе пусто
        label = soup.find_all('span', class_='k-numeric-wrap k-state-default k-expand-padding')
        try:
            year = str(label[1])[str(label[1]).find('aria-valuenow')+15:str(label[1]).find('aria-valuenow')+17]
        except:
            year = np.NaN


        # Пробуем достать количество месяцев
        try:
            month = str(label[2])[str(label[2]).find('aria-valuenow')+15:str(label[2]).find('aria-valuenow')+17]
            if month == '0':
                month = np.NaN
        except:
            month = np.NaN


        # Пробуем достать количество дней
        try:
            day = str(label[3])[str(label[3]).find('aria-valuenow')+15:str(label[3]).find('aria-valuenow')+16]
            if day == '0':
                day = np.NaN
        except:
            day = np.NaN

        ###############################################################################################################

        # Реестровый номер процедуры АРИП
        arip_proc = soup.find_all('a')                                                                              # Реестровый номер процедуры АРИП
        for i in arip_proc:
            if 'purchaseId' in str(i):
                arip_proc = i.text.strip()


        # Дата публикации
        pub_date = np.NaN
        try:
            pub_date = soup.find_all('span', class_='ng-binding')[4].text.strip()[0:10]                                 # Дата публикации

        except:
            pass

        # Дата начала приема заявок
        start_date = soup.find_all('span', class_='ng-binding')[5].text.strip()[0:10]                               # Дата начала приема заявок


        # Дата окончания приема заявок
        end_date = soup.find_all('span', class_='ng-binding')[5].text.strip()[19:29]                                # Дата окончания приема заявок


        # Дата аукциона
        auc_date = soup.find_all('span', class_='ng-binding')[6].text.strip()[0:10]                                 # Дата аукциона


        # Количество лотов

        num_lot = ''
        try:
            num_lot = soup.find_all('span', class_='ng-binding')[7].text.strip()                                        # Количество лотов

        except:
            num_lot = ''


        ######################################################################################################
        # Пробую переходить в реестр торгов

        driver.back()
        time.sleep(sleep)
        driver.back()
        time.sleep(sleep)
        button = driver.find_element_by_id("dtrog")
        button.click()
        time.sleep(sleep)
        pagesource = driver.page_source
        time.sleep(sleep)
        soup = BeautifulSoup(pagesource, features='lxml')
        a = soup.find_all('a')
        link = 'https://easuz.mosreg.ru/Arip/'
        for i in a:
            if 'class="ng-binding" href="#/Purchase/' in str(i):
                link_1 = i.get('href')
                link = link + link_1

        driver.get(link)
        time.sleep(sleep)
        page = driver.page_source
        soup = BeautifulSoup(page, features='lxml')
        time.sleep(sleep)

        # Получаем все данные с классом, их более 50 и в них содержится вся нужная инфа, получим их по индексу
        data_auc = soup.find_all('div', attrs={'class': 'col-lg-3 col-xs-3 ng-binding'})
        #for i in enumerate(data_auc):
        #    print(i)

        # Номер аукциона
        num_auc = data_auc[25].text.strip()                                                                             # Номер аукциона
        if 'Земельный участок' in num_auc:
            num_auc = data_auc[23].text.strip()
        elif 'Аукцион' in num_auc:
            num_auc = data_auc[22].text.strip()
        elif 'Право на заключение договора аренды имущества' in num_auc:
            num_auc = data_auc[23].text.strip()

        # Номер на ГИСах
        num_gis = data_auc[20].text.strip()                                                                             # Номер на ГИСах
        if 'Торги опубликованы' in num_gis:
            num_gis = data_auc[18].text.strip()
        elif 'Прием заявок' in num_gis:
            num_gis = data_auc[17].text.strip()

        # Размер задатка
        zadatok = data_auc[13].text.strip()                                                                             # Размер задатка
        test = ['1','2','3','4']
        for i in test:
            if zadatok == i:
                zadatok = data_auc[12].text.strip()
        if zadatok == '':
            zadatok = data_auc[12].text.strip()
        if zadatok == np.NaN:
            zadatok = data_auc[12].text.strip()

        # Реестровый номер процедуры АРИП (Реестр торгов)
        arip_proc = data_auc[18].text.strip()                                                                           # Реестровый номер процедуры АРИП (Реестр торгов)
        if 'Торги объявлены' in arip_proc:
            arip_proc = data_auc[15].text.strip()
        elif arip_proc == '':
            arip_proc = data_auc[16].text.strip()
        if 'ПЗЭ' in num_auc:
            arip_proc = data_auc[16].text.strip()

        time.sleep(2)
        ######################################################################################################


        url = 'https://easuz.mosreg.ru/Arip/#/ObjectPurchaseSearch'                                                     # Снова перехожу на поисковую страницу АРИП
        driver.get(url)
        time.sleep(sleep)

        ################
        driver.find_elements_by_class_name('hide_block')[0].click()                                                     # Тут я открываю строку с поиском
        driver.find_elements_by_xpath('//*[@id="main"]/div/div[3]/div[2]/div[4]/div[1]/input')[0].clear()
        #######################

        list_data = [vid_torg,object_type, vri,place,kn,square,date_omsu,omsu_num,date_mvk,mvk_num,prot_mvk,mvk_day,number_arip,nmc,zadatok,year,month,day,pub_date,start_date,end_date,auc_date,num_auc,num_gis,arip_proc,auc_date]

        return list_data

    # Создаю датафрейм для данных и заполняю его
    index = ['Вид торгов','Описание объекта', 'Функциональное назначение / вид разрешенного использования земельного участка',
             'Адрес объекта', 'Краткая характеристика объекта/кадастровый номер', 'Площадь', 'Дата входящего обращения ОМСУ',
             'Номер входящего обращения ОМСУ', 'Дата протокола МВК', 'Номер протокола МВК', 'Вопрос протокола МВК', 'Количество дней с МВК',
             'Реестровый номер объекта в ЕАСУЗ АРИП', 'Начальная цена (предмет аукциона)', 'Размер задатка', 'Срок договора, лет', 'Срок договора, месяцев',
             'Срок договора, дней', 'Дата публикации', 'Дата начала заявочной кампании', 'Дата окончания заявочной кампании',
             'Дата рассмотрения заявок/определения участников', 'Номер аукциона', 'Номер извещения на Официальном сайте', 'Реестровый номер процедуры в АРИП', 'Дата аукциона']
    columns = list(range(0, len(list_kn)))
    data = pd.DataFrame(index=index, columns=columns)

    for column in data.columns:
        data[column] = parse_arip(list_kn[column])
        print(column)
    data = data.T

    # Заполнил таблицу данными, перевернул и теперь добавляю пустые столбцы как в ГТ

    # Тут названия колонок и индексы куда их нужно вставить для цикла

    col = ['№ п/п', 'Наименование городского округа, муниципального района или органа исполнительной власти',
           'Отметка об ограничении/Временно-учтенные','Код',
           'Тип объекта', 'Этажность', 'Начальная цена (вгод)', 'Документ, подтверждающий начальную цену предмета аукциона',
           'Дата документа', 'Срок актуальности', 'Номер лота/Кол-во объектов',
           'Место проведения аукциона', 'Реестровый номер процедуры на РТС-Тендер', 'Статус процедуры', 'Даты переносов аукциона',
           'Количество продлений заявочной кампании', 'Ответственный сотрудник (ГКУ РЦТ)', 'Дата направления отказа', 'Исходящий номер отказа',
           'Дата исходящего обращения','Номер исходящего обращения']
    ind = [0,1,2,3,7,9,20,25,26,27,33,36,37,39,40,41,42,43,44,45,46]

    for i in range(len(col)):
        data.insert(ind[i],col[i],np.NaN)

    # Меняю количество дней с МВК на NaN так как эта инфа неактуальна
    data['Количество дней с МВК'] = np.NaN

    # Ставлю НМЦ вгод как НМЦ (предмет аукциона), в дальнейшем нужно учесть если НМЦ в месяц умножать на 10 в аренде имущества
    data['Начальная цена (вгод)'] = data['Начальная цена (предмет аукциона)']

    # Ставлю Статус процедуры (Заявочная кампания)
    data['Статус процедуры'] = 'Заявочная кампания'

    # Место проведения аукциона ЭП
    data['Место проведения аукциона'] = 'ЭП'

    # Номер на ГИСах = Номер на РТС-Тендер
    data['Реестровый номер процедуры на РТС-Тендер'] = data['Номер извещения на Официальном сайте']

    # Из номера аукциона беру Код
    try:
        for i in range(len(data['Номер аукциона'])):
            data.loc[i,'Код'] = data.loc[i,'Номер аукциона'].split('-')[0].strip()
    except:
        pass

    # Меняю ВИД торгов в соответствии с кодом
    try:
        for i in range(len('Код')):
            if data.loc[i,'Код'] == 'АЗЭ':
                data.loc[i,'Вид торгов'] = 'аренда земельного участка в электронной форме'
            elif data.loc[i,'Код'] == 'АЗГЭ':
                data.loc[i, 'Вид торгов'] = 'аренда земельного участка для граждан в электронной форме'
            elif data.loc[i,'Код'] == 'ПЗЭ':
                data.loc[i, 'Вид торгов'] = 'продажа земельного участка в электронной форме'
            elif data.loc[i,'Код'] == 'АЗПЭ':
                data.loc[i, 'Вид торгов'] = 'аренда земельного участка в электронной форме для СМП'
            elif data.loc[i,'Код'] == 'ПЭ':
                data.loc[i, 'Вид торгов'] = 'продажа помещения/здания в электронной форме'
            elif data.loc[i,'Код'] == 'АЭ':
                data.loc[i, 'Вид торгов'] = 'аренда помещения/здания в электронной форме'
    except:
        pass


    # Добавляю ИЖС/ЛПХ
    for i in range(len(data['Функциональное назначение / вид разрешенного использования земельного участка'])):
        if 'ведения личного подсобного хозяйства' in data.loc[i,'Функциональное назначение / вид разрешенного использования земельного участка']:
            data.loc[i,'Тип объекта'] = 'ИЖС/ЛПХ'
        elif 'индивидуального жилищного строительства' in data.loc[i,'Функциональное назначение / вид разрешенного использования земельного участка']:
            data.loc[i, 'Тип объекта'] = 'ИЖС/ЛПХ'
        else:
            data.loc[i, 'Тип объекта'] = np.NaN

    # Добавляю городской округ

    city = ['Волоколамский г.о.','Воскресенский г.о.','Дмитровский г.о.','Зарайск г.о.','Истра г.о.','Клин г.о.',
    'Коломенский г.о.','Красногорск г.о.','Ленинский г.о.','Лотошино г.о.','Луховицы г.о.','Люберцы г.о.','Можайский г.о.',
    'Наро-Фоминский г.о.','Богородский г.о.','Одинцовский г.о.','Орехово-Зуевский г.о.','Павловский Посад г.о.','Пушкинский г.о.',
    'Раменский г.о.','Рузский г.о.','Сергиево-Посадский г.о.','Солнечногорск г.о.','Ступино г.о.','Талдомский г.о.','Чехов г.о.',
    'Шатура г.о.','Щелково г.о.','Балашиха г.о.','Бронницы г.о.','Власиха г.о.', 'Долгопрудный г.о.',
    'Домодедово г.о.','Дубна г.о.','Егорьевск г.о.','Жуковский г.о.','Звёздный городок г.о.', 'Кашира г.о.',
    'Королёв г.о.','Котельники г.о.','Краснознаменск г.о.','Лобня г.о.','Лосино-Петровский г.о.',
    'Лыткарино г.о.','Молодежный г.о.','Мытищи г.о.','Подольск г.о.','Протвино г.о.','Пущино г.о.','Реутов г.о.',
    'Серебряные Пруды г.о.','Серпухов г.о.','Фрязино г.о.','Химки г.о.','Черноголовка г.о.','Шаховская г.о.','Электрогорск г.о.','Электросталь г.о.']

    code_city = ['ВОЛ', 'ВОС','ДМ','ЗР','ИСТР','КЛН','ГОКО','КР','ЛЕН','ЛОТ','ЛУХ','ЛЮБ','МОЖ','НФ','БГР','ОД','ОЗГО','ПП','ПУШ','РАМ','РУЗ',
                 'СП','СГ','СТУ','ТЛ','ЧЕХ','ШАТ','ЩЕЛК','БАЛ','БР','ВЛ','ДП','ДО','ДУБ','ЕГ','ЖУК','ЗВГ','КАШ','КОР',
                 'КОТ','КЗН','ЛОБ','ЛП','ЛЫТ','МОЛ','МЫТ','ПДЛГО','ПР','ПУЩ','РЕУ','СЕР','СРП','ФР','ХИМ','ЧГ','ШАХ','ЭГ','ЭС']
    try:
        for i in range(len(data['Номер аукциона'])):
            data.loc[i, 'Наименование городского округа, муниципального района или органа исполнительной власти'] = city[code_city.index(data.loc[i, 'Номер аукциона'].split('-')[1].strip().split('/')[0].strip())]
    except:
        pass

    ######################################################################################

    file = 'C:\\Users\\KonovalovAlE\\Desktop\\Парсинг Арип\\zayavki.xlsx'  # указываем здесь путь к файлам
    wb = openpyxl.load_workbook(file)
    ws = wb.active

    from openpyxl.utils.dataframe import dataframe_to_rows
    rows = dataframe_to_rows(data, index=False)

    for r_idx, row in enumerate(rows, 1):
        for c_idx, value in enumerate(row, 1):
            ws.cell(row=r_idx, column=c_idx, value=value)

    wb.save(file)

    print('Готово!')

if __name__ == "__main__":
    main()