from urllib.parse import unquote
from pylanguagetool import api

from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromiumService
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

from result_in_xlsx import save_in_file

import time, logging

# Инициализируем логгер
logger = logging.getLogger(__name__)


class Parser:

    # статметод для приведения текста к шаблону
    @staticmethod
    def _extract_words(text_: str) -> str:

        words_list = list(set([word.lower() for word in text_.split() if 4 <= len(word) <= 12 and word.isalpha()]))
        clean_text = '; '.join(words_list)
        result = 'Проверить слова: ' + clean_text
        return result

    # конструктор вебдрайвера
    def __init__(self):

        option = webdriver.ChromeOptions()
        option.add_argument('--disable-blink-features=AutomationControlled')  # отключение автодетекта браузера
        self.driver = webdriver.Chrome(options=option, service=ChromiumService(ChromeDriverManager().install()))
        self.driver.maximize_window()

    # служебный метод для поиска и обработки каждой карточки на сайте и формирования результата
    def _search_cards(self, name_scrin: str, count: int, pagination: int, _extract_words=_extract_words) -> int | None:

        count = count

        # закрытие алерта "сделать Яндекс Браузер основным"
        try:
            self.driver.find_element(By.XPATH, '/html/body/main/div[3]/button').click()
        except Exception:
            pass
            # logger.exception('Алерт "сделать Яндекс Браузер основным" не найден')

        cards = self.driver.find_elements(By.CSS_SELECTOR, '#search-result > li')

        for card in cards:
            elem_card = card.find_elements(By.XPATH, './/*')
            text = ''
            for elem in elem_card:
                try:
                    text += elem.text.replace('\n', ' ') + ' '
                    if elem.text == 'Читать ещё':
                        elem.click()
                    if 'Показать' in elem.text.split() and '+7' in elem.text.split():
                        elem.click()
                    logger.info(f'это текст элемента: {elem.text}')
                except Exception:
                    pass
                    # logger.exception('В карточке не найден элемент или текст')
            text = _extract_words(text_=text)

            try:
                tl = api.check(input_text=text, api_url='https://api.languagetool.org/v2/', lang='auto')
                if tl['matches'] and \
                        tl['matches'][0]['replacements'] and \
                        tl['matches'][0]['shortMessage'] == 'Орфографическая ошибка' and \
                        len(tl['matches'][0]['replacements']) == 1 and \
                        len(tl['matches'][0]['replacements'][0]['value'].split()) == 1 and \
                        tl['matches'][0]['replacements'][0]['value'] not in ['супермаркет']:
                    link = unquote(card.find_element(By.TAG_NAME, 'a').get_attribute("href"))
                    correction = tl['matches'][0]['replacements'][0]['value']
                    context = tl['matches'][0]['context']['text']
                    name_scrin = name_scrin + str(count) + '.png'
                    self.driver.save_screenshot(name_scrin)
                    save_in_file(context=context, correction=correction, link=link, name_scrin=name_scrin)
            except Exception:
                logger.exception('Ошибка:')

            self.driver.execute_script("return arguments[0].scrollIntoView(true);", card)

            count += 1
            time.sleep(6)

        if pagination == 1:
            self.driver.find_element(By.XPATH,  # переходим на вторую страницу пагинации
                                     '/html/body/main/div[2]/div[2]/div/div/div[1]/nav/div/div[2]/a').click()
            time.sleep(4)
            return count
        else:
            return None

    # метод для перехода на сайт Яндекса
    def go_to_yandex(self):

        self.driver.get("https://www.google.com/webhp?hl=ru&sa=X&ved=0ahUKEwjc5Lrv8MCEAxXDi8MKHfpCDfoQPAgJ")
        self.driver.find_element(By.TAG_NAME, 'textarea').send_keys('яндекс поиск', Keys.ENTER)  # вводим в поиск запрос
        self.driver.find_element(By.TAG_NAME, 'h3').click()  # кликаем по первой выдаче - ссылка на яндекс поиск
        logger.info('Переход на Яндекс поиск')
        time.sleep(10)

    # метод для поиска на Яндексе
    def search_on_yandex(self, timer_capcha: int, search_phrase: str = '', _search_cards=_search_cards) -> None:

        count = 1
        name_scrin = '_'.join(search_phrase.split()) + '_'

        self.driver.find_element(By.TAG_NAME, 'input').clear()
        self.driver.find_element(By.TAG_NAME, 'input').send_keys(search_phrase, Keys.ENTER)

        if timer_capcha == 0:
            time.sleep(30)  # время для ручного решения капчи при первом входе (если она появится)

        for i in range(1, 3):
            count = self._search_cards(name_scrin=name_scrin, count=count, pagination=i)
            logger.info(f'Пагинация {i} обработана')

    # закрытие драйвера
    def close_driver(self):
        if self.driver:
            self.driver.quit()
            logger.info(f'Экземпляр вэбдрайвера удалён')
