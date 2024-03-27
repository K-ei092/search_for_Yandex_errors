from urllib.parse import unquote

from pylanguagetool import api
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromiumService
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager

from result_in_xlsx import save_in_file

import logging, os, time

# Инициализируем логгер
logger = logging.getLogger(__name__)


class Parser:

    # конструктор вебдрайвера
    def __init__(self):

        self.count = 1
        self.first_name_scrin = ''
        self.full_name_scrin = ''

        option = webdriver.ChromeOptions()
        option.add_argument('--disable-blink-features=AutomationControlled')  # отключение автодетекта браузера
        self.driver = webdriver.Chrome(options=option, service=ChromiumService(ChromeDriverManager().install()))
        self.driver.maximize_window()

    # служебный статметод для приведения текста к шаблону
    @staticmethod
    def _extract_words(text_: str) -> str:

        words_list = list(set([word.lower() for word in text_.split() if 4 <= len(word) <= 12 and word.isalpha()]))
        clean_text = '; '.join(words_list)
        result = 'Проверить слова: ' + clean_text
        return result

    # метод для перехода на сайт Яндекса
    def go_to_yandex(self):

        self.driver.get("https://www.google.com/webhp?hl=ru&sa=X&ved=0ahUKEwjc5Lrv8MCEAxXDi8MKHfpCDfoQPAgJ")
        self.driver.find_element(By.TAG_NAME, 'textarea').send_keys('яндекс поиск', Keys.ENTER)  # вводим в поиск запрос
        self.driver.find_element(By.TAG_NAME, 'h3').click()  # кликаем по первой выдаче - ссылка на яндекс поиск
        logger.info('Переход на Яндекс поиск')
        time.sleep(5)

    # метод для поиска на Яндексе
    def search_on_yandex(self, search_phrase: str) -> None:

        self.first_name_scrin = '_'.join(search_phrase.split()) + '_'
        self._search_captcha()
        self._search_alert()

        try:
            self.driver.find_element(By.CSS_SELECTOR,  # поиск и очистка поискового поля яндекса
                                     'input.HeaderDesktopForm-Input.'
                                     'mini-suggest__input[name="text"]').clear()
            self.driver.find_element(By.CSS_SELECTOR,  # вводим поисковую фразу
                                     'input.HeaderDesktopForm-Input.'
                                     'mini-suggest__input[name="text"]').send_keys(search_phrase, Keys.ENTER)
        except Exception:
            pass

        try:
            self.driver.find_element(By.CSS_SELECTOR,  # поиск и очистка поискового поля яндекса
                                     'input.search3__input.mini-'
                                     'suggest__input[name="text"]').clear()
            self.driver.find_element(By.CSS_SELECTOR,  # вводим поисковую фразу
                                     'input.search3__input.mini-'
                                     'suggest__input[name="text"]').send_keys(search_phrase, Keys.ENTER)
        except Exception:
            pass

        for i in [1, 2]:
            self._search_cards(pagination=i)
            logger.info(f'Пагинация {i} обработана, имя последнего скрина с ошибкой = {self.full_name_scrin}')

    # служебный метод для поиска и обработки каждой карточки на сайте и формирования результата
    def _search_cards(self, pagination: int, _extract_words=_extract_words):

        self._search_captcha()
        self._search_alert()

        time.sleep(0.5)
        cards = self.driver.find_elements(By.CSS_SELECTOR, 'li.serp-item.serp-item_card')  # '#search-result > li')

        for card in cards:
            elem_card = card.find_elements(By.CSS_SELECTOR, '*')
            result_text_card = ''
            flag = True
            for elem in elem_card:

                try:
                    if flag:
                        elem.find_element(By.CSS_SELECTOR, 'span.Link.Link_theme_neoblue.ExtendedText-Toggle').click()
                        flag = False
                except:
                    pass

                try:
                    elem.find_element(By.CSS_SELECTOR, 'button.Link.Link_theme_normal.CoveredPhone-Button').click()
                except:
                    pass

                try:
                    elem_text = ' '.join(elem.text.replace('\n', ' ').split())
                    result_text_card += elem_text
                except:
                    pass

            text = _extract_words(text_=result_text_card)

            try:

                tl = api.check(input_text=text, api_url='https://api.languagetool.org/v2/', lang='auto')

                if tl['matches'] and \
                        tl['matches'][0]['replacements'] and \
                        tl['matches'][0]['shortMessage'] == 'Орфографическая ошибка' and \
                        len(tl['matches'][0]['replacements']) == 1 and \
                        len(tl['matches'][0]['replacements'][0]['value'].split()) == 1 and \
                        tl['matches'][0]['replacements'][0]['value'] not in ['супермаркет']:
                    context = tl['matches'][0]['context']['text']
                    correction = tl['matches'][0]['replacements'][0]['value']
                    link = unquote(card.find_element(By.TAG_NAME, 'a').get_attribute("href"))
                    self.full_name_scrin = self.first_name_scrin + str(self.count) + '.png'

                    while True:
                        res = self.driver.save_screenshot(self.full_name_scrin)
                        logger.info('сохраняем скрин')
                        if res:
                            break

                    time.sleep(0.2)
                    save_in_file(context=context, correction=correction, link=link, name_scrin=self.full_name_scrin)
                    os.remove(self.full_name_scrin)

            except IndexError:
                pass
            except Exception:
                logger.exception('Ошибка:')

            self.driver.execute_script("return arguments[0].scrollIntoView(true);", card)

            self.count += 1
            time.sleep(4)

        # переходим на вторую страницу пагинации
        if pagination == 1:
            self.driver.find_element(By.CSS_SELECTOR,
                                     'a.VanillaReact.Pager-Item.Pager-Item_type_page').click()
            time.sleep(2)
            logger.info(f'переход на вторую страницу поиска')

    # служебный метод проверки наличия каптчи и ожидание её решения пользователем
    def _search_captcha(self):

        try:
            element_captcha = self.driver.find_element(By.CSS_SELECTOR,
                                                       "div.Theme.Theme_color_yandex-default.Theme_root_default")
            while element_captcha:
                element_captcha = self.driver.find_element(By.CSS_SELECTOR,
                                                           "div.Theme.Theme_color_yandex-default.Theme_root_default")
                time.sleep(3)
        except Exception:
            pass

    # служебный метод закрытие алерта "сделать Яндекс Браузер основным"
    def _search_alert(self):

        time.sleep(2)
        try:
            self.driver.find_element(By.CSS_SELECTOR,
                                     'button.Button2.Button2_pin_circle-circle.Button2_size_l.'
                                     'DistributionSplashScreenModal-CloseButtonOuter').click()
        except Exception:
            pass

        try:
            self.driver.find_element(By.CSS_SELECTOR, 'button.simple-popup__close.').click()
        except Exception:
            pass

    # закрытие драйвера
    def close_driver(self):

        if self.driver:
            self.driver.quit()
            logger.info(f'Экземпляр вэбдрайвера удалён')
