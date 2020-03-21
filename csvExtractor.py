import logging.config
import os
import sys
from xlrd import open_workbook

LOG_CONF = "./logging.conf"
logging.config.fileConfig(LOG_CONF)

from kivy.app import App
from kivy.config import Config

Config.set('modules', 'inspector', '')  # Inspectorを有効にする
Config.set('graphics', 'width', 480)
Config.set('graphics', 'height', 280)
Config.set('graphics', 'maxfps', 20)  # フレームレートを最大で20にする
Config.set('graphics', 'resizable', 0)  # Windowの大きさを変えられなくする
Config.set('input', 'mouse', 'mouse,disable_multitouch')
from kivy.core.text import LabelBase, DEFAULT_FONT
from kivy.core.window import Window
from kivy.resources import resource_add_path
from kivy.uix.screenmanager import Screen

if hasattr(sys, "_MEIPASS"):
    resource_add_path(sys._MEIPASS)

EMPTY = ""
INDEX_TWITTER = 16
INDEX_EMAIL = 8
INDEX_FIRST_NAME = 3
INDEX_LAST_NAME = 2
INDEX_ADDRESS = 7
INDEX_POST_CODE = 4
INDEX_CITY = 6
INDEX_REGION = 5
INDEX_PHONE = 8
INDEX_PAY_TYPE = 13
INDEX_CARD_NUMBER = 14
INDEX_CARD_LIMIT_MONTH = 15
INDEX_CARD_LIMIT_YEAR = 16
INDEX_CARD_CVV = 17
INDEX_ITEM_NO = 4
INDEX_ITEM_SIZE = 2

ID_MESSAGE = "message"

OUT_FILE_NAME = "result.csv"
UTF8 = "utf8"
SJIS = "sjis"

EXT_TXT = ".txt"
EXT_XLS = ".xlsx"
EXT_XLSX = ".xlsx"
EXT_XLSM = ".xlsm"

KEY_ORDER_NUMBER = "注文オーダー番号"
KEY_MAIL_ADDRESS = "購入者メールアドレス"
KEY_ITEM_NAME = "商品名"

CONFIG_TXT = "./config.txt"
CONFIG_DICT = {}
CONFIG_KEY_OUTPUT_CSV_NAME = "OUTPUT_CSV_NAME"
CONFIG_KEY_OUTPUT_CSV_CHAR_SET = "OUTPUT_CSV_CHAR_SET"
CONFIG_KEY_INPUT_TEXT_CHAR_SET = "INPUT_TEXT_CHAR_SET"
ORDER_ITEM_LIST_DICT = {}
ORDER_NUM_DICT = {}
TWITTER_DICT = {}

excel_proc_line_num = 0
text_proc_line_num = 0
already_read_text = False
already_read_excel = False


class MainScreen(Screen):
    def __init__(self, **kwargs):
        super(MainScreen, self).__init__(**kwargs)
        self._file = Window.bind(on_dropfile=self._on_file_drop)

    def _on_file_drop(self, window, file_path):
        file_path = file_path.decode(UTF8)
        root, ext = os.path.splitext(file_path)
        if ext == EXT_TXT:
            self.parse_text_file(file_path)
        elif ext in [EXT_XLS, EXT_XLSM, EXT_XLSX]:
            self.parse_excel_file(file_path)

        if already_read_text and already_read_excel:
            self.dump_csv()

    def dump_csv(self):
        out_file_name = CONFIG_DICT[CONFIG_KEY_OUTPUT_CSV_NAME]
        try:
            self.dump_csv_core()
        except Exception as e:
            err_msg = "{}の出力に失敗しました。".format(out_file_name)
            self.disp_messg_err(err_msg)
            log.exception(err_msg, e)

    def dump_csv_core(self):
        out_file_name = CONFIG_DICT[CONFIG_KEY_OUTPUT_CSV_NAME]
        twitter_item_dict = self.mk_twitter_item_dict()
        self.dump_twitter_and_item_list(out_file_name, twitter_item_dict)
        self.disp_messg("{}を出力しました".format(out_file_name))

    @staticmethod
    def dump_twitter_and_item_list(out_file_name, twitter_item_dict):
        with open(out_file_name, "w", encoding=CONFIG_DICT[CONFIG_KEY_OUTPUT_CSV_CHAR_SET]) as f:
            for item in sorted(twitter_item_dict.items()):
                f.write("{}".format(item[0]))
                for item_name in sorted(item[1]):
                    f.write(",{}\n".format(item_name))

    @staticmethod
    def mk_twitter_item_dict():
        twitter_item_dict = {}
        for item in ORDER_ITEM_LIST_DICT.items():
            mail = item[0]
            order_item_list = item[1]
            twitter = TWITTER_DICT.get(mail)
            if twitter is None:
                log.warn("アドレス {} はExcelファイルに存在しません。処理をスキップします。".format(mail))
            else:
                list = twitter_item_dict.get(twitter, [])
                list.extend(order_item_list)
                twitter_item_dict[twitter] = list
        return twitter_item_dict

    def parse_excel_file(self, file_path):
        global excel_proc_line_num
        global already_read_excel
        try:
            parse_excel_file_core(file_path)
            already_read_excel = True
            self.disp_messg("{}を読み込みました。\n続いてテキストファイルをドラッグ&ドロップしてください".format(
                os.path.basename(file_path)))
        except Exception as e:
            file_name = os.path.basename(file_path)
            err_msg = "{}の読込処理に失敗しました。\nエラー発生行番号={}。".format(file_name, excel_proc_line_num)
            self.disp_messg_err(err_msg)
            log.exception(err_msg, e)
            already_read_excel = False

    def parse_text_file(self, file_path):
        global text_proc_line_num
        global already_read_text
        try:
            parse_text_file_core(file_path)
            already_read_text = True
            self.disp_messg("{}を読み込みました。\n続いてExcelファイルをドラッグ&ドロップしてください".format(
                os.path.basename(file_path)))
        except Exception as e:
            file_name = os.path.basename(file_path)
            err_msg = "{}の読込処理に失敗しました。\nエラー発生行番号={}。".format(file_name, text_proc_line_num)
            self.disp_messg_err(err_msg)
            log.exception(err_msg, e)
            already_read_text = False

    def dump_out_file(self, file_path):
        global log
        try:
            self.dump_out_file_core(file_path)
        except Exception as e:
            self.disp_messg_err("{}の出力に失敗しました。".format(OUT_FILE_NAME))
            log.exception("{}の出力に失敗しました。%s".format(OUT_FILE_NAME), e)

        self.disp_messg("{}を出力しました".format(OUT_FILE_NAME))

    def disp_messg(self, msg):
        self.ids[ID_MESSAGE].text = msg
        self.ids[ID_MESSAGE].color = (0, 0, 0, 1)

    def disp_messg_err(self, msg):
        self.ids[ID_MESSAGE].text = "{}\n詳細はログファイルを確認してください。".format(msg)
        self.ids[ID_MESSAGE].color = (1, 0, 0, 1)

    @staticmethod
    def format_size(size):
        global log
        log.info(size)


class CsvExtractorApp(App):
    def build(self):
        return MainScreen()


def setup_config():
    load_config()


def load_config():
    for line in open(CONFIG_TXT, "r", encoding=SJIS):
        items = line.replace("\n", "").split("=")

        if len(items) != 2:
            continue

        CONFIG_DICT[items[0]] = items[1]


def parse_excel_file_core(file_path):
    global TWITTER_DICT
    global excel_proc_line_num

    TWITTER_DICT = {}
    excel_proc_line_num = 1

    workbook = open_workbook(file_path)
    sheet = workbook.sheet_by_index(0)
    for i in range(1, sheet.nrows):
        row = sheet.row(i)
        mail = row[INDEX_EMAIL].value
        twitter = row[INDEX_TWITTER].value
        if not (mail in TWITTER_DICT):
            TWITTER_DICT[mail] = twitter

        excel_proc_line_num += 1


def parse_text_file_core(file_path):
    global ORDER_NUM_DICT
    global ORDER_ITEM_LIST_DICT
    global text_proc_line_num
    ORDER_NUM_DICT = {}
    ORDER_ITEM_LIST_DICT = {}
    text_proc_line_num = 1
    is_item_line = False
    is_mail_line = False
    is_order_num_line = False

    for line in open(file_path, "r", encoding=CONFIG_DICT[CONFIG_KEY_INPUT_TEXT_CHAR_SET]):
        line = line[:-1]

        if is_item_line:
            item_name = line

        elif is_mail_line:
            mail_address = line

        elif is_order_num_line:
            order_num = line
            if not (order_num in ORDER_NUM_DICT):
                order_list = ORDER_ITEM_LIST_DICT.get(mail_address, [])
                order_list.append(item_name)
                ORDER_ITEM_LIST_DICT[mail_address] = order_list
                ORDER_NUM_DICT[order_num] = True
            else:
                log.warn("注文オーダー番号 {} はすでに読込済みのため、読込をスキップします。行番号={}".format(order_num, text_proc_line_num))

        if line == KEY_ITEM_NAME:
            is_item_line = True
            is_mail_line = False
            is_order_num_line = False

        elif line == KEY_MAIL_ADDRESS:
            is_item_line = False
            is_mail_line = True
            is_order_num_line = False

        elif line == KEY_ORDER_NUMBER:
            is_item_line = False
            is_mail_line = False
            is_order_num_line = True

        else:
            is_item_line = False
            is_mail_line = False
            is_order_num_line = False

        text_proc_line_num += 1


if __name__ == '__main__':
    log = logging.getLogger('my-log')
    setup_config()
    LabelBase.register(DEFAULT_FONT, "ipaexg.ttf")
    CsvExtractorApp().run()
