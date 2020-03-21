import logging.config
import os
import sys
from collections import OrderedDict
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
INDEX_TWITTER = 5
INDEX_EMAIL = 1

ID_MESSAGE = "message"

OUT_FILE_NAME = "result.csv"
UTF8 = "utf8"
SJIS = "sjis"

EXT_TXT = ".txt"
EXT_XLS = ".xlsx"
EXT_XLSX = ".xlsx"
EXT_XLSM = ".xlsm"

KEY_MAIL_ADDRESS = "メールアドレス"

CONFIG_TXT = "./config.txt"
CONFIG_DICT = {}
CONFIG_KEY_OUTPUT_CSV_NAME = "OUTPUT_CSV_NAME"
CONFIG_KEY_OUTPUT_CSV_CHAR_SET = "OUTPUT_CSV_CHAR_SET"
CONFIG_KEY_INPUT_TEXT_CHAR_SET = "INPUT_TEXT_CHAR_SET"
MAIL_ADDRESS_DICT_FROM_TXT = OrderedDict()
DISCORD_ID_DICT_FROM_EXCEL = {}
MAIL_ADDRESS_DICT_FROM_EXCEL = {}

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
        self.dump_twitter_and_item_list(out_file_name)
        self.disp_messg("{}を出力しました".format(out_file_name))

    @staticmethod
    def dump_twitter_and_item_list(out_file_name):
        with open(out_file_name, "w", encoding=CONFIG_DICT[CONFIG_KEY_OUTPUT_CSV_CHAR_SET]) as f:
            already_dumped_discord_id_dict = {}
            for mail in MAIL_ADDRESS_DICT_FROM_TXT.keys():
                discord_id = MAIL_ADDRESS_DICT_FROM_EXCEL.get(mail)
                if discord_id is None:
                    log.warn("アドレス {} はExcelファイルに存在しません。処理をスキップします。".format(mail))
                    continue
                if discord_id in already_dumped_discord_id_dict:
                    continue

                f.write("{}".format(discord_id))
                for mail in DISCORD_ID_DICT_FROM_EXCEL[discord_id]:
                    f.write(",{}\n".format(mail))
                
                already_dumped_discord_id_dict[discord_id] = True


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
    global MAIL_ADDRESS_DICT_FROM_EXCEL
    global excel_proc_line_num

    MAIL_ADDRESS_DICT_FROM_EXCEL = {}
    excel_proc_line_num = 1

    workbook = open_workbook(file_path)
    sheet = workbook.sheet_by_index(0)
    for i in range(1, sheet.nrows):
        row = sheet.row(i)
        mail = row[INDEX_EMAIL].value
        twitter = row[INDEX_TWITTER].value
        if not (mail in MAIL_ADDRESS_DICT_FROM_EXCEL):
            mail_list = DISCORD_ID_DICT_FROM_EXCEL.get(twitter, [])
            mail_list.append(mail)
            DISCORD_ID_DICT_FROM_EXCEL[twitter] = mail_list
            MAIL_ADDRESS_DICT_FROM_EXCEL[mail] = twitter
        else:
            log.warn("メールアドレス {} はすでに読込済みのため、読込をスキップします。ファイル={} 行番号={}".format(mail, file_path, excel_proc_line_num))

        excel_proc_line_num += 1


def parse_text_file_core(file_path):
    global MAIL_ADDRESS_DICT_FROM_TXT
    global text_proc_line_num
    MAIL_ADDRESS_DICT_FROM_TXT = OrderedDict()
    text_proc_line_num = 1
    is_item_line = False
    is_mail_line = False
    is_order_num_line = False

    for line in open(file_path, "r", encoding=CONFIG_DICT[CONFIG_KEY_INPUT_TEXT_CHAR_SET]):
        line = line[:-1]
        if is_mail_line:
            mail_address = line
            if not (mail_address in MAIL_ADDRESS_DICT_FROM_TXT):
                MAIL_ADDRESS_DICT_FROM_TXT[mail_address] = True
            else:
                log.warn("メールアドレス {} はすでに読込済みのため、読込をスキップします。ファイル={} 行番号={}".format(mail_address, file_path, text_proc_line_num))

        if line == KEY_MAIL_ADDRESS:
            is_mail_line = True
        else:
            is_mail_line = False

        text_proc_line_num += 1


if __name__ == '__main__':
    log = logging.getLogger('my-log')
    setup_config()
    LabelBase.register(DEFAULT_FONT, "ipaexg.ttf")
    CsvExtractorApp().run()
