import re
import shutil
import time
from json import JSONDecodeError
from typing import List
import requests
import hashlib
import winreg
import json
import uuid
import docx
import os

alphabets = "([A-Za-z])"
prefixes = "(Mr|St|Mrs|Ms|Dr)[.]"
suffixes = "(Inc|Ltd|Jr|Sr|Co)"
starters = "(Mr|Mrs|Ms|Dr|He\s|She\s|It\s|They\s|Their\s|Our\s|We\s|But\s|However\s|That\s|This\s|Wherever)"
acronyms = "([A-Z][.][A-Z][.](?:[A-Z][.])?)"
websites = "[.](com|net|org|io|gov)"


class TOOLS:
    @staticmethod
    def get_desktop() -> str:
        """
        获取用户桌面路径。
        """
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                             r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')
        return winreg.QueryValueEx(key, "Desktop")[0]

    @staticmethod
    def input_verify(prompt: str, *args):
        """
        校验用户输入
        :param prompt: 输入提示
        :return: 经用户校验了的 input() 返回值
        """
        temp = input(f'请输入{prompt}（{"、".join(args)}）：')
        while True:
            os.system("cls")
            checksum = input(f'[{prompt}]是“{temp}”，请确认（[y]/n）')
            os.system("cls")
            return temp if checksum in ['y', 'Y', ''] else TOOLS.input_verify(prompt, *args)

    @staticmethod
    def merge_segment(files_path, target_name):
        files_list = os.listdir(files_path)
        files_list.sort(key=lambda a: int(a.split('.', 1)[0]))
        with open(target_name, 'ab') as complete_file:
            for file in files_list:
                complete_file.write(open(f'{files_path}{file}', 'rb').read())

    @staticmethod
    def split_into_sentences(text):
        text = " " + text + "  "
        text = text.replace("\n", " ")
        text = re.sub(prefixes, "\\1<prd>", text)
        text = re.sub(websites, "<prd>\\1", text)
        if "Ph.D" in text: text = text.replace("Ph.D.", "Ph<prd>D<prd>")
        text = re.sub("\s" + alphabets + "[.] ", " \\1<prd> ", text)
        text = re.sub(acronyms + " " + starters, "\\1<stop> \\2", text)
        text = re.sub(alphabets + "[.]" + alphabets + "[.]" + alphabets + "[.]", "\\1<prd>\\2<prd>\\3<prd>", text)
        text = re.sub(alphabets + "[.]" + alphabets + "[.]", "\\1<prd>\\2<prd>", text)
        text = re.sub(" " + suffixes + "[.] " + starters, " \\1<stop> \\2", text)
        text = re.sub(" " + suffixes + "[.]", " \\1<prd>", text)
        text = re.sub(" " + alphabets + "[.]", " \\1<prd>", text)
        text = re.sub("(\d+)[.](\d+)", " \\1<prd>\\2 ", text)
        if "”" in text: text = text.replace(".”", "”.")
        if "\"" in text: text = text.replace(".\"", "\".")
        if "!" in text: text = text.replace("!\"", "\"!")
        if "?" in text: text = text.replace("?\"", "\"?")
        text = text.replace(".", ".<stop>")
        text = text.replace("?", "?<stop>")
        text = text.replace("!", "!<stop>")
        text = text.replace("<prd>", ".")
        sentences = text.split("<stop>")
        sentences = sentences[:-1]
        sentences = [s.strip() for s in sentences]
        return sentences


class MANUSCRIPT:
    def __init__(self, file_path='manuscript.txt'):
        self.__file_path = file_path
        self.file_name, self.file_type = os.path.splitext(self.__file_path)
        self.manuscript: List[str] = self.__get_content()
        self.preprocessed_sentence = self.__truncate()

    def __get_content(self) -> List[str]:
        if self.file_type == '.txt':
            with open(self.__file_path, 'r') as fd:
                return fd.read().split('\n')
        elif self.file_type == '.docx':
            return [prg.text for prg in docx.Document(self.__file_path).paragraphs]

    def __truncate(self) -> List[str]:
        temp = []
        for i in range(len(self.manuscript)):
            temp += TOOLS.split_into_sentences(self.manuscript[i])  # 继续拆分段落
        return temp


class TTS:
    def __init__(self):
        self.YOUDAO_URL, self.APP_KEY, self.APP_SECRET = '', '', ''
        self.SAVE_IN = ''
        self.fragment_path = './data/voice_fragment/'
        self.__load_config()

    def __load_config(self):
        try:  # 尝试读取 ./data 路径下合乎规范的 TTS_Config.json 文件。
            if not os.path.exists(self.fragment_path):
                os.makedirs(self.fragment_path)
            config_fd = open('./data/TTS_Config.json', 'r', encoding='UTF-8')
            config_info: dict = json.load(config_fd)
            self.YOUDAO_URL = config_info['youdao_url']
            certificate = config_info['certificate']
            self.APP_KEY = certificate['app_key']
            self.APP_SECRET = certificate['app_secret']
            self.SAVE_IN = config_info['save_in']
        except FileNotFoundError:  # 未在指定的路径找到，则新建。写入配置后加载
            self.__int_config()
            self.__load_config()
        except JSONDecodeError:  # 找到了对应文件但不符合规范，清空内容。写入配置后加载
            self.__int_config()
            self.__load_config()
        except KeyError:  # 找到了对应文件但不存在对应的键（疑似编辑错误），清空内容。写入配置后加载
            self.__int_config()
            self.__load_config()

    @staticmethod
    def __int_config():
        config_fd = open('./data/TTS_Config.json', 'w', encoding='UTF-8')  # 重置配置文件
        config_template = dict(youdao_url='https://openapi.youdao.com/ttsapi',
                               certificate={'app_key': '', 'app_secret': ''},
                               save_in=TOOLS.get_desktop() + '\\')
        config_template['certificate']['app_key'] = TOOLS.input_verify('应用ＩＤ', 'APP_KEY')
        config_template['certificate']['app_secret'] = TOOLS.input_verify('应用密钥', 'APP_SECRET')
        json.dump(config_template, config_fd, indent=2)

    def __encrypt_signature(self, q):
        salt = str(uuid.uuid1())
        signature = self.APP_KEY + q + salt + self.APP_SECRET
        hash_sign = hashlib.md5()
        hash_sign.update(signature.encode('UTF-8'))
        return dict(langType='en', appKey=self.APP_KEY, q=q, salt=salt, sign=hash_sign.hexdigest())

    def get_voice(self, q, q_id: int):
        headers = {'Content-Type': 'application/x-www-form-urlencoded'}
        response = requests.post(self.YOUDAO_URL, data=self.__encrypt_signature(q), headers=headers)
        if response.headers['Content-Type'] == "audio/mp3":
            file_path = self.fragment_path + str(q_id) + ".mp3"
            with open(file_path, 'wb') as voice:
                voice.write(response.content)
        else:
            print(response.content)


def main(file_path: str, subject: str, sleep_minute=5):
    manuscript = MANUSCRIPT(file_path)
    text2speech = TTS()
    sentences: list = manuscript.preprocessed_sentence
    for index in range(len(sentences)):
        print(f'正在合成第 {index + 1} 个句子：{sentences[index]}\n')
        time.sleep(sleep_minute)
        text2speech.get_voice(sentences[index], index + 1)
    os.system("cls")
    print(f'共计 {len(sentences)} 个句子的语音片段合成完毕，正在拼接……')
    complete_file = f'{text2speech.SAVE_IN}{subject}_{time.strftime("%Y%m%d")}({time.strftime("%H-%M-%S")}).mp3'
    TOOLS.merge_segment(text2speech.fragment_path, complete_file)
    shutil.rmtree(text2speech.fragment_path)


if __name__ == '__main__':
    while True:
        os.system("cls")
        p = input('请输入待转语音的 [.txt]/[.docx] 文件完整路径：')
        if os.path.exists(p):
            s = input('请输入语段的主题：')
            main(p, s, sleep_minute=6)
            break
        else:
            continue
