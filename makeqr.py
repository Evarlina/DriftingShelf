from json import load
from os import getcwd
from time import sleep

import win32.lib.win32con as win32con
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from win32 import win32gui


class MakeQR:

    def __init__(self):
        """ Create explorer.
        """
        # Ignore certain log.
        option = webdriver.ChromeOptions()
        option.add_experimental_option(
            'excludeSwitches', ['enable-automation', 'enable-logging'])

        # Instantialize chrome explorer.
        try:
            self.chrome = webdriver.Chrome(
                executable_path='./rsc/chromedriver.exe', chrome_options=option)
        except Exception as e:
            print(e)

        # Maximize chrome window.
        self.chrome.maximize_window()

    def log_in(self, username, password):
        """ Log in.
        """
        # Access log-in page.
        try:
            self.chrome.get(url='https://user.cli.im/login')
        except Exception as e:
            print(e)

        # Fill username frame.
        usrnm = self.chrome.find_element_by_id('loginemail')
        usrnm.send_keys(username)

        # Fill password frame.
        psw = self.chrome.find_element_by_id('loginpassword')
        psw.send_keys(password)

        # Click log-in button.
        login_btn = self.chrome.find_element_by_id('login-btn')
        login_btn.click()

    def go_to_template(self):
        """ Go to the batch production template.
        """
        # Wait till button <批量模板> is ready. Time-out is set to 20 secs.
        target_a = WebDriverWait(self.chrome, 20, 0.5).until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="1000200$Menu"]/li[2]/a'))
        )
        target_url = target_a.get_attribute('href')

        # Access <批量模板> page.
        try:
            self.chrome.get(target_url)
        except Exception as e:
            print(e)

    def process_dialog(self, filename):
        """ Deal with dialog frame using module "win32gui" and "win32con".
        """
        # Locate dialog frame.
        dialog = win32gui.FindWindow('#32770', '打开')

        # Locate filepath input frame.
        ComboBoxEx32 = win32gui.FindWindowEx(dialog, 0, 'ComboBoxEx32', None)
        ComboBox = win32gui.FindWindowEx(ComboBoxEx32, 0, 'ComboBox', None)
        Edit = win32gui.FindWindowEx(ComboBox, 0, 'Edit', None)

        # Locate open button.
        open_btn = win32gui.FindWindowEx(dialog, 0, 'Button', None)

        # Send message to two frames.
        win32gui.SendMessage(Edit, win32con.WM_SETTEXT,
                             0, f'{getcwd()}\\{filename}')
        win32gui.SendMessage(dialog, win32con.WM_COMMAND, 1, open_btn)

    def upload_bookinfo(self):
        """ Upload file "书籍信息.xlsx".
        """
        # Wait till button <我已填好Excel，下一步> is ready. Time-out is set to 30 secs.
        next_step_btn = WebDriverWait(self.chrome, 30, 0.5).until(
            EC.presence_of_element_located(
                (By.XPATH, '/html/body/div[11]/div/div[2]/div/div[2]/div/div/div[2]/div[2]/button[2]'))
        )

        # Click button <我已填好Excel，下一步>.
        next_step_btn.click()

        # Wait till button <上传Excel> is ready. Time-out is set to 30 secs.
        upload_btn = WebDriverWait(self.chrome, 30, 0.5).until(
            EC.presence_of_element_located(
                (By.XPATH, '/html/body/div[11]/div/div[2]/div/div[2]/div/div/div[2]/div[1]/span/div/span/div/button'))
        )

        # Click button <上传Excel>.
        upload_btn.click()

        # Wait for the dialog frame to open.
        sleep(4)

        # Open file "书籍信息.xlsx".
        self.process_dialog(f'rsc\\书籍信息.xlsx')

    def upload_proof(self):
        """ Upload file "捐赠证明.xlsx".
        """
        # Wait till button <我已填好Excel，下一步> is ready. Time-out is set to 30 secs.
        next_step_btn = WebDriverWait(self.chrome, 30, 0.5).until(
            EC.presence_of_element_located(
                (By.XPATH, '/html/body/div[11]/div/div[2]/div/div[2]/div/div/div[2]/div[2]/button[2]'))
        )

        # Click button <我已填好Excel，下一步>.
        next_step_btn.click()

        # Wait till button <上传Excel> is ready. Time-out is set to 30 secs.
        upload_btn = WebDriverWait(self.chrome, 30, 0.5).until(
            EC.presence_of_element_located(
                (By.XPATH, '/html/body/div[11]/div/div[2]/div/div[2]/div/div/div[2]/div[1]/span/div/span/div/button'))
        )

        # Click button <上传Excel>.
        upload_btn.click()

        # Wait for the dialog frame to open.
        sleep(4)

        # Open file "书籍信息.xlsx".
        self.process_dialog(f'rsc\\捐赠证明.xlsx')

    def create_qrcode(self):
        """ Create QR Code.
        """
        # Wait till button <开始生码> is ready. Time-out is set to 60 secs.
        start_mkqr_btn = WebDriverWait(self.chrome, 60, 0.5).until(
            EC.presence_of_element_located(
                (By.XPATH, '/html/body/div[11]/div/div[2]/div/div[2]/div/div/div[4]/button'))
        )

        # Click button <开始生码>.
        start_mkqr_btn.click()

    def download_qrcode(self):
        """ Download QR Code.
        """
        # Wait till button <下载二维码> is ready. Time-out is set to 60 secs.
        download_btn = WebDriverWait(self.chrome, 60, 0.5).until(
            EC.presence_of_element_located(
                (By.XPATH, '/html/body/div[11]/div/div[2]/div/div[2]/div/div/div[4]/button[2]'))
        )

        # Click button <下载二维码>.
        download_btn.click()

        # Wait till button <下载> is ready. Time-out is set to 20 secs.
        final_btn = WebDriverWait(self.chrome, 20, 0.5).until(
            EC.presence_of_element_located(
                (By.XPATH, '/html/body/div[11]/div/div[2]/div/div[2]/div[3]/div/div[2]/button[2]'))
        )

        # Click button <下载>.
        final_btn.click()


if __name__ == '__main__':
    i_username = input('账号>> ').strip()
    i_password = input('密码>> ').strip()
    mkqr = MakeQR()
    mkqr.log_in(i_username, i_password)
    mkqr.go_to_template()

    mkqr.upload_bookinfo()
    mkqr.create_qrcode()
    mkqr.download_qrcode()

    mkqr.upload_proof()
    mkqr.create_qrcode()
    mkqr.download_qrcode()
    sleep(10)
