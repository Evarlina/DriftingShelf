from proceed import SheetOperation
from makeqr import MakeQR
from time import sleep

# Proceed sheets.
shtopr = SheetOperation()
shtopr.load_json()
shtopr.start_loop()
shtopr.close()

# Make QR Code.
print('-' * 60)
print('请输入草料二维码平台的账号信息。')
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
