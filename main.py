from proceed import SheetOperation
from makeqr import MakeQR
from time import sleep

# Proceed sheets.
shtopr = SheetOperation()
shtopr.load_json()
shtopr.start_loop()
shtopr.close()

# Make QR Code.
mkqr = MakeQR()
mkqr.log_in('18962388966', 'siaoca708401')
mkqr.go_to_template()
mkqr.upload_bookinfo()
mkqr.create_qrcode()
mkqr.download_qrcode()
mkqr.upload_proof()
mkqr.create_qrcode()
mkqr.download_qrcode()
sleep(10)
