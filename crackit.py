#BruteForce script for protected exel file
import sys
import pandas as pd
openedDoc = pd.ExcelFile('/content/111.xlsx')
filename= sys.argv[1]

password_file = open ( '/content/cmnd.txt', 'r' )  # BURAYA WORDLIST YAZILACAK.
passwords = password_file.readlines()
password_file.close()

passwords = [item.rstrip('\n') for item in passwords]

results = open('results.txt', 'w') #BURASI DEGISTIRILMEYECEK, SIFRE TESPIT EDILDIGINDE BU DOSYAYA YAZACAK.

for password in passwords:
	print(password)
	try:
		wb = openedDoc.Workbooks.Open(filename, False, True, None, password)
		print("Success! Password is: "+password)
		results.write(password)
		results.close()
		break
	except:
		print("Incorrect password")
pass
