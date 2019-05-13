#BruteForce script for protected exel file
import sys
import win32com.client
openedDoc = win32com.client.Dispatch("Excel.Application")
filename= sys.argv[1]

password_file = open ( 'C:\Users\someone\Desktop\crack\wordlist.txt', 'r' )  # BURAYA WORDLIST YAZILACAK.
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