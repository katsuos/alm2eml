#!python3

#ALM(AL-Mail)ファイルをEMLファイルに変換する
#2024/10/13 katsuos

import os
import sys
import configparser
import glob
import re
import email
from email.header import decode_header
from email.policy import default

#フォルダ作成
def MyMakeDirs(FolderName):
	print('makedirs {}'.format(FolderName))
	if not os.path.isdir(FolderName):
		try:
			os.makedirs(FolderName)
			pass
		except FileExistsError:
			pass
		except OSError as e:
			print('OSError:{}'.format(e.strerror))
			sys.exit(-1)


#CP932で読み込んでUTF-8で保存する
#Content-Typeヘッダを書き換える
def Convert(filename,DstFolder):
	try:
		with open(filename,encoding='cp932',errors='backslashreplace') as fin:
			msg = email.message_from_file(fin, policy=default)
			for part in msg.walk():
				#print('1:',part.get_content_type(),':',part.get_content_charset(''))
				if part.get_content_type()=='text/plain':
					if part.get_content_charset('')=='iso-2022-jp':
						part.set_charset('UTF-8')
				#print('2:',part.get_content_type(),':',part.get_content_charset(''))
	except FileNotFoundError:
		#ファイルが無ければ次
		return

	f=os.path.splitext(os.path.basename(filename))[0]+'.eml'
	f=os.path.join(DstFolder,f)	
	print("\r{}".format(f),end="")
	with open(f,mode='w',encoding='utf-8') as fout:
		fout.write(msg.as_string(policy=default))


#Userフォルダの処理
def ProcUserFolder(SrcFolder,DstFolder):
	print('UserFolder({},{})'.format(SrcFolder,DstFolder))

	#Entry.ini 読み込み
	f=os.path.join(SrcFolder,'Entry.ini')
	if not os.path.isfile(f):
		return
	#print('Entry.ini {}'.format(f))

	ENTRYINI = configparser.ConfigParser()
	ENTRYINI.read(f,encoding='cp932')
	try:
		title=ENTRYINI['Property']['Title']	
	except KeyError:
		if(os.path.basename(SrcFolder)=='Inbox.box' ):
			title='郵便受け'
		elif( os.path.basename(SrcFolder)=='Outbox.box' ):
			title='送信箱'
		else:
			print('Entry.ini 読み取りエラー')
			sys.exit(-1)
	print('Title {} {}mails'.format(title,ENTRYINI['Property']['Total']))

	#フォルダ作成
	d=os.path.join(DstFolder,title)
	MyMakeDirs(d)

	#Entry.lst処理
	fn=os.path.join(SrcFolder,'Entry.lst')
	with open(fn,encoding='cp932',errors='backslashreplace') as fp:
		s = fp.readlines()
	for t in s:
		fn=os.path.join(SrcFolder,t[:8] + '.ALM')
		Convert(fn,d)

	print('\n')


#Trash.boxの処理
def ProcTrashFolder(SrcFolder,DstFolder):
	print('TrashFolder({},{})'.format(SrcFolder,DstFolder))

	#フォルダ作成
	d=os.path.join(DstFolder,'ごみ箱')
	MyMakeDirs(d)

	fn=os.path.join(SrcFolder,'Trash.box')
	with open(fn,encoding='cp932',errors='backslashreplace') as fp:
		s = fp.readlines()

	#1個上のフォルダ
	SrcFolder=os.path.dirname(SrcFolder)

	for t in s:
		x=re.match(r'.+?=',t)
		if x is not None:
			f=os.path.join(SrcFolder,x.group()[:-1]+'.ALM')
			Convert(f,d)

	print('\n')



#Accountフォルダの処理
def ProcAccountFolder(SrcFolder,DstFolder):
	print('AccountFolder({},{})'.format(SrcFolder,DstFolder))

	#Account.ini 読み取り
	f=os.path.join(SrcFolder,'Account.INI')
	#print('Account.ini {}'.format(f))
	ACCOUNTINI = configparser.ConfigParser()
	ACCOUNTINI.read(f,encoding='cp932')
	try:
		title=ACCOUNTINI['Property']['Title']
	except KeyError:
		print('読み取りエラー{}'.format(f))
		sys.exit(-1)
	#ファイル名に使えない文字を取り除く
	title = re.sub(r'[\\/:*?"<>|]+','',title)

	#フォルダ作成
	d=os.path.join(DstFolder,title)
	MyMakeDirs(d)

	#ごみ箱の処理
	ProcTrashFolder(SrcFolder,DstFolder)

	#郵便受けと送信済みフォルダの処理
	f=os.path.join(SrcFolder,'Inbox.box')
	ProcUserFolder(f,d)
	f=os.path.join(SrcFolder,'Outbox.box')
	ProcUserFolder(f,d)
		
	#User0000.boxフォルダの処理
	userfolders=glob.glob(r'User[0-9][0-9][0-9].box',root_dir=SrcFolder)
	for usr in userfolders:
		x=os.path.join(SrcFolder,usr)
		ProcUserFolder(x,d)

	
	print('')


##############################################################
# MAIN
if __name__ == "__main__":
	
	print('AL-Mail to EML Converter v1.0 2024/10/13 katsuos\n')

	if( len(sys.argv)!=3 ):
		print('使い方 {} [AL-Mailのメールボックスフォルダ] [EMLファイル出力フォルダ]'.format(sys.argv[0]))
		sys.exit(-1)

	SrcFolder = sys.argv[1]
	DstFolder = sys.argv[2]
	#print('MailboxFolder {}'.format(SrcFolder))
	#print('OutputFolder  {}'.format(DstFolder))

	#OutputFolder 作成
	MyMakeDirs(DstFolder)

	#MAINBOX.INI 存在チェック
	f=os.path.join(SrcFolder,'MAILBOX.INI')
	#print('MAILBOX.INI:{}'.format(f))
	MAILBOXINI = configparser.ConfigParser()
	MAILBOXINI.read(f, encoding='cp932')
	try:
		x=MAILBOXINI['Property']['AddressFile']
		#print('AddressFile {}'.format(x))
	except KeyError:
		print('{} が見つからない'.format(f))
		sys.exit(-1)
		
	#Accountフォルダの処理
	accountfolders=glob.glob(SrcFolder+'/Account[0-9]')
	for acc in accountfolders:
		ProcAccountFolder(acc,DstFolder)



