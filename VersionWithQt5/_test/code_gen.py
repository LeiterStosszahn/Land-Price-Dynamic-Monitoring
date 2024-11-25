#month-day[0:4]-year-day[3:8]-code
import math,hashlib

def only_contains(string, chars):
    return set(string) <= set(chars)

def encode_key(y,m,d):
	key=[]

	key_moth=1<<m
	key_month=hex(key_moth)
	key.append(key_month)

	key_day=d<<3
	key_day=bin(key_day)[2:]
	key.append(key_day[0:4])

	key_year=math.log(y)*(10**6)
	key_year=hex(math.ceil(key_year))[2:]
	key.append(key_year)

	key.append(key_day[4:8])

	veryfication_code=(y+m+d)<<1
	md5hash=hashlib.md5(str(veryfication_code).encode())
	veryfication_code=md5hash.hexdigest()
	key.append(veryfication_code[8:14])

	keystr=key[0]+"-"+key[1]+"-"+key[2]+"-"+key[3]+"-"+key[4]
	return keystr

def decode_key(key):
	key_list=key.split("-")
	if len(key_list)!=5:
		return 0,0,0,0

	#key_year=log(y)*10^6->HEX
	year=key_list[2]
	year=int(year,16)/1000000
	if year<700:
		year=int(math.exp(year))
	if year>2100 or year<2024:
		year=1999
	#key_moth=2^m->HEX
	month=key_list[0]
	month=int(math.log(int(month,16),2))
	if month>12:
		month=1
	#key_day=d*8->BIN
	day=key_list[1]+key_list[3]
	if only_contains(day,"01"):
		day=int(int(day,2)>>3)
	else:
		day=1
	if day>30 or (month==2 and day>28):
		day=1
	#veryfication code=(y+m+d)*2->hash
	code=key_list[4]
	veryfication_code=(year+month+day)<<1
	md5hash=hashlib.md5(str(veryfication_code).encode())
	veryfication_code=md5hash.hexdigest()
	if code==veryfication_code[8:14]:
		hash_result=0
	else:
		hash_result=1
	return year,month,day,hash_result

if __name__ == '__main__':
	# y=input("年：")
	# m=input("月：")
	# d=input("日：")
	y=2024
	m=6
	d=30
	key=encode_key(y,m,d)
	print(key)