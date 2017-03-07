#coding:utf-8
import os
import re
import xlsxwriter as xlsx
import sys
import time

res=[]



if len(sys.argv)>1 and sys.argv[1]=='-w':
	if os.path.exists(sys.argv[2]):
		os.chdir(sys.argv[2])
		print('work dir->'+sys.argv[2])
		work=sys.argv[3:]
	else:
		print(sys.argv[2]+' not exists')
		sys.exit()
else:
	work=sys.argv[1:]
#print work

if len(work)>0:
	for v in  work:
		if re.match(r'[A-Za-z0-9]+\.log',v) and v!='exec.log':
			res.append(v)
		else:
			print('file '+v+'  is invalid')

if len(res)==0 and os.path.exists('config.log'):
	
	f=open("config.log")
	for v in f:
		if re.match(r'[a-z0-9]+\.log',v):
			res.append(v.strip('\n').strip('\r'))
		else:
			print('config->file '+v+' is invalid')

if len(res)==0:
	for v in  os.listdir(os.getcwd()):
		#print(v)
		if re.match(r'[a-z0-9]+\.log',v):
			res.append(v.strip('\n').strip('\r'))
		

if len(res)==0:
	print('no valid file')
	sys.exit()

def getinfo(v):
	words={}
	
	p=re.compile(r'\s+')
	f=open(v,'r')
	i=0
	for v in f.readlines():
		if 'Isotropic =' in v:
			#i=i+1
			tmp=p.split(v.strip())
			#print tmp
			if tmp[1] in words:
				words[tmp[1]].append([tmp[0],tmp[4],tmp[7]])
			else:
				words[tmp[1]]=[[tmp[0],tmp[4],tmp[7]]]
	f.close()
	#print i
	return words


def saveexcel(name):
	word=getinfo(name)
	#print name
	'''
	print word
	print len(word['C'])
	print len(word['H'])
	print len(word['O'])
	print len(word['N'])
	'''
	#sys.exit()
	s=name.split('.')[0]+'.xlsx'
	if os.path.exists(s):
		os.remove(s)

	f=xlsx.Workbook(s)




	for k in word:
		print(len(word[k]))
		
		t=f.add_worksheet(k)
		t.write(0,0,'No.')
		t.write(0,1,'Atom')
		t.write(0,2,'Isotropic')
		t.write(0,3,'Anisotropy')
		i=0
		for v in enumerate(word[k]):#print k,v[0],v[1]


			t.write(v[0]+1,0,v[1][0])
			t.write(v[0]+1,1,k)
			t.write(v[0]+1,2,v[1][1])
			t.write(v[0]+1,3,v[1][2])

	f.close()

	


if os.path.exists("exec.log"):
	os.remove("exec.log")

fp=open("exec.log","w+")

#print res
for v in res:
	#print v
	if len(v)>0:
		t=time.time()
		try:
			saveexcel(v)
			fp.write('file '+v+' '+str(time.time()-t)+"\n")
		except Exception as ex:
			fp.write("\n"+str(ex)+"\n")

fp.close()
	

