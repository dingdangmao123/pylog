#coding:utf-8
import os
import re
import xlwt
import sys
import time

res=[]
style = xlwt.XFStyle()
font = xlwt.Font() 
font.bold='True'
style.font = font


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
		if re.match(r'[a-z0-9]+\.log',v):
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
	for v in f:
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

	f=xlwt.Workbook()
	for k in word:
		print(len(word[k]))
		t=f.add_sheet(k)
		t.write(0,0,'No.',style)
		t.write(0,1,'Atom',style)
		t.write(0,2,'Isotropic',style)
		t.write(0,3,'Anisotropy',style)
		for v in enumerate(word[k]):
			#print k,v[0],v[1]
			t.write(v[0]+1,0,v[1][0],style)
			t.write(v[0]+1,1,k,style)
			t.write(v[0]+1,2,v[1][1],style)
			t.write(v[0]+1,3,v[1][2],style)
	s=name.split('.')[0]+'.xls'
	try:
		if os.path.exists(s):
			os.remove(s)
		f.save(name.split('.')[0]+'.xls')
		print('file '+name+' is ok')
	except Exception,e:
		fp.write("\n"+str(e)+"\n")

	



fp=open("exec.log","w+")

#print res
for v in res:
	#print v
	if len(v)>0:
		t=time.time()
		try:
			saveexcel(v)
			fp.write('file '+v+' '+str(time.time()-t)+"\n")
		except Exception:
			fp.write("\n"+str('')+"\n")
fp.close()
	

