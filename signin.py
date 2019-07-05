import cv2
import random as rd
import face_recognition as fc
import smtplib
import xlsxwriter as xl
import pandas as pd
import numpy as np
import xlrd
import time
import datetime as dt
v=cv2.VideoCapture(0)
fd=cv2.CascadeClassifier(r'haarcascade_frontalface_alt2.xml')
face_img=0
column=0
length=0
naam=0
ind=0
reData=pd.read_excel('attendance.xlsx')
df1=pd.DataFrame(reData)
samay=dt.datetime.now()
#dates=samay.strftime('%H:%M')
l=df1.shape[1]
df1.insert(l,'81:00','')
while(True):
    ret,img=v.read()
    grey_img=cv2.cvtColor(img,cv2.COLOR_BGR2GRAY)
    face=fd.detectMultiScale(grey_img)
    if len(face)==0:
        continue
    for x,y,w,h in face:
        face_img=img[y:y+h,x:x+h].copy()
    cv2.imshow('face',face_img)
    key=cv2.waitKey(5)
    if key==ord('r'):
        face_loc=fc.face_locations(face_img)
        face_encod=fc.face_encodings(face_img,face_loc)
        if len(face_encod)==0:
            continue
        Data=pd.read_excel('registration.xlsx')
        df=pd.DataFrame(Data)
        length=len(df['Encodings'])
        for i in range(length):
            column=df['Encodings'][i]
            jin=len(column)
            column=column[1:-2]
            print(column)
            a=column.split(' ')
            newlist=[]
            r=[x for x in a if x!='']
            for x in r:
                if '\n' in x:
                    f=float(x[:-2])
                    newlist.append(f)
                elif '[' in x:
                    f=float(x[1:])
                    newlist.append(f)
                elif ']' in x:
                    pass
                    #f=float(x[:-2])
                    #newlist.append(f)
                else:
                    newlist.append(float(x))
            print(newlist)
            new_arr=np.array(newlist,dtype='float64')
            f=fc.compare_faces([new_arr],face_encod[0])
            if f==[False]:
                continue
            else:
                naam=(df['Name'][i])
                ind=i
                df1['8:00'][ind]='P'
                df1.to_excel('attendance.xlsx',index=False)
                print(naam+' lag gayi')
            
            
        
        
    
