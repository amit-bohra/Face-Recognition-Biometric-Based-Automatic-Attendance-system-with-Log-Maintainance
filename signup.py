import cv2
import random as rd
import face_recognition as fc
import smtplib
import xlsxwriter as xl
import pandas as pd
import numpy as np
'''workbook=xl.Workbook('registration.xlsx')
worksheet=workbook.add_worksheet()
worksheet.write('A1','Encodings')
worksheet.write('B1','Name')
worksheet.write('C1','Mobile Number')
worksheet.write('D1','Email ID')
worksheet.write('E1','Category')
worksheet.write('F1','Gender')
worksheet.write('G1','OTP SENT')'''
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login("ENTER YOUR EMAIL HERE", "ENTER YOUR PASSWORD") #Senders email and password
v=cv2.VideoCapture(0)
fd=cv2.CascadeClassifier(r'haarcascade_frontalface_alt2.xml')
face_img=0
face_list=[]
name=0
mobile=0
email=0
category=0
gender=0
otps=0
otpr=0
otp_count=0
face_encod=0
serial=0
content=[]
while(True):
    ret,img=v.read()
    grey_img=cv2.cvtColor(img,cv2.COLOR_BGR2GRAY)
    face=fd.detectMultiScale(grey_img)
    for x,y,w,h in face:
        face_img=img[y:y+h,x:x+h].copy()
    cv2.imshow('face',face_img)
    key=cv2.waitKey(5)
    if key==ord('d'):
        column=0
        name=input('Enter your name.\n').upper()
        mobile=int(input('Enter your mobile number.\n'))
        email=input('Enter your email id.\n')
        category=input('Enter "S" for Student and "T" for Teacher.\n')
        gender=input('Enter gender "M" for Male and "F" for Female.\n')
        otps=rd.randint(1000,10001)
        print('\nOK '+name+' \nCheck your email, an OTP is sent to you.\n')
        msg='Greetings '+name+' \nYour OTP for AI based Biometric Registration is:\n'+str(otps)+' .'
        server.sendmail('tomjerrybhaibhai@gmail.com',email,msg)
        flag=0
        while(otps!=otpr):
            if flag==0 and otpr!=otps:
                otpr=int(input('Enter the OTP\n'))
                flag=1
            elif flag==1 and otpr!=otps and otp_count!=2:
                otpr=int(input('\nIncorrect OTP......!!!!\nEnter Again\n\n'))
                otp_count+=1
            elif otp_count==2:
                print('Quitting Program')
                quit()
                break
        serial+=1
        print('Your otp Verified\nTHANK YOU\n')
        face_loc=fc.face_locations(face_img)
        face_encod=fc.face_encodings(face_img,face_loc)
        content.append(face_encod)
        content.append(name)
        content.append(mobile)
        content.append(email)
        content.append(category)
        content.append(gender)
        content.append(otps)
        Data=pd.read_excel('registration.xlsx')
        df=pd.DataFrame(Data)
        dataf=pd.DataFrame({'Serial':[serial],
                            'Encodings':list(face_encod),
                            'Name':[name],
                            'Mobile Number':[mobile],
                            'Email ID':[email],
                            'Category':[category],
                            'Gender':[gender],
                            'OTP SENT':[otps]})
        df=df.append(dataf,ignore_index=True,sort=False)
        df.to_excel('registration.xlsx',index=False)
        reData=pd.read_excel('attendance.xlsx')
        df=pd.DataFrame(reData)
        redataf=pd.DataFrame({'Serial':[serial],
                              'Name':[name],
                              'Email Id':[email]})
        df=df.append(redataf,ignore_index=True,sort=False)
        df.to_excel('attendance.xlsx',index=False)
        print('Data Entered')
            
                
        
        
        
