# coding=utf-8
import xlwings as xw
import os
def func(c,n):
    return chr(ord(c)+n)
def movexy(string,x,y):
    return ((chr(ord(string[0])+x)+(str(int(string[1:])+y))).upper())
def numtoxy(a,b):
    return (func('a',a-1)+str(b)).upper()
def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        pass
 
    try:
        import unicodedata
        unicodedata.numeric(s)
        return True
    except (TypeError, ValueError):
        pass
 
    return False


xlsfile=[]
sheet=[]
numT=''
numM=''
numP=''
numS=''
knife={}
cs=['A36','A283C','A285C','A516-70','SS400','SM400B','SM400B','SB410','SGV480','A234WPB','PT410','PT480','A53B','A106','A106S','A105','A181-60','A181-70','SF440A','SF490A','SF440A','SF490A']
ss=['A240-304','A240-304L','A240-316','A240-316L','SUS304','SUS304L','SUS316','SUS316L','A403WP304','A402WP304L','A403WP316','A403WP316L','SUS304','SUS304L','SUS316','SUS316L','A312-TP304','A312-TP304L','A312TP316','A312-TP316L','SUS304TP','SUS304LTP','SUS316TP','SUS316LTP','A182-F304','A182-F304L','A182-F316','A182-F316L','A293-F316L','SUS-F304','SUS-F304L','SUS-F316','SUS-F316L']
csVc={'9.8':511,'10.2':491,'16.1':311,'16.4':305,'18.5':270,'19.2':261,'19.27':260,'19.6':255,'25':200,'25.6':195,'25.67':195,'25.9':193,'31':161,'無法辨視':0}
ssVc={'9.8':312,'10.2':300,'16.1':190,'16.4':186,'18.5':165,'19.2':159,'19.27':158,'19.6':156,'25':122,'25.6':119,'25.67':119,'25.9':118,'31':98,'無法辨視':0}
OD=[9.8,10.2,16.1,16.4,18.5,19.2,19.27,19.6,25,25.6,25.67,25.9,31]
for i in os.listdir('.'):
    #print(i)
    if(i.split('.')[1]=='xls'):
        if(i[:8]=='CNC機台週產值'):
            xlsfile.append(i)
            exname='xls'#副檔名
    elif(i.split('.')[1]=='xlsx'):
        if(i[:8]=='CNC機台週產值'):
            xlsfile.append(i)
            exname='xlsx'
for filename in xlsfile:
    counter=2
    xlbook = xw.Book(filename)
    for j in xlbook.sheets:
        sheetsname=str(j)
        if('~' in sheetsname):
            sheet.append(sheetsname.split(exname)[1][1:-1])
    if(sheetsname.split(exname)[1][1:-1]!='caculate'):
        xlbook.sheets.add('caculate',after=j)
    sheet_calc=xlbook.sheets['caculate']
    sheet_calc.range(numtoxy(1,1)).value='刀具規格'
    sheet_calc.range(numtoxy(2,1)).value='已使用(M)'
    sheet_calc.range(numtoxy(3,1)).value='工令'
    sheet_calc.range(numtoxy(4,1)).value='工件名稱'
    sheet_calc.range(numtoxy(5,1)).value='尺寸規格(T)'
    sheet_calc.range(numtoxy(6,1)).value='材質'
    sheet_calc.range(numtoxy(7,1)).value='件數'
    sheet_calc.range(numtoxy(8,1)).value='刀具外徑'
    sheet_calc.range(numtoxy(9,1)).value='孔數'
    sheet_calc.range(numtoxy(10,1)).value='進給速度'
    sheet_calc.range(numtoxy(11,1)).value='每孔秒數'
    sheet_calc.range(numtoxy(12,1)).value='所需時間(小時)(僅鑽孔)'
    sheet_calc.range(numtoxy(13,1)).value='校正時間(小時)'
    sheet_calc.range(numtoxy(14,1)).value='引孔時間'
    for j in sheet:
        sht=xlbook.sheets[j]
        flag=0
        #print(sht.range(numtoxy(1,1)).value)
        for a in range(1,10):
            for b in range(1,10):
                #print(numtoxy(a,b))
                if((sht.range(numtoxy(a,b)).value)=='尺寸規格'):
                    numT = a ##厚度X軸
                elif((sht.range(numtoxy(a,b)).value)=='材質'):
                    numM = a ##Material
                elif((sht.range(numtoxy(a,b)).value)=='件數'):
                    numQ = a ##QTY
                elif((sht.range(numtoxy(a,b)).value)=='加工內容'):
                    numP = a ##Process
                elif((sht.range(numtoxy(a,b)).value)=='工件名稱'):
                    numN = a ##Name
                elif((sht.range(numtoxy(a,b)).value)=='項次'):
                    numS = b+1 ##Start
                #print ((sht.range(numtoxy(a,b)).value))
        for b in range(numS,100): #找出底部資料
            if(sht.range(numtoxy(numP,b)).value== None):
                if(sht.range(numtoxy(numP,b+1)).value==None):
                    if(sht.range(numtoxy(numP,b+2)).value==None):
                        if(flag==0):
                            numE = b
                            flag=1
        #print(numE)
        for i in range(numS,numE):
            print(sht.range(numtoxy(numN,i)).value)
            strn = sht.range(numtoxy(numN,i)).value
            if(strn!=None):
                if('板'in strn):  #只針對板材
                    dataT = str(sht.range(numtoxy(numT,i)).value).split('T')[0]
                    if('T' not in sht.range(numtoxy(numT,i)).value):
                        data_OD = 'Φ'
                    else:
                        data_OD =str(sht.range(numtoxy(numT,i)).value).split('T')[1]
                    if('Φ' in data_OD):
                        data_OD = data_OD.split('Φ')[1]
                    else:
                        data_OD = 0
                    if('*' in dataT):
                        dataT= dataT.split('*')[-1]
                    if(is_number(dataT)):
                        #print(dataT)
                        sheet_calc.range(numtoxy(numT,counter)).value = dataT
                        dataM = str(sht.range(numtoxy(numM,i)).value)
                        dataM = dataM.replace('SA','A')
                        dataM = dataM.replace('N','')
                        dataM = dataM.split('(')[0]
                        dataM = dataM.split('+')[0]
                        if(dataM in cs):
                            sheet_calc.range(numtoxy(numM,counter)).value = 'cs'
                        elif(dataM in ss):
                            sheet_calc.range(numtoxy(numM,counter)).value = 'ss'
                        else:
                            sheet_calc.range(numtoxy(numM,counter)).value = '無法辨視'
                        sheet_calc.range(numtoxy(3,counter)).value = sht.range(numtoxy(3,i)).value
                        sheet_calc.range(numtoxy(4,counter)).value = sht.range(numtoxy(4,i)).value
                        sheet_calc.range(numtoxy(7,counter)).value = sht.range(numtoxy(7,i)).value
                        dataP = sht.range(numtoxy(8,i)).value
                        if(('Φ') not in dataP):
                            sheet_calc.range(numtoxy(8,counter)).value = '無法辨視'
                        else:
                            if(('+')in dataP.split('Φ')[1].split('*')[0]):
                                sheet_calc.range(numtoxy(8,counter)).value = dataP.split('Φ')[1].split('*')[0].split('+')[0]
                            elif(('，')in dataP.split('Φ')[1].split('*')[0]):
                                sheet_calc.range(numtoxy(8,counter)).value = dataP.split('Φ')[1].split('*')[0].split('，')[0]
                            else:
                                sheet_calc.range(numtoxy(8,counter)).value = dataP.split('Φ')[1].split('*')[0]
                            if('*' not in str(dataP.split('Φ')[1])):
                                dataP2=''
                            else:
                                dataP2 = str(dataP.split('Φ')[1].split('*')[1])
                        #print(dataP2)
                        if(('孔') not in dataP2):
                            sheet_calc.range(numtoxy(9,counter)).value = '無法辨視'
                        else:
                            sheet_calc.range(numtoxy(9,counter)).value = dataP2.split('孔')[0]
                        #print(dataP2)
                        if(sheet_calc.range(numtoxy(8,counter)).value!='無法辨視'):
                            dictvalue = str(min(OD, key=lambda x:abs(x-float(sheet_calc.range(numtoxy(8,counter)).value))))
                        else:
                            dictvalue = '無法辨視'
                        if(sheet_calc.range(numtoxy(6,counter)).value=='cs'):
                            sheet_calc.range(numtoxy(10,counter)).value=csVc[dictvalue]
                        elif(sheet_calc.range(numtoxy(6,counter)).value=='ss'):
                            sheet_calc.range(numtoxy(10,counter)).value=ssVc[dictvalue]
                        else:
                            sheet_calc.range(numtoxy(10,counter)).value=''
                        #if(sheet_calc.range(numtoxy(6,counter)).value)=='cs'
                        if((sheet_calc.range(numtoxy(5,counter)).value=='無法辨視')or (sheet_calc.range(numtoxy(6,counter)).value=='無法辨視') or(sheet_calc.range(numtoxy(7,counter)).value==None)  or(sheet_calc.range(numtoxy(8,counter)).value=='無法辨視') or(sheet_calc.range(numtoxy(9,counter)).value=='無法辨視') or(sheet_calc.range(numtoxy(9,counter)).value==None) or(sheet_calc.range(numtoxy(10,counter)).value==None) or(sheet_calc.range(numtoxy(10,counter)).value==0) ):
                            sheet_calc.range(numtoxy(11,counter)).value=None
                        else:
                            sheet_calc.range(numtoxy(11,counter)).value = float(sheet_calc.range(numtoxy(5,counter)).value) * 1.15/float(sheet_calc.range(numtoxy(10,counter)).value)*60
                            sheet_calc.range(numtoxy(12,counter)).value = float(sheet_calc.range(numtoxy(7,counter)).value) * float(sheet_calc.range(numtoxy(11,counter)).value) * float(sheet_calc.range(numtoxy(9,counter)).value) / 60/60
                            sheet_calc.range(numtoxy(13,counter)).value = float(data_OD)/1000+0.5
                            knifedata = str(sheet_calc.range(numtoxy(8,counter)).value)
                            kinfevalue = round(float(sheet_calc.range(numtoxy(5,counter)).value) * float(sheet_calc.range(numtoxy(7,counter)).value) * float(sheet_calc.range(numtoxy(9,counter)).value),2)
                            #print(kinfevalue)
                            if(knifedata in knife):
                                knife[knifedata] = knife[knifedata] + kinfevalue
                            else:
                                knife[knifedata] = kinfevalue
                        #print(str(sheet_calc.range(numtoxy(8,counter)).value))
                        if(sheet_calc.range(numtoxy(9,counter)).value =='無法辨視'):
                            sheet_calc.range(numtoxy(14,counter)).value = ''
                        else:
                            coefficient = sheet_calc.range(numtoxy(9,counter)).value /1000*0.2+1.2
                            if(sheet_calc.range(numtoxy(6,counter)).value=='cs'):
                                v_giant = 80
                            if(sheet_calc.range(numtoxy(6,counter)).value=='ss'):
                                v_giant = 50
                            print(coefficient)
                            sheet_calc.range(numtoxy(14,counter)).value = (sheet_calc.range(numtoxy(9,counter)).value*8/v_giant*coefficient+0.1*sheet_calc.range(numtoxy(9,counter)).value)/60
                    sheet_calc.autofit()
                    counter = counter+1
    #print(knife)
    counter = 2
    for va in knife:
        sheet_calc.range(numtoxy(1,counter)).value = va
        sheet_calc.range(numtoxy(2,counter)).value = knife[va]/1000
        counter = counter+1
    sheet_calc.autofit()
           # a = numtoxy(1,i)
            #xy(a,0,1)
            #if(sht.range(a).value=='1'):
                #print(sht.range(numtoxy(numT,i)).value)
                #print(xy(a,1,0))
                #print(sht.range(xy(a,1,0)).value)
