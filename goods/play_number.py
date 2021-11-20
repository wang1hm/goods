#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Jun  2 22:33:54 2021

@author: haoxinchiguren
"""
import pandas
#import tkinter
import sys
import numpy as np
import xlsxwriter
import time
address = sys.path[0]


def auto_table(name,mode):
    data=pandas.read_excel(address+'/'+name+'.xlsx',sheet_name=None)#更换成你自己文件的地址
    res=pandas.DataFrame(columns=['昵称','角色','总点数','肾款','备用'])
    if mode == '1':
        a = 0
        for m in range(len(data)):
            sheet_n = input("请告诉我第"+str(m+1)+"张sheet的名字是（默认请输入‘Sheet1’）：")
            num = (len(data[sheet_n].iloc[0])-3)/2
            for i in range(len(data[sheet_n])):
                for j in range(int(num)):
                    lis1 = list(res['昵称'])
                    lis2 = list(res['角色'])
                    if data[sheet_n].iloc[i][2*j+1] not in lis1 and str(data[sheet_n].iloc[i][2*j+1]) != 'nan':
                        res=res.append({'角色':1},ignore_index=True)
                        res['昵称'][a]=data[sheet_n].iloc[i][2*j+1]
                        res['角色'][a]=[data[sheet_n].iloc[:,0][i]]
                        res['总点数'][a]=int(data[sheet_n].iloc[i][2*j+2])
                        res['肾款'][a]=(data[sheet_n]['均价'][i]+data[sheet_n]['调价'][i])*data[sheet_n].iloc[i][2*j+2]
                        res['备用'][a]=data[sheet_n].iloc[:,0][i]+str(int(data[sheet_n].iloc[i][2*j+2]))+' '
                        a += 1
                    if data[sheet_n].iloc[i][2*j+1] in lis1 and str(data[sheet_n].iloc[i][2*j+1]) != 'nan':
                        b = lis1.index(data[sheet_n].iloc[i][2*j+1])
                        res['肾款'][b] += ((data[sheet_n]['均价'][i]+data[sheet_n]['调价'][i]))*data[sheet_n].iloc[i][2*j+2]
                        res['角色'][b].append(data[sheet_n]['角色'][i])
                        res['总点数'][b]+=int(data[sheet_n].iloc[i][2*j+2])
                        print(res)
                        res['备用'][b]+=data[sheet_n]['角色'][i]+str(int(data[sheet_n].iloc[i][2*j+2]))+' '
         
        for i in range(len(res)):
            """lis3 = []
            for j in range(len(res['角色'][i])):
                if res['角色'][i][j] not in lis3:
                    res['备用'][i]+=(res['角色'][i][j]+str(res['角色'][i].count(res['角色'][i][j]))+' ')
                    lis3.append(res['角色'][i][j])"""
            res['角色'][i] = res['备用'][i]     
            #print(res['角色'][i])
        
        #res['角色'][i] = res['备用'][i]
        res=res.drop(columns='备用')

        for i in range(len(res)):
            res['肾款'][i]='%.7g' % res['肾款'][i]

    if mode == '2':
        a = 0
        for m in range(len(data)):
            sheet_n = input("请告诉我第"+str(m+1)+"张sheet的名字是（默认请输入‘Sheet1’）：")
            num = (len(data[sheet_n].iloc[0])-3)
            for i in range(len(data[sheet_n])):
                for j in range(int(num)):
                    lis1 = list(res['昵称'])
                    lis2 = list(res['角色'])
                    if data[sheet_n].iloc[i][j+1] not in lis1 and str(data[sheet_n].iloc[i][j+1]) != 'nan':
                        res=res.append({'角色':1},ignore_index=True)
                        res['昵称'][a]=data[sheet_n].iloc[i][j+1]
                        res['角色'][a]=[data[sheet_n]['角色'][i]]
                        res['总点数'][a]=1
                        res['肾款'][a]=(data[sheet_n]['均价'][i]+data[sheet_n]['调价'][i])
                        #res['备用'][a]=data[str(m+1)].iloc[:,0][i]+','
                        a += 1
                    if data[sheet_n].iloc[i][j+1] in lis1 and str(data[sheet_n].iloc[i][j+1]) != 'nan':
                        b = lis1.index(data[sheet_n].iloc[i][j+1])
                        res['肾款'][b] += ((data[sheet_n]['均价'][i]+data[sheet_n]['调价'][i]))
                        res['角色'][b].append(data[sheet_n]['角色'][i])
                        res['总点数'][b]+=1
                        #res['备用'][b]+= (data[str(m+1)].iloc[:,0][i]+',')
                        
        for i in range(len(res)):   
            res['备用'][i]=''
            
        for i in range(len(res)):
            lis3 = []
            for j in range(len(res['角色'][i])):
                if res['角色'][i][j] not in lis3:
                    res['备用'][i]+=(res['角色'][i][j]+str(res['角色'][i].count(res['角色'][i][j]))+' ')
                    lis3.append(res['角色'][i][j])
            res['角色'][i] = res['备用'][i]     
            #print(res['角色'][i])
    
        res=res.drop(columns='备用')

        for i in range(len(res)):
            res['肾款'][i]='%.7g' % res['肾款'][i]
    return res

def auto_return(name1,mode1,name2,mode2):
    pre = auto_table(name1,mode1)
    post = auto_table(name2,mode2)
    result = pandas.DataFrame(columns=('昵称','角色','已肾','应肾','退补'))
    nickname = list(set(pre['昵称']).union(set(post['昵称'])))
    for l in range(len(nickname)):
        a = 0
        for i in range(len(post['昵称'])):
            if nickname[l]==post['昵称'][i]:
                dataf2 = pandas.DataFrame([[post['昵称'][i], post['角色'][i],'0',post['肾款'][i],'0']], columns=['昵称','角色','已肾','应肾','退补'])
                result=result.append(dataf2, ignore_index=True)
            else:
                a += 1
        if a == len(post['昵称']):
            dataf2 = pandas.DataFrame([[nickname[l], '无','0','0','0']], columns=['昵称','角色','已肾','应肾','退补'])
            result=result.append(dataf2, ignore_index=True)
    for j in range(len(result['昵称'])):
        for i in range(len(pre['昵称'])):
            if pre['昵称'][i]==result['昵称'][j]:
                result['已肾'][j]=pre['肾款'][i]
    
    for m in range(len(result['昵称'])):
        #print(result['应肾'][m])
        result['退补'][m]=float(result['应肾'][m])-float(result['已肾'][m])
    return result

print('注意：请把程序和待处理文件置于同一文件夹下')
func = input('请选择功能（肾表/退补）:')
if func == '肾表':
    name_1 = input('请在此输入文件名:')
    mode_1 = input('请选择模式（输入1/2）:')
    res=auto_table(name_1,mode_1)
    #res.insert(len(res.iloc[0]),'昵称 ',res['昵称'])
    res.to_excel(address+'/'+name_1+'肾表.xlsx',encoding='utf-8-sig')
    book = xlsxwriter.Workbook(address+'/'+name_1+'肾表.xlsx')
    sheet = book.add_worksheet('demo')
    bold = book.add_format({
        'bold':  False,  # 字体加粗
        'border': 1,  # 单元格边框宽度
        'align': 'left',  # 水平对齐方式
        'valign': 'vcenter',  # 垂直对齐方式
        'fg_color': 'white',  # 单元格背景颜色
        'text_wrap': True,  # 是否自动换行
        'font_name': '宋体 (正文)'
    })
    width = np.max(res['总点数'])*3
    sheet.set_column(1,1, width)
    sheet.write_row("A1",['昵称','角色','总点数','肾款','昵称'],bold)
    sheet.write_column("A2",res['昵称'],bold)
    sheet.write_column("B2",res['角色'],bold)
    sheet.write_column("C2",res['总点数'],bold)
    sheet.write_column("D2",res['肾款'],bold)
    sheet.write_column("E2",res['昵称'],bold)
    today=time.strftime('%Y-%m-%d',time.localtime(time.time()+86400*7))
    text = '请备注：'+name_1+'+cn\n截止日期：'+today+' 24:00\n微信支付请每100元+0.1元手续费'
    options = {
        'width':256,
        'height':100,
    }
    sheet.insert_textbox(0,5,text,options)
    sheet.insert_image('F6','./二维码.png', {'x_offset': 1, 'y_offset': 0,'x_scale': 0.6, 'y_scale': 0.6})
    book.close()
if func == '退补':
    name_1 = input('请在此输入预排文件名:')
    mode_1 = input('请选择模式（输入1/2）:')
    name_2 = input('请在此输入结果文件名:')
    mode_2 = input('请选择模式（输入1/2）:')
    res=auto_return(name_1,mode_1,name_2,mode_2)
    res.to_excel(address+'/'+name_1+'退补.xlsx',encoding='utf-8-sig')
    book = xlsxwriter.Workbook(address+'/'+name_1+'退补.xlsx')
    sheet = book.add_worksheet('demo')
    bold1 = book.add_format({
        'bold':  False,  # 字体加粗
        'border': 1,  # 单元格边框宽度
        'align': 'left',  # 水平对齐方式
        'valign': 'vcenter',  # 垂直对齐方式
        'fg_color': 'white',  # 单元格背景颜色
        'text_wrap': True,  # 是否自动换行
        'font_name': '宋体 (正文)'
    })
    bold2 = book.add_format({
        'bold':  False,  # 字体加粗
        'border': 1,  # 单元格边框宽度
        'align': 'left',  # 水平对齐方式
        'valign': 'vcenter',  # 垂直对齐方式
        'fg_color': '#ffb3b3',  # 单元格背景颜色
        'text_wrap': True,  # 是否自动换行
        'font_name': '宋体 (正文)'
    })
    #width = np.max(res['应肾'])/3
    width = 40
    sheet.set_column(1,1, width)
    sheet.write_row("A1",['昵称','角色','已肾','总肾','退补','昵称'],bold1)
    sheet.write_column("A2",res['昵称'],bold1)
    sheet.write_column("B2",res['角色'],bold1)
    sheet.write_column("C2",res['已肾'],bold1)
    sheet.write_column("D2",res['应肾'],bold1)
    sheet.write_column("E2",res['退补'],bold2)
    sheet.write_column("F2",res['昵称'],bold1)
    today=time.strftime('%Y-%m-%d',time.localtime(time.time()+86400*7))
    text = '请备注：'+name_1+'+cn\n截止日期：'+today+' 24:00\n微信支付请每100元+0.1元手续费\n*红色为应补足/退款部分'
    options = {
        'width':256,
        'height':100,
    }
    sheet.insert_textbox(0,6,text,options)
    sheet.insert_image('G6','./二维码.png', {'x_offset': 1, 'y_offset': 0,'x_scale': 0.6, 'y_scale': 0.6})
    book.close()

