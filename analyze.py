import gpxpy
import gpxpy.gpx

import numpy as np
import matplotlib.pyplot as plt
import pandas as pd

from datetime import timedelta
from datetime import datetime

import csv
import xlwt

import math
import os

# fn = 'C:/Users/fw552fw131/Documents/projects/mariagps/data/2016-08-10 08.57.35 Giorno.gpx'
fn = 'C:/Users/fw552fw131/Documents/projects/mariagps/data/2016-08-11 07.27.02 Giorno.gpx'


def read_gpx(fn):
    gpx_file = open(fn, 'r')
    gpx = gpxpy.parse(gpx_file)
    pts = []

    for track in gpx.tracks:
        for segment in track.segments:
            segment.points[0].speed = 0
            segment.points[-1].speed = 0
            for p in segment.points:
                p.idx = len(pts)
                pts.append(p)
        
    gpx.add_missing_speeds()        
    n = len(pts)

    # post processing
    timeshift = timedelta(hours=8)
    for i in range(n):        
        pts[i].t = (pts[i].time-pts[0].time).total_seconds()
        pts[i].rtime = pts[i].time + timeshift

    return pts


def tag_sections(pts):
    # cars
    n = len(pts)
    for i in range(n):
        pts[i].tag = 3 if pts[i].speed*3.6>10 else 0
            

    #fill car gaps        
    i = 1
    while i<n:
        if pts[i-1].tag==3 and pts[i].tag!=3:
            j = i
            while j<n and pts[j].tag!=3: j+=1
            if j>=n: break
            assert pts[i-1].tag==3 and pts[j].tag==3
            if pts[j].t-pts[i].t<60*5:
                for k in range(i,j): pts[k].tag=3
            i=j
        else:
            i+=1

    # trova le soste
    i = 0
    while i<n:
        j = i
        while j+1<n and pts[j+1].tag!=3 and pts[i].distance_2d(pts[j+1])<50: j+=1
        if j-i<4 or pts[j].t-pts[i].t < 5*60:
            # sosta troppo breve
            if j>i: i=j
            else: i+=1
            continue
        else:
            for r in range(i,j): pts[r].tag = 1
            i = j

    # camminate
    i = 0
    while True:
        # cerco una possibile camminata
        while i<n and pts[i].tag!=0: i+=1
        if i>=n: break
        #gli estremi
        a = i
        b = a
        while b+1<n and pts[b+1].tag==0: b+=1

        assert a<=b<n
        assert pts[a].tag==0 and pts[b].tag==0
        assert a==0 or pts[a-1].tag!=0
        assert b+1>=n or pts[b+1].tag!=0
        
        if a==b or (pts[b].t-pts[a].t)<120*2 and pts[b].distance_2d(pts[a])<50:
            # camminata insignificante
            if a>0 and pts[a-1].tag==3 or b+1<n and pts[b+1].tag==3:
                for j in range(a,b+1): pts[j].tag = 3
            elif a>0 and pts[a-1].tag==1 or b+1<n and pts[b+1].tag==1:
                for j in range(a,b+1): pts[j].tag = 1
            else:
                #??            
                for j in range(a,b+1): pts[j].tag = 2
        else:
            for j in range(a,b+1): pts[j].tag = 2
        i = b+1

def get_average_position(pts, a, b):
    lat,lon=0,0
    for j in range(a,b+1):
        lat += pts[a].latitude
        lon += pts[a].longitude
    m = b-a+1
    return lat/m,lon/m

def format_position(pos):
    return "%.7f, %.7f"%pos

def get_average_speed(pts, a,b):
    speed = 0
    for j in range(a,b+1):
        speed += pts[j].speed
    return speed/(b-a+1)

def get_total_distance(pts, a,b):
    s = 0
    for j in range(a,b):
        s += pts[j].distance_2d(pts[j+1])
    return s

    
wb = xlwt.Workbook()

base = 'font: name Arial, height 100;'
pat_white = ''
pat_yellow = 'pattern: pattern solid, fore_colour yellow;'
pat_green = 'pattern: pattern solid, fore_colour light_green;'

align_left = 'align: horiz left;'
align_right = 'align: horiz right;'

styles = {}
styles[1] = xlwt.easyxf(base + pat_yellow + align_left)
styles[2] = xlwt.easyxf(base + pat_green + align_left)
styles[3] = xlwt.easyxf(base + pat_white + align_left)
styles[-1] = xlwt.easyxf(base + pat_yellow + align_right)
styles[-2] = xlwt.easyxf(base + pat_green + align_right)
styles[-3] = xlwt.easyxf(base + pat_white + align_right)

style0 = xlwt.easyxf('font: name Times New Roman, bold on',
    num_format_str='#,##0.00')
style1 = xlwt.easyxf()
yellow = xlwt.easyxf()


column_titles = ["From","To","Begin","End","Duration","Type","Speed","Distance","Pos"]

sheet_count = 0

folder = 'C:/Users/fw552fw131/Documents/projects/mariagps/data/'

for file in os.listdir(folder):
    if not file.endswith(".gpx"): continue
    
    print("Processing ",file)
    fn = folder + file

    pts = read_gpx(fn)
    if len(pts)<50: continue
    tag_sections(pts)
    day = pts[len(pts)//2].rtime.strftime("%d.%b.%Y")

    sheet_count += 1
    ws = wb.add_sheet(day)

    ws.write(0,0,"Date")
    ws.write(0,1,pts[len(pts)//2].rtime.strftime("%A, %d %B %Y"))

    ws.write(1,0,"File")
    ws.write(1,1,file)
    
    ws.write(2,0,"Gps points")
    ws.write(2,1,len(pts))

    for col in range(len(column_titles)):
        ws.write(4, col, column_titles[col],style0)

    ws.col(7).width = 256*12
    ws.col(8).width = 256*25

    row = 5
    n = len(pts)
    i = 0
    while i<n:
        a = i
        b = i
        while b+1<n and pts[b+1].tag==pts[a].tag: b+=1
        tag_str = ["?","Rest","Walk","Car"][pts[a].tag]

        tag = pts[a].tag
        st = styles[tag]
        st_r = styles[-tag]
        
        col = 0
        ws.write(row, col, a,st_r); col+=1
        ws.write(row, col, b,st_r); col+=1
        
        ws.write(row, col, pts[a].rtime.strftime("%H:%M:%S"),st); col+=1
        ws.write(row, col, pts[b].rtime.strftime("%H:%M:%S"),st); col+=1
        ws.write(row, col, str(pts[b].rtime-pts[a].rtime),st); col+=1

        ws.write(row, col, tag_str,st); col+=1


        if pts[a].tag == 1:
            # sosta
            ws.write(row, col, "",st); col+=1
            ws.write(row, col, "",st); col+=1
            pos = format_position(get_average_position(pts,a,b))
            ws.write(row, col, pos,st); col+=1
        else:
            

            val = "%.0f km/h"%get_average_speed(pts,a,b)    
            ws.write(row, col, val,st_r); col+=1

            if pts[a].tag == 2:
                val = "%.0f m"%(get_total_distance(pts,a,b) )
            else:
                val = "%.3f km"%(0.001*get_total_distance(pts,a,b) )
                
            ws.write(row, col, val,st_r); col+=1

            ws.write(row, col, "",st); col+=1
        
        row += 1
        
        i = b+1

wb.save('example2.xls')
