import time
from datetime import datetime, timezone
from tkinter import *
import streamlit as st
import pandas as pd
import numpy as np
import serial
import modbus_tk
from fontTools.merge.util import current_time
from modbus_tk import utils
from modbus_tk import modbus
import modbus_tk.defines as cst
import modbus_tk.modbus_rtu as modbus_rtu
from threading import Timer
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.animation as animation
matplotlib.use('TkAgg')
from collections import deque
delay_pu=[]
act_time=[]
data_one=[]

fig=plt.figure()
ax=fig.add_subplot(1,1,1)

def animate(i):
    with open('text.txt', 'r') as f:
         file=deque(f, 20)
    #print(file)
    data=open('text.txt', 'r').read()
    #print(data)
    lines=data.split('\n')[:-1]
    #print(lines)
    xs=[]
    ys=[]

    #for line in lines:
    for line in file:
        print(line)
        y,x=line.split(', ')
        #print(x)
        #print(y)
        xs.append(x)
        #print(xs)
        ys.append(y)
        #print(ys)
    ax.clear()
    ax.plot(xs,ys)
    plt.xticks(rotation=90)

    plt.xlabel('Время')
    plt.ylabel('Параметр')
    plt.title('График')


#ani=animation.FuncAnimation(fig, func=animate, interval=1000)
#plt.show()






def repeater(interval, function):
    Timer(interval, repeater, [interval, function]).start()
    function()

def modbus_req():
    PORT = "COM1"
    # Заголовок приложения
    # st.title('Тестовое приложения для Modbus RTU')

    logger = modbus_tk.utils.create_logger('console')

    master = modbus_rtu.RtuMaster(
        serial.Serial(port=PORT, baudrate=9600, bytesize=8, parity='N', stopbits=1, xonxoff=0)
    )
    master.set_timeout(3.0)
    master.set_verbose(True)
    logger.info('connected')
    f=open('text.txt', 'w' )
    try:
            get_param = master.execute(14, cst.READ_HOLDING_REGISTERS, 514, 12)
            #print(get_param)
            #current_time=current_time.strftime("%d-%m-%Y %H:%M")
            current_time_sec=int(round(time.time()))
            current_time=datetime.fromtimestamp(current_time_sec)
            current_time=current_time.strftime('%Y-%m-%d %H:%M:%S')

            print (current_time)
            delay_pu.append(get_param[0])
            act_time.append(current_time)
            #data_one.append([get_param[0], int(str(current_time)[-2::])])
            data_one.append([get_param[0], str(current_time)[-8::]])
            print(delay_pu)
            print(act_time)
            print(data_one)
            data={
                "delay":delay_pu,
                "time":act_time
            }
            rows_table=list(range(1,len(delay_pu)+1))
            #print(data)
            i=0
            for index in data_one:
                i+=1
                if i==len(data_one):
                    f.write(str(index)[1:-1])
                else:
                    f.write(str(index)[1:-1] + '\n')

            f.close()
            #ani = animation.FuncAnimation(fig, animate, interval=1000)
            #plt.show()

            #print(rows_table)
            #df=pd.DataFrame(data=data, index=rows_table)
            #print(df)
            #df.to_csv('data.csv')



    except modbus_tk.modbus.ModbusError as exc:
            logger.error("%s- Code=%d", exc, exc.get_exception_code())

repeater(2, modbus_req)
ani = animation.FuncAnimation(fig, func=animate, interval=1000)
plt.show()







