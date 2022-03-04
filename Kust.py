"""Print of real-time temperature, stirring speed and oxygen saturation data on screen and save in Excel file.

Usage:
    ./stationinfo.py

Author:
    Ursina Gnädinger - 27.02.2022

License:
    MIT License

    Copyright (c) 2022 Ursina Gnädinger

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.
"""

# Import of packages
import re
import sys
import time
import serial
import xlsxwriter
import signal
import datetime
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import matplotlib.dates as mdates


DEBUG_KUST = True
''' Set to True for debug messages to console'''

#-------------------------------------------------------------------------------
class Kust():
    '''Kust (Kommunikation & Steuerung)'''
    comm = None

    def connect(self, device):
        '''Connect to Kust
        
        Parameter: device eg. 'COM1'
        Return: True if successful
        '''
        try:
            self.comm = self.SerialCommunication(device)
            return True
        except serial.SerialException as e:
            self.kust_debug(e.args)
            return False

    def kust_debug(self, resp):
        '''Print a debug string to stdout if globally enabled'''
        if DEBUG_KUST:
            print(f'DEBUG_KUST> {resp}')

    def check_response(self, resp, cmd):
        '''Check if the response contains the correct command and the error code 00'''
        if resp['Command'] != cmd or resp['ErrCode'] != '00':
            self.kust_debug(resp)
            return False
        return True

    def is_raedy(self):
        '''Checks whether the interface box is available
        
        Return: True if Kust is available
        '''
        if self.comm == None:
            self.kust_debug('Not connected')
            return False
        return self.comm.is_raedy()

    def get_firmware_version(self):
        '''Return Version srting, empty if no communication'''
        resp = self.comm.req_resp('IBRF')
        if not self.check_response(resp, 'RF'):
            return ''           
        return resp['Value']

    def get_temperatures(self):
        '''Read all 4 Temperature Sensors

        Return: Array, n=4, type=float, unit=°C
        '''
        t = [None]*4
        for i in range(4):
            resp = self.comm.req_resp('IBRT', i+1)
            if not self.check_response(resp, 'RT'):
                return []
            t[i] = float(resp['Value'])/10
        return t

    def get_rotational_frequency(self):
        '''Read all 6 rotational frequencies

        Return: Array, n=6, type=int, unit=rpm
        '''
        s = [None]*6
        for i in range(6):
            resp = self.comm.req_resp('IBRR', i+1)
            if not self.check_response(resp, 'RR'):
                return []
            s[i] = int(resp['Value'])
        return s

    def get_oxigen_sensor(self):
        '''Read current of the oxigen sensor

        Return: Float, unit=mA
        '''
        resp = self.comm.req_resp('IBRI')
        if not self.check_response(resp, 'RI'):
            return []
        return float(resp['Value'])/1000

    def reset_errors(self):
        '''Reset pending errors'''
        resp = self.comm.req_resp('IBEI')
        self.check_response(resp, 'EI')

    #- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    class SerialCommunication():
        '''Communication class to interface box'''

        def __init__(self, device):
            '''Initialize the Kust object but don't open the searial port.
            This has to be done separately using open_port'''
            self.ser = serial.Serial(
                port     = device,
                baudrate = 19200,
                parity   = serial.PARITY_EVEN,
                stopbits = serial.STOPBITS_ONE,
                bytesize = serial.EIGHTBITS,
                timeout  = 2 # seconds
            )

            # Regular Expression to parse response
            self.pattern_inc_value = re.compile(r'IB(\w{2})er(\d{2})\w*\s?\+?(\d{5}|.*)')

        def open_port(self):
            '''Open the serial port if not already'''
            if not self.ser.isOpen():
                self.ser.open()

        def req_resp(self, req, nbr=''):
            '''Send a request via serial port wait for the response

            The complete request is build up by req and a additional
            number nbr at the end of the command

            Parameter: req
            Parameter: nbr
            Return: Dictionary (empty if not successful)
            '''
            try:
                self.open_port()
                #self.ser.flush()
                byte_str = f'{req}{nbr}\r\n'.encode(encoding = 'UTF-8')
                self.ser.write(byte_str)
                resp = self.ser.readline().decode('UTF-8')
                return self.parse_response(resp)
            except:
                return {'Command':'', 'ErrCode':'', 'Value':''} 

        def is_raedy(self):
            '''Checks whether the interface box is available'''
            resp = self.req_resp('IBRF')
            if resp['Command'] == 'RF': # Intentionally no check for error codes
                return True
            return False

        def parse_response(self, resp):
            '''Convertrs a string response into a dictionary'''
            parsed = self.pattern_inc_value.match(resp)
            if parsed != None:
                return { 'Command':parsed[1], \
                        'ErrCode':parsed[2], \
                        'Value':parsed[3] }
            else:
                return {'Command':'', 'ErrCode':'', 'Value':''}

#-------------------------------------------------------------------------------
# Test Program
#-------------------------------------------------------------------------------
if __name__ == "__main__":
    PORT = 'COM1'

    kust = Kust()
    if not kust.connect(PORT):    # Package pyserial wird benoetigt
        sys.exit()
   
    kust.reset_errors()
    print( kust.get_firmware_version() )
    print('\n\n\n')
    
    rows=1
    now = datetime.datetime.now()
    date = '{}-{}-{}'.format(now.day , now.month,now.year)
    filename = 'Measures' + '_' + date + '.xlsx'
    
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()
    format1 = workbook.add_format({'border': 1})
    format2 = workbook.add_format({'border': 1, 'bold':True})
    format1.set_align('center')
    format2.set_align('center')
    worksheet.merge_range('A1:A2', 'Time',format2)
    worksheet.merge_range('B1:E1', 'Temperature',format2)
    worksheet.write('B2', 'T1',format1)
    worksheet.write('C2', 'T2',format1)
    worksheet.write('D2', 'T3',format1)
    worksheet.write('E2', 'T4',format1)
    worksheet.merge_range('G1:G2', 'Time',format2)
    worksheet.merge_range('H1:H2', 'Oxygen',format2)
    worksheet.merge_range('J1:J2', 'Time',format2)
    worksheet.merge_range('K1:P1', 'Stirring',format2)
    worksheet.write('K2', 'Stirring 1',format1)
    worksheet.write('L2', 'Stirring 2',format1)
    worksheet.write('M2', 'Stirring 3',format1)
    worksheet.write('N2', 'Stirring 4',format1)
    worksheet.write('O2', 'Stirring 5',format1)
    worksheet.write('P2', 'Stirring 6',format1)
  
    run = True

    def signal_handler(signal, frame):
        global run
        print ("exiting")
        run = False

    signal.signal(signal.SIGINT, signal_handler)
    
   
    plt.ion()
    
    x=np.array([])
    y=np.array([])
    y2=np.array([])
    y3=np.array([])
    y4=np.array([])
    y5=np.array([])
    
    fig = plt.figure(figsize=(9, 6))
    plt.subplots_adjust(wspace= 0.25, hspace= 0.25)
    sub1 = fig.add_subplot(2,2,1) 
    sub2 = fig.add_subplot(2,2,2)
    sub3 = fig.add_subplot(2,2,(3,4))
    
    

  
    
    while run :
        t1 = time.time()
        temperatures = kust.get_temperatures()
        speeds = kust.get_rotational_frequency()
        oxigen = kust.get_oxigen_sensor()
        t2 = time.time()
        print( time.strftime('%T') )
        print( f'Temp\t[°C]:\t{temperatures}' )
        print( f'Speed\t[rpm]:\t{speeds}' )
        print( f'Oxigen\t[mA]:\t{oxigen}' )
        print( f'\tdt={(t2-t1):.3f} [s]') #Time spent for requests
        print('\n\n\n')
        time.sleep(0.1)
        rows+=1
        #####Just for tests#########################
        if temperatures==[]: temperatures=[0,0,0,0]
        ############################################
        
        # Write excel file
        worksheet.write(rows, 0, time.strftime('%T'),format1)
        worksheet.write(rows, 1, temperatures[0],format1)
        worksheet.write(rows, 2, temperatures[1],format1)
        worksheet.write(rows, 3, temperatures[2],format1)
        worksheet.write(rows, 4, temperatures[3],format1)
        worksheet.write(rows, 6, time.strftime('%T'),format1)
        worksheet.write(rows, 7, oxigen,format1)
        worksheet.write(rows, 9, time.strftime('%T'),format1)
        worksheet.write(rows, 10, speeds[0],format1)
        worksheet.write(rows, 11, speeds[1],format1)
        worksheet.write(rows, 12, speeds[2],format1)
        worksheet.write(rows, 13, speeds[3],format1)
        worksheet.write(rows, 14, speeds[4],format1)
        worksheet.write(rows, 15, speeds[5],format1)
        new_x=time.strftime('%T')
        
        # oxygen plot
        sub1.set_title('Oxygen Measurement', fontweight='bold')
        sub1.set_ylabel('% DO')
        new_y=oxigen;
        x=np.append(x,new_x)
        y=np.append(y,new_y)
        line, =sub1.plot(x,y,'red')
        sub1.autoscale('x')
        sub1.set_ylim([0,100])
        tick=[x[0],x[-1]]
        sub1.set_xticks(tick,tick)
        sub1.relim()
       
        
       
        # Temperature
        y2=np.append(y2,temperatures[0])
        y3=np.append(y3,temperatures[1])
        y4=np.append(y4,temperatures[2])
        y5=np.append(y5,temperatures[3])
        sub2.plot(x,y2,'blue')
        sub2.plot(x,y3,'red')
        sub2.plot(x,y4,'green')
        sub2.plot(x,y5,'yellow')
        sub2.legend(['T1','T2','T3','T4'])
        sub2.set_title('Temperature Measurement', fontweight='bold')
        sub2.set_ylabel('°C')
        sub2.autoscale('both')
        sub2.set_ylim([20,25])
        sub2.set_xticks(tick,tick)
        sub2.relim()
       
        # Speed
        speedStirring=[speeds[0],speeds[1],speeds[2],speeds[3],speeds[3],speeds[5]]
        motors=['Stirrer 1','Stirrer 2','Stirrer 3','Stirrer 4','Stirrer 5','Stirrer 6']
        bar=sub3.bar(motors,speedStirring)
        bar[0].set_color('blue')
        bar[1].set_color('brown')
        bar[2].set_color('green')
        bar[3].set_color('yellow')
        bar[4].set_color('grey')
        bar[5].set_color('violet')
        sub3.set_title('Stirring Speed Recording', fontweight='bold')
        sub3.set_ylabel('Speed [rpm]')
        sub3.autoscale()
        
        plt.pause(0.8)
        plt.show()

       
    workbook.close()
  
  
    


    
  
    
  
    
  
    
  
    
  
  