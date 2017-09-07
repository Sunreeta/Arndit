import sys, os, csv
import numpy as N
import wx
import os.path
import xlrd
import xlsxwriter
import matplotlib
import numpy
import math
#import Tkinter
#import tkFileDialog
from datetime import datetime
from pylab import *
from scipy import *
from scipy import optimize
from bisect import bisect_left


from matplotlib.figure import Figure
from matplotlib.pyplot import *
from matplotlib.backends.backend_wxagg import \
    FigureCanvasWxAgg as FigCanvas, \
    NavigationToolbar2WxAgg as NavigationToolbar

# global variables
offset_diam = 0
no_of_points=1
x_axis=[]
y_axis=[]
x_axis_val=[]
y_axis_val=[]
flag=0
sno = []
volatile_date=[]
high_val = []
low_val = []
close_val = []
time_diff = []
date_val = []
stock_val=[]
mvg_avg = []
volatile_value = []
mvg_avg_option=[]
rd_path = ""
wr_path = ""
text_path=""
x_name=""
y_name=""
mvg_avg_name=""
correl_val=0.0


class FinalProject(wx.Frame):
    title = 'Arndit_Project'

    def __init__(self):
        global offset_diam

        wx.Frame.__init__(self, None, -1, self.title)

        # al init
        offset_diam = float(1.2)

        self.create_menu()
        self.create_status_bar()
        self.create_main_panel()
# Menubar design
    def create_menu(self):

        self.menubar = wx.MenuBar()
        menu_file = wx.Menu()
        menu_edit = wx.Menu()

        m_load = menu_file.Append(-1, "&Load excel File\tCtrl-L", "Load Raw data from file")  # al
        self.Bind(wx.EVT_MENU, self.on_load_file, m_load)  # al

        m_expt = menu_file.Append(-1, "&Save plot\tCtrl-S", "Save plot to file")
        self.Bind(wx.EVT_MENU, self.on_save_plot, m_expt)
        menu_file.AppendSeparator()
        m_export = menu_file.Append(-1, "&Export data\tCtrl-X", "Export")
        self.Bind(wx.EVT_MENU, self.on_export_data, m_export)

        m_exit = menu_file.Append(-1, "&Exit\tCtrl-X", "Exit")
        self.Bind(wx.EVT_MENU, self.on_exit, m_exit)
        self.menubar.Append(menu_file, "&File")
        self.SetMenuBar(self.menubar)

# mainpanel design
    def create_main_panel(self):

        global offset_diam
        self.panel = wx.Panel(self)
        self.dpi = 100
        self.fig = Figure((15,100), dpi=self.dpi)

        # graph canvas
        self.fig.subplots_adjust(hspace=0.3, wspace=0.3)  # space at the bottom for axes labels
        self.canvas = FigCanvas(self.panel,-1, self.fig)
        self.axes = self.fig.add_subplot(1,1,1)
        self.axes.set_xlabel('X-Axis')  # al
        self.axes.set_ylabel('Y-Axis')  # al

        # control Buttons
        self.refresh = wx.Button(self.panel,-1, "Display Graph", size=(200,25))
        self.Bind(wx.EVT_BUTTON, self.on_draw_refresh, self.refresh)
        self.reset = wx.Button(self.panel, -1, "Reset Data", size=(200,25))
        self.Bind(wx.EVT_BUTTON, self.on_draw_reset, self.reset)
        self.displayButton = wx.Button(self.panel, -1, "Display Data", size=(200,25))
        self.Bind(wx.EVT_BUTTON, self.on_draw_display, self.displayButton)

        # functional buttons
        self.movingAvg = wx.Button(self.panel, -1, "Apply Moving Average", size=(2,25))
        self.Bind(wx.EVT_BUTTON, self.on_draw_movingAvg, self.movingAvg)
        self.checkVolatility = wx.Button(self.panel, -1, "Volatility check",size=(200,25))
        self.Bind(wx.EVT_BUTTON, self.on_draw_volatile, self.checkVolatility)
        self.correlation = wx.Button(self.panel, -1, "Apply Correlation",size=(250,25))
        self.Bind(wx.EVT_BUTTON, self.on_draw_correlation, self.correlation)

        # Text Panel
        self.my_text = wx.TextCtrl(self.panel,-1,style=wx.TE_MULTILINE,size=(400,500))


        # Drop down list
        correlation_option = ['high', 'low', 'close', 'stock', 'volatility']
        self.x_box = wx.ComboBox(self.panel, choices=correlation_option, size=(150, 20))
        self.text_xbox = wx.StaticText(self.panel, -1, 'Parameter 1 :      ', size=(100, 22))
        self.y_box = wx.ComboBox(self.panel, choices=correlation_option, size=(150, 20))
        self.text_ybox = wx.StaticText(self.panel, -1, 'Parameter 2 :      ', size=(100, 22))

        self.moving_avg_box=wx.ComboBox(self.panel,choices=correlation_option,size=(150,20))
        self.text_mvgavg_box=wx.StaticText(self.panel,-1,' Parameter :    ',size=(100,22))

        # moving avg points
        self.avg_points_txt = wx.StaticText(self.panel, label=" No. of Points :    ",size=(100,22))
        self.avg_points_txt_val = wx.TextCtrl(self.panel, value="1",size=(150,20))


        #no of points moving avg
        flags = wx.ALIGN_CENTER


        #text options
        self.movingAvg_Option_text = wx.BoxSizer(wx.HORIZONTAL)
        self.movingAvg_Option_text.Add(self.avg_points_txt, 0, border=10, flag=flags)
        self.movingAvg_Option_text.Add(self.avg_points_txt_val, 0, border=10, flag=flags)


        # moving average options
        self.movingAvg_Option = wx.BoxSizer(wx.HORIZONTAL)
        self.movingAvg_Option.Add(self.text_mvgavg_box, 0, border=3, flag=flags)
        self.movingAvg_Option.Add(self.moving_avg_box, 0, border=3, flag=flags)


        #correlation checkboxes1

        self.correlation_checkbox1=wx.BoxSizer(wx.HORIZONTAL)
        self.correlation_checkbox1.Add(self.text_xbox,0,border=10,flag=flags)
        self.correlation_checkbox1.Add(self.x_box, 0, border=10, flag=flags)

        self.correlation_checkbox2=wx.BoxSizer(wx.HORIZONTAL)
        self.correlation_checkbox2.Add(self.text_ybox, 0, border=10, flag=flags)
        self.correlation_checkbox2.Add(self.y_box, 0, border=10, flag=flags)


        #correlation button
        self.correlation_button=wx.BoxSizer(wx.HORIZONTAL)
        self.correlation_button.Add(self.correlation,0,border=10,flag=flags)



        flags = wx.ALIGN_CENTRE | wx.EXPAND  # wx.ALIGN_LEFT  #
        # taskbarvbox 1
        self.taskbarvbox1 = wx.BoxSizer(wx.VERTICAL)
        self.taskbarvbox1.Add(self.refresh, 0, border=10, flag=flags)
        self.taskbarvbox1.Add(self.reset, 0,border=10, flag=flags)
        self.taskbarvbox1.Add(self.displayButton, 0, border=10, flag=flags)

        # taskbarvbox 2
        self.taskbarvbox2 = wx.BoxSizer(wx.VERTICAL)
        self.taskbarvbox2.Add(self.movingAvg_Option, 0, border=10, flag=flags)
        self.taskbarvbox2.Add(self.movingAvg_Option_text,0,border=10,flag=flags)
        self.taskbarvbox2.Add(self.movingAvg, 0, border=10, flag=flags)

        # taskbarvbox 3
        flags = wx.ALIGN_BOTTOM
        self.taskbarvbox3 = wx.BoxSizer(wx.VERTICAL)
        self.taskbarvbox3.Add(self.checkVolatility, 0, border=10, flag=flags)

        # taskbarvbox 4
        self.taskbarvbox4 = wx.BoxSizer(wx.VERTICAL)
        self.taskbarvbox4.Add(self.correlation_checkbox1, 0, border=10, flag=flags)
        self.taskbarvbox4.Add(self.correlation_checkbox2, 0, border=10, flag=flags)
        self.taskbarvbox4.Add(self.correlation_button, 0, border=10, flag=flags)

        # taskbarvbox 5
        self.taskbarvbox5 = wx.BoxSizer(wx.VERTICAL)
       # self.taskbarvbox5.Add(self.partition, 0, border=10, flag=flags)
        # taskbar
        self.taskbar = wx.BoxSizer(wx.HORIZONTAL)
        self.taskbar.Add(self.taskbarvbox1, 0,border=10,flag=flags)
        self.taskbar.Add(self.taskbarvbox2, 0, border=10, flag=flags)
        self.taskbar.Add(self.taskbarvbox3, 0, border=10, flag=flags)
        self.taskbar.Add(self.taskbarvbox4, 0, border=10, flag=flags)
        self.taskbar.Add(self.taskbarvbox5, 0, border=10, flag=flags)

       # databox
        self.databox = wx.BoxSizer(wx.VERTICAL)
        self.databox.Add(self.my_text, wx.ALL | wx.EXPAND)

        # mainpanel
        flags = wx.ALL|wx.EXPAND
        flags = wx.GROW  # wx.ALIGN_TOP | wx.ALL
        self.mainpanel = wx.BoxSizer(wx.HORIZONTAL)
        self.mainpanel.Add(self.databox, 1, wx.GROW)
        self.mainpanel.Add(self.canvas, 1, wx.GROW )  # al
        self.toolbar = NavigationToolbar(self.canvas)

        # 3 horizontal boxes
        self.vbox = wx.BoxSizer(wx.VERTICAL)
        self.vbox.Add(self.toolbar, 1, wx.TOP | wx.EXPAND , border=0)  # al
        self.vbox.Add(self.taskbar, 1, border=0)
        self.vbox.Add(self.mainpanel, 1, wx.GROW , border=0)  # al

        self.panel.SetSizer(self.vbox)
        self.vbox.Fit(self)

# StatusBar design
    def create_status_bar(self):
        self.statusbar = self.CreateStatusBar()

# Refresh button
    def on_draw_refresh(self, event):
        global flag
        self.axes.clear()
        self.canvas.draw()
        self.fig.clf()
        print "Button 1 Clicked"
        print flag
        global x_axis_val
        global y_axis_val
        global volatile_date
        global x_name
        global y_name

        # vlotaility graph
        if flag==2:
           if (len(sno) > 0):
              print "Drawing plot ..... "
              self.axes = self.fig.add_subplot(1,1,1)

              self.axes.plot(x_axis_val,y_axis_val, marker='o', linestyle='--', color='red')
              self.axes.set_xlabel('date')
              self.axes.set_ylabel('volatile_value')
              # if (len(diameter_mm_smooth) > 0):
             #   self.axes.plot(frame_no, diameter_mm_smooth, 'black')
           self.canvas.draw()

        if flag==3:
            if (len(sno) > 0):
                    print "Drawing plot ..... "
                #if((x_axis==0) & (y_axis==1)):
                   # print x_axis
                    #print y_axis
                    self.axes = self.fig.add_subplot(1, 1, 1)
                    self.axes.plot(x_axis_val, y_axis_val, marker='o', linestyle=' ', color='red')
                    self.axes.set_xlabel(x_name)
                    self.axes.set_ylabel(y_name)
                # if (len(diameter_mm_smooth) > 0):
                #   self.axes.plot(frame_no, diameter_mm_smooth, 'black')

                    self.canvas.draw()


# Display button
    def on_draw_display(self, event):
        temp=[]
        global mvg_avg
        global volatile_value
        global high_val
        global low_val
        global stock_val
        global sno
        global close_val
        global x_axis
        global y_axis
        global flag
        if flag==0:

            temp1 = list(high_val)
            temp2 = list(low_val)
            temp0 = list(date_val)
            temp3 = list(close_val)
            temp4 = list(stock_val)
            #print temp1, temp2
            f = open('myfile_data.txt', 'w')
            f.write('DAY' + '\t\t' + 'HIGH' + '\t' + 'LOW' + '\t' + 'CLOSE' + '\t' + 'STOCK' + '\n')
            for i in range(0, len(sno)):
                f.write(str(temp0[i]) + '\t' + str(temp1[i]) + '\t' + str(temp2[i]) + '\t' + str(temp3[i]) + '\t' + str(
                    temp4[i]) + '\n')
            f.close()
            if os.path.exists('myfile_data.txt'):
              with open('myfile_data.txt') as fobj:
                 for line in fobj:
                     self.my_text.WriteText(line)
                 #print line
#moving average
        if flag==1:
            global mvg_avg_name
            self.my_text.WriteText("moving Average("+mvg_avg_name +")\n")
            for line in mvg_avg:
              self.my_text.WriteText(str(line)+"\n")
# volatility
        if flag==2:
            global volatile_date
            global volatile_value
            temp1=list(volatile_date)
            temp2=list(volatile_value)
            self.my_text.WriteText('Date \t\t Volatility\n')
            for i in range(0,len(volatile_date)):
                self.my_text.WriteText(str(temp1[i])+' \t '+str(temp2[i])+'\n')


            #if os.path.exists('myfile.txt'):
              #with open('myfile.txt') as fobj:
                 #for line in fobj:
                     #self.my_text.WriteText(line)
                 #print line

                 # correlation
        if flag == 3:
            self.my_text.WriteText('Correlation between '+ str(x_name) +' & '+str(y_name)+' is : '+ str(correl_val))
            #if os.path.exists('myfile_correlation.txt'):
              # with open('myfile_correlation.txt') as fobj:
                  #  for line in fobj:
                       #  self.my_text.WriteText(line)
                    #print line


#moving Algorithm
#moving average
    def on_draw_movingAvg(self, event):
        global no_of_points
        global mvg_avg_option
        global mvg_avg_name

        no_of_points=self.avg_points_txt_val.GetValue()
        #print no_of_points
        if (self.moving_avg_box.GetValue() == 'high'):
            mvg_avg_option = high_val
            mvg_avg_name='High Value'
        elif (self.moving_avg_box.GetValue() == 'low'):
            mvg_avg_option = low_val
            mvg_avg_name='Low Value'
        elif (self.moving_avg_box.GetValue() == 'close'):
            mvg_avg_option = close_val
            mvg_avg_name='Close Value'
        elif (self.moving_avg_box.GetValue() == 'stock'):
            mvg_avg_option = stock_val
            mvg_avg_name='Stock Value'
        elif (self.moving_avg_box.GetValue() == 'volatility'):
            volatile_check()
            mvg_avg_option= volatile_value
            mvg_avg_name='Volatility'
        moving_avg_filter()


            # volatile check
    def on_draw_volatile(self,event):
        volatile_check()
        temp1=list(volatile_value)
        temp2=list(volatile_date)
        for i in range(0,len(sno)):
            if int(temp1[i]) != 0 :
                x_axis_val.append(temp2[i])
                y_axis_val.append(temp1[i])
        print x_axis_val
        print y_axis_val
#correlation
    def on_draw_correlation(self,event):
        global x_axis
        global y_axis
        global x_name
        global y_name
        x_name = self.x_box.GetValue()
        y_name=self.y_box.GetValue()
        if(self.x_box.GetValue()== 'high' ):
            x_axis=high_val
        elif(self.x_box.GetValue()== 'low' ):
            x_axis=low_val
        elif(self.x_box.GetValue()== 'close' ):
            x_axis=close_val
        elif(self.x_box.GetValue()== 'stock' ):
            x_axis=stock_val
        elif(self.x_box.GetValue()== 'volatility') :
            volatile_check()
            x_axis=volatile_value

        if (self.y_box.GetValue() == 'high'):
            y_axis = high_val
        elif (self.y_box.GetValue() == 'low'):
            y_axis = low_val
        elif (self.y_box.GetValue() == 'close'):
            y_axis = close_val
        elif (self.y_box.GetValue() == 'stock'):
            y_axis = stock_val
        elif (self.y_box.GetValue() == 'volatility'):
            volatile_check()
            y_axis = volatile_value


        correlation()


    def on_draw_reset(self, event):
        global flag
        reset_all_global()
        flag=0
        FinalProject.on_draw_refresh(self, event)

    def on_load_file(self, event):
        global rd_path
        reset_all_global()
        wr_path = ""
        file_choices = "XLSX (*.xlsx)|*.xlsx"

        dlg = wx.FileDialog(
            self,
            message="Load File...",
            defaultDir=os.getcwd(),
            defaultFile="",
            wildcard=file_choices,
            style=wx.FC_OPEN | wx.FD_FILE_MUST_EXIST)

        if dlg.ShowModal() == wx.ID_OK:
            rd_path = dlg.GetPath()
            print rd_path
            read_file(rd_path)

#            FinalProject.on_draw_button1(self, event)

         #   self.editxt2.SetValue(str(offset_diam))



    def on_save_plot(self, event):
        file_choices = "PNG (*.png)|*.png"

        dlg = wx.FileDialog(
            self,
            message="Save plot as...",
            defaultDir=os.getcwd(),
            defaultFile="plot.png",
            wildcard=file_choices,
            style=wx.FC_SAVE)

        if dlg.ShowModal() == wx.ID_OK:
            path = dlg.GetPath()
            self.canvas.print_figure(path, dpi=self.dpi)
            # self.flash_status_message("Saved to %s" % path)

    def on_export_data(self, event):
        global rd_path
        global wr_path

        file_choices = "XLSX (*.xlsx)|*.xlsx"

        dlg = wx.FileDialog(
            self,
            message="Export Data File...",
            defaultDir=os.getcwd(),
            defaultFile="export_data.txt",
            wildcard=file_choices,
            style=wx.FC_SAVE)

        if dlg.ShowModal() == wx.ID_OK:
            wr_path = dlg.GetPath()
            # self.flash_status_message("File Loaded %s" % path)
            export_data(wr_path)

    def on_exit(self, event):
        self.Destroy()


# Supplementary functions
def read_file(filename):
    reset_all_global()
    global sno
    global date_val
    global high_val
    global low_val
    global close_val
    global offset_diam
    global stock_val

    book=xlrd.open_workbook(filename)
    sheet=book.sheet_by_index(0)
    serial=0
    # data=[[sheet.cell_value(r,c)for c in range (sheet.ncols)]for r in range(sheet.nrows)]
    for r in range(sheet.nrows):
      if sheet.cell_value(r,0)!='':
        if serial!=0:
            date=sheet.cell_value(r,0)
            #print date
            if sheet.cell_value(r,0)!='':
                date_val.append(xlrd.xldate.xldate_as_datetime(int(sheet.cell_value(r,0)), book.datemode))
            else:
                date_val.append(0)
            if sheet.cell_value(r,1)!='':
                high_val.append(int(sheet.cell_value(r,1)))
            else:
                high_val.append(0)
            if sheet.cell_value(r, 2) != '':
                low_val.append(int(sheet.cell_value(r,2)))
            else:
                low_val.append(0)
            if sheet.cell_value(r, 3) != '':
                close_val.append(int(sheet.cell_value(r,3)))
            else:
                close_val.append(0)
            if sheet.cell_value(r, 4) != '':
                stock_val.append(sheet.cell_value(r,4))
            else:
                stock_val.append(0)
            sno.append(serial)
        data_flag = 0
        serial=serial+1

    #print date_val

    offset_diam = sno[0]
   # print "Read from file Offset Diameter : " + str(offset_diam)
   # print "Completed Reading File...."
    # print frame_no

#volatile check
def volatile_check():
    global sno
    global volatile_value
    global high_val
    global low_val
    global volatile_date
    global date_val
    global flag
    global x_axis_val
    flag=2
    temp=[]

    temp1=[]
    temp2=[]
    volatile_value=[]
    temp1= list(high_val)
    temp2=list(low_val)
    temp=list(date_val)
   # print temp1,temp2
    f = open('myfile.txt', 'w')
    f.write('DAY'+'\t\t'+'HIGH' + '\t' + 'LOW' + '\t' + 'VOLATILE' + '\n')
    for i in range(0,len(sno)):
        if (int(temp1[i])!=0 and int(temp2[i])!=0):
            dif=int(temp1[i])-int(temp2[i])
        else:
            dif=0
        #print dif

        f.write(str(temp[i])+'\t'+str(temp1[i])+'\t'+str(temp2[i])+'\t'+str(dif)+'\n')

        volatile_value.append(dif)
        volatile_date.append(temp[i])
    #print volatile_value
    print volatile_value
    f.close()

# moving average filter
def moving_avg_filter():
    global time_diff
    global mvg_avg
    global high_val
    global date_val
    global flag
    global volatile_value
    global no_of_points
    global mvg_avg_option
    temp=[]
    temp2=[]
    mvg_avg=[]
    temp = mvg_avg_option
    #print temp
    flag=1
    for r in range(0,len(temp)):
        if temp[r]!=0 :
            temp2.append(temp[r])
    print temp2

    k = int(no_of_points)
    print k# for k point moving average
    print "Applying " + str(k)+ " Point Moving Average filter...."
    for i in range(0,(len(temp2)-k+1)):
        sum = 0
        # print "original diameter:" + str(diameter_mm_mvg_avg[i])+ "\t"
        for j in range(i,i+k):
            sum = sum + int(temp2[j])

        avg=sum/k
        mvg_avg.append(avg)
        # print "new diameter:" + str(diameter_mm_mvg_avg[i]) + "\n"
        # diameter_mm[i] = (diameter_mm[i-1] + diameter_mm[i] + diameter_mm[i+1])/3;
   # diameter_mm_smooth = list(diameter_mm_mvg_avg)
    #print mvg_avg

def correlation():
    global high_val
    global low_val
    global stock_val
    global sno
    global close_val
    global volatile_value
    global x_axis
    global y_axis
    global flag
    global x_axis_val
    global y_axis_val
    global correl_val
    flag=3
    temp1 = list(high_val)
    temp2 = list(low_val)
    temp0= list(date_val)
    temp3=list(close_val)
    temp4=list(stock_val)
    temp5=list(volatile_value)
    k=0
    sum1=0
    sum2=0
    sum_a_square=0;
    sum_b_square=0;
    sum_ab=0

    for i in range(0, len(sno)):
        if (int(x_axis[i]) != 0 and int(y_axis[i]) != 0):
            k=k+1
            sum1=sum1+x_axis[i]
            sum2=sum2+y_axis[i]
    print k
    avg_val1=sum1/k
    avg_val2=sum2/k

    x_axis_val=[]
    y_axis_val=[]

    for i in range(0, len(sno) ):
        if(int(x_axis[i])!=0 and int(y_axis[i])!=0):
            x_axis_val.append(x_axis[i])
            y_axis_val.append(y_axis[i])

            a_val=int(x_axis[i])-avg_val1
            b_val=int(y_axis[i])-avg_val2
            a_square=a_val*a_val
            b_square=b_val*b_val
            ab_mul=(a_val*b_val)
            sum_ab=sum_ab+ab_mul
            sum_a_square=sum_a_square+a_square
            sum_b_square=sum_b_square+b_square
    #print x_axis_val
    #print y_axis_val
    #print sum_b_square
    #print sum_ab
    div=math.sqrt(sum_a_square*sum_b_square)
    correl_val=sum_ab/div
    print 'The correlation of '+x_name+' & '+ y_name+' is...'
    print correl_val
    #print N.corrcoef(x_axis,y_axis)
    #print temp1, temp2
    f = open('myfile_correlation.txt', 'w')
    f.write('DAY' + '\t\t' + 'HIGH' + '\t' + 'LOW' + '\t' + 'CLOSE' + '\t'+'STOCK'+'\n')
    for i in range(0, len(sno)):
        f.write(str(temp0[i]) + '\t' + str(temp1[i]) + '\t' + str(temp2[i]) + '\t' + str(temp3[i])+'\t'+str(temp4[i]) + '\n')
    f.close()


def export_data(filename):

    global mvg_avg
    global x_axis_val
    global y_axis_val
    global correl_val
    global flag


    global rd_path
    global wr_path
    temp1=[]
    temp2=[]
#    fr=open(wr_path,'r')

    fw = xlsxwriter.Workbook(filename)
    sheet=fw.add_worksheet()
    if flag == 1:
        print "Exporting Data File.....\n"

        temp1=list(mvg_avg)
        sheet.write(0,0, mvg_avg_name)
        for i in range(0,len(mvg_avg)):
            sheet.write(i,0,temp1[i])

    if flag == 2:
        print "Exporting Data File.....\n"

        temp1=list(volatile_date)
        temp2=list(volatile_value)
        sheet.write(0,0, 'Date')
        sheet.write(0,1,'Volatility')

        for i in range(0,len(volatile_date)):
            sheet.write(i,0,temp1[i])
            sheet.write(i,1,temp2[i])

    if flag == 3:
        print "Exporting Data File.....\n"

        temp1=list(x_axis_val)
        temp2=list(y_axis_val)
        sheet.write(0,0, x_name)
        sheet.write(0,1,y_name)
        sheet.write(0,2,'Correlation')
        for i in range(1,len(x_axis_val)):
            sheet.write(i,0,temp1[i])
            sheet.write(i,1,temp2[i])
        sheet.write(1,2,correl_val)

    fw.close()

def reset_all_global():
    global sno
    global date_val
    global high_val
    global low_val
    global close_val
    global offset_diam
    global stock_val
    global x_axis
    global y_axis
    global x_axis_val
    global y_axis_val
    global flag
    global no_of_points
    global volatile_date
    global volatile_value
    global mvg_avg
    global mvg_avg_option
    global rd_path
    global wr_path
    global x_name
    global y_name
    global mvg_avg_name
    global correl_val
    global text_path

    offset_diam = 0
    no_of_points = 1
    x_axis = []
    y_axis = []
    x_axis_val = []
    y_axis_val = []
    flag = 0
    sno = []
    volatile_date = []
    high_val = []
    low_val = []
    close_val = []
    time_diff = []
    date_val = []
    stock_val = []
    mvg_avg = []
    volatile_value = []
    mvg_avg_option = []
    rd_path = ""
    wr_path = ""
    text_path = ""
    x_name = ""
    y_name = ""
    mvg_avg_name = ""
    correl_val = 0.0

if __name__ == '__main__':
    app = wx.App(False)
    app.frame = FinalProject()
    app.frame.Show()
    app.MainLoop()
    del app

