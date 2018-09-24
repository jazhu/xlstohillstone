# -*- coding: utf-8 -*-
import xlrd
import threading
import os
import datetime
import wx
import wx.xrc

###########################################################################
## Class ui
###########################################################################

class xlsui ( wx.Frame ):

    def __init__( self, parent ):
        wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title ='excle配置生成工具v0.1', pos = wx.DefaultPosition, size = wx.Size( 1000,680 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL )

        self.SetSizeHints( wx.DefaultSize, wx.DefaultSize )
        self.SetBackgroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_WINDOW ) )

        table = wx.GridBagSizer( 0, 0 )
        table.SetFlexibleDirection( wx.BOTH )
        table.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )

        self.file = wx.FilePickerCtrl( self, wx.ID_ANY, wx.EmptyString, u"Select a file", u"*.*", wx.DefaultPosition, wx.Size( 400,-1 ), wx.FLP_DEFAULT_STYLE )
        self.file.SetBackgroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_INACTIVECAPTION ) )

        table.Add( self.file, wx.GBPosition( 1, 1 ), wx.GBSpan( 1, 1 ), wx.ALL, 5 )

        self.m_staticText1 = wx.StaticText( self, wx.ID_ANY, u"开始行：", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText1.Wrap( -1 )
        table.Add( self.m_staticText1, wx.GBPosition( 1, 2 ), wx.GBSpan( 1, 1 ), wx.ALL, 5 )

        self.output = wx.Button( self, wx.ID_ANY, u"导出", wx.DefaultPosition, wx.DefaultSize, wx.NO_BORDER )
        self.output.SetBackgroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_INACTIVECAPTION ) )

        table.Add( self.output, wx.GBPosition( 1, 5 ), wx.GBSpan( 1, 1 ), wx.ALL, 5 )

        self.create = wx.Button( self, wx.ID_ANY, u"生成配置", wx.DefaultPosition, wx.DefaultSize, wx.NO_BORDER )
        self.create.SetBackgroundColour( wx.SystemSettings.GetColour( wx.SYS_COLOUR_INACTIVECAPTION ) )

        table.Add( self.create, wx.GBPosition( 1, 4 ), wx.GBSpan( 1, 1 ), wx.ALL, 5 )

        self.row = wx.SpinCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size( 80,-1 ), wx.SP_ARROW_KEYS, 0, 100, 5 )
        table.Add( self.row, wx.GBPosition( 1, 3 ), wx.GBSpan( 1, 1 ), wx.ALL, 5 )

        self.config = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size( 960,400 ), wx.TE_MULTILINE )
        table.Add( self.config, wx.GBPosition( 2, 1 ), wx.GBSpan( 1, 5 ), wx.ALL, 5 )

        self.info = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size( 960,160 ), wx.TE_MULTILINE )
        table.Add( self.info, wx.GBPosition( 3, 1 ), wx.GBSpan( 1, 5 ), wx.ALL, 5 )


        self.SetSizer( table )
        self.Layout()

        self.Centre( wx.BOTH )

        # Connect Events
        self.output.Bind( wx.EVT_BUTTON, self.out )
        self.create.Bind( wx.EVT_BUTTON, self.cr )


    def __del__( self ):
        pass


    # Virtual event handlers, overide them in your derived class
    def out( self, event ):
        self.getconfig()

    def cr( self, event ):
        t=threading.Thread(target=self.create_config, args=(self.file.GetPath(),self.row.GetValue(),))
        t.start()

    def create_config(self,x,r):
        workbook = xlrd.open_workbook(x)
        sheetbook= workbook.sheet_names()
        self.info.AppendText('正在获取表信息....')
        b=1
        while len(sheetbook)>b:
            sh=workbook.sheet_by_name(sheetbook[b])
            self.info.AppendText('正在读取'+sheetbook[b]+'\n')
            self.config.AppendText('#######################'+sheetbook[b]+'############################\n')
            rownum=sh.nrows
            i=r
            while rownum>i:
                try:
                    row=sh.row_values(i)
                    des=row[2].split('访问')
                    srcip=row[4].split('\n')
                    dstip=row[8].split('\n')
                    ser=row[10].split('\n')
                    zone=row[11].split('->')
    
                    rule='rule\n action permit\n src-zone '+zone[0]+'\n dst-zone '+zone[1]+'\n'
                    #源地址
                    src_addr=''
                    for sip in srcip:
                        src_addr+='address '+des[0]+'-'+sip+'\n ip '+sip+'/32\nexit\n'
                        rule+=' src-addr '+des[0]+'-'+sip+'\n'
                    #目标地址
                    dst_addr=''
                    for dip in dstip:
                        dst_addr+='address '+des[1]+'-'+dip+'\n ip '+sip+'/32\nexit\n'
                        rule+=' dst-addr '+des[1]+'-'+dip+'\n'
                    #生成服务配置  
                    service=''
                    for s in ser:
                        service+='service '+s+'\n '
                        port=s.split('-')
                        #判断是否为端口范围
                        if len(port)>2:
                            service+=port[0]+' dst-port '+port[1]+' '+port[2]+'\n'
                        else:
                            service+=port[0]+' dst-port '+port[1]+'\n'
                        rule+=' service '+s+'\n'
    
                    service+='exit\n'
                    rule+='exit\n'
                    
                    self.config.AppendText(src_addr)
                    self.config.AppendText(dst_addr)
                    self.config.AppendText(service)
                    self.config.AppendText(rule)
                    i+=1
                except:
                    self.info.AppendText('配置生成失败：'+sheetbook[b]+' '+row[2]+'\n')
                    i+=1
            b+=1

    def getconfig(self):
        config=self.config.GetValue()
        nowTime=datetime.datetime.now().strftime('%Y%m%d%H%M%S')
        newconfig=open('hillstone_config_'+nowTime+'.txt','wb')
        newconfig.write(config.encode())
        newconfig.close()
        self.info.AppendText('文件导出成功！配置文件路径在：'+os.getcwd()+'\\hillstone_config'+nowTime+'.txt')
            
if __name__ == '__main__':
    app = wx.App(False)
    frame= xlsui(None)
    frame.Show()
    app.MainLoop()
    pass