import sys
import os
import pandas as pd
import sqlite3
import time

import numpy as np
from itertools import groupby
from collections import defaultdict
from PyQt5.QtWidgets import *
from PyQt5 import uic
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import BaseDocTemplate, Table, TableStyle, Paragraph, Frame, PageTemplate, PageBreak, SimpleDocTemplate
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont




def displayHistory(self):
    try:
            start_time = time.time()
            global table, headers,result,right,amt,gst,total,date, saleman,addedby,invnum,quote, name, sadd1, sadd2, sadd3, scode, sphone, spstno, lstreet, lname, ladd1, ladd2, ladd3, lcode
            conn = sqlite3.connect(':memory:')
            cur = conn.cursor()

            # load excel file
            dfi = pd.read_excel("Data/pri_hd.xlsx", sheet_name="pri_hd")            
            dfi.to_sql(name='pri_hd', con=conn, if_exists='append')
            dfq = pd.read_excel("Data/prq_hd.xlsx", sheet_name="prq_hd")
            dfq.to_sql(name='prq_hd', con=conn, if_exists='append')
            dfq = pd.read_excel("Data/prw_hd.xlsx", sheet_name="prw_hd")
            dfq.to_sql(name='prw_hd', con=conn, if_exists='append')

            #get table
            
            cur.execute("SELECT 'INV',pri_hd.OINO, ONAME||COALESCE(OFNAME,' ') FROM pri_hd WHERE pri_hd.PR_STATUS = 'P'")
            num_i = cur.fetchall()
            cur.execute("SELECT 'QU',prq_hd.OQNO, ONAME||COALESCE(OFNAME,' ') FROM prq_hd WHERE prq_hd.PR_STATUS = 'P'")
            num_q = cur.fetchall()
            cur.execute("SELECT 'WO',prw_hd.OWNO, ONAME||COALESCE(OFNAME,' ') FROM prw_hd WHERE prw_hd.PR_STATUS = 'P'")
            num_w = cur.fetchall()
            result = num_i + num_q+num_w

            self.tableWidget_his.setRowCount(0)
            for row_number, row_data in enumerate(result):
                self.tableWidget_his.insertRow(row_number)
                for column_number, data in enumerate(row_data):
                    self.tableWidget_his.setItem(row_number, column_number, QTableWidgetItem(str(data)))
            #reprint button
            for index in range(self.tableWidget_his.rowCount()):
                self.btn_reprint = QPushButton('Reprint')                
                self.tableWidget_his.setCellWidget(index,3,self.btn_reprint)
                self.btn_reprint.clicked.connect(self.reprintBut)


    except Exception as e:
        print(e)







