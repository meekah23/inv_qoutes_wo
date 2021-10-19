import sys
import os
import pandas as pd
import sqlite3
import time

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



def getData(self):
        try:
            start_time = time.time()
            global table, headers,result,right,amt,gst,total,date, saleman,addedby,wonum,datereq,quote, name, sadd1, sadd2, sadd3, scode, sphone, spstno, lstreet, lname, ladd1, ladd2, ladd3, lcode
            conn = sqlite3.connect(':memory:')
            cur = conn.cursor()

            # load excel file
            df = pd.read_excel("Data1/prw_hd.xlsx", sheet_name="prw_hd")
            
            df.to_sql(name='prw_hd', con=conn, if_exists='append')
            df1 = pd.read_excel("Data1/prw_det.xlsx", sheet_name="prw_det")
            df1.to_sql(name='prw_det', con=conn, if_exists='append')

            #get headers
            cur.execute("SELECT OWDATE,OSLSMAN,ADDEDBY,OWNO,ORDATE FROM prw_hd WHERE PR_STATUS is null")
            result = cur.fetchall()
            date = [a[0] for a in result] 
            saleman = [a[1] for a in result]
            addedby = [a[2] for a in result]
            wonum = [a[3] for a in result]
            datereq = [a[4] for a in result] 
            
            
            
            #get ship details
            cur.execute("SELECT ONAME || ', ' || OFNAME,OADDR1, OADDR2,OADDR3,OPCODE,OBPHONE,OPSTNO,OSTREET,OINAME,OIADDR1,OIADDR2,OIADDR3,OIPCODE FROM prw_hd WHERE  PR_STATUS is null")
            shipping = cur.fetchall()
            name = [a[0] for a in shipping]
            sadd1 = [a[1] for a in shipping]
            sadd2 = [a[2] for a in shipping] 
            sadd3 = [a[3] for a in shipping] 
            scode= [a[4] for a in shipping]
            sphone = [a[5] for a in shipping] 
            spstno = [a[6] for a in shipping] 
            lstreet = [a[7] for a in shipping] 
            lname = [a[8] for a in shipping] 
            ladd1 = [a[9] for a in shipping] 
            ladd2 = [a[10] for a in shipping] 
            ladd3 = [a[11] for a in shipping] 
            lcode = [a[12] for a in shipping] 
                        
            #get inv details
            cur.execute("SELECT OD_ITEM, "
                        "CASE WHEN round(OD_QTY,2) = 0 AND round(OD_PRICE,2) = 0 AND round(OD_AMOUNT,2) = 0 THEN ''"
                        "ELSE round(OD_QTY)"
                        "END, "
                        "COALESCE(OD_DESCR, ' ') as descr,"
                        "CASE WHEN round(OD_QTY,2) = 0 AND round(OD_PRICE,2) = 0 AND round(OD_AMOUNT,2) = 0 THEN '' "
                        "ELSE round(OD_PRICE,2)"
                        "END, "
                        "CASE WHEN round(OD_QTY,2) = 0 AND round(OD_PRICE,2) = 0 AND round(OD_AMOUNT,2) = 0 THEN '' "
                        "ELSE round(OD_AMOUNT,2) "
                        "END, "
                        "prw_hd.OWNO "
                        #"sum(round(OD_AMOUNT,2)), sum(round(prw_hd.OGST,2)), sum(round(OD_AMOUNT,2))+ sum(round(prw_hd.OGST,2))"
                       "FROM prw_det "
                        "INNER JOIN prw_hd ON prw_det.OD_UNO = prw_hd.ONUMBER "
                        "WHERE prw_hd.PR_STATUS is null")
            table = cur.fetchall()
            b = [el[5] for el in table]
            newlist = list(dict.fromkeys(b))            
            finalist=[]

            #get total
            for c in newlist:                
                cur.execute("SELECT SUM(round(OD_AMOUNT,2)),OGST,SUM(round(OD_AMOUNT,2))+OGST FROM prw_det INNER JOIN prw_hd ON prw_det.OD_UNO = prw_hd.ONUMBER WHERE OWNO = ?",(c,))
                d = cur.fetchall()               
                finalist.append(d)
            
            amt = [a[0][0] for a in finalist] 
            gst=[a[0][1] for a in finalist]
            total = [a[0][2] for a in finalist]

            self.tableWidget_wo.setRowCount(0)
            for row_number, row_data in enumerate(table):
                self.tableWidget_wo.insertRow(row_number)
                for column_number, data in enumerate(row_data):
                    self.tableWidget_wo.setItem(row_number, column_number, QTableWidgetItem(str(data)))
            
            if not table:
                QMessageBox.information(self, "Done!", "No more invoice to print.")
            else:
                end_time = time.time()
                execution_time = end_time - start_time
                self.time_label.setText("Data1 fetched for " + str(execution_time) + " secs")

        except Exception as e:
                print(e)
                QMessageBox.information(self,"Error", "Failed to run script.")

def toPDF(self):
        start_time = time.time()
        directory = "inv_pdf"
        all_data = [[x for x in g] for x, g in groupby(table, key = lambda x: x[5])]
        
        try:
            os.makedirs(directory, exist_ok = True)
            print("Directory '%s' created successfully" % directory)
            #selecting simultaneous lists using zip()
            for data, sumamt, ogst, ototal, odate, osale, oadd, owonum, ordate, oname, osadd1, osadd2, osadd3, oscode, osphone, ospstno, olstr, olname, oladd1, oladd2, oladd3, olcode in zip(all_data, amt, gst, total, date, saleman,addedby,wonum,datereq,name, sadd1, sadd2, sadd3, scode, sphone, spstno, lstreet, lname, ladd1, ladd2, ladd3, lcode):

                style2 = ParagraphStyle(
                    name='Normal',
                    fontName='ArialBd',
                    fontSize=22,
                    alignment=1,
                    spaceAfter=15,
                )
                style3 = ParagraphStyle(
                    name='Normal',
                    fontName='Arial',
                    fontSize=12,
                    alignment=1,
                    spaceBefore=20,

                )

                addstyle = ParagraphStyle(
                    name='Normal',
                    fontName='ArialBd',
                    fontSize=10,
                    alignment=1,
                                      
                )

                tablestyle1 = TableStyle([
                    ('FONTSIZE', (0, 0), (-1, -1), 10),
                    ('FONTNAME', (0, 0), (-1, -1), 'ArialBd'),
                    #('BOX', (0, 0), (-1, -1), 0.5, colors.black),

                 ])
                style4 = TableStyle([
                    ('FONTSIZE', (0, 0), (-1, -1), 12),
                    ('FONTNAME', (0, 0), (-1, -1), 'Arial'),
                    ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
                    ('ALIGN',(0,0),(-1,-1), 'CENTER'),
                 ])

                tablestyle2 = TableStyle([

                     ('FONTSIZE', (0, 0), (-1, -1), 10),
                     ('FONTNAME', (0, 0), (-1, -1), 'ArialBd'),
                     ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
                     ('LINEBEFORE', (1, 0), (1, -1), 0.5, colors.black)

                 ])

                headerstyle = TableStyle([

                     ('FONTSIZE', (0, 0), (-1, 0), 9),
                     ('FONTNAME', (0, 0), (-1, 0), 'ArialBd'),
                     ('LINEABOVE', (0, 0), (-1, -1), 1, colors.black),
                     ('LINEBEFORE', (0, 0), (0, -1), 1, colors.black),
                     ('LINEAFTER', (4, 0), (4, -1), 1, colors.black),
                     ])

                footerstyle = TableStyle([

                     ('FONTSIZE', (0, 0), (-1, -1), 8),
                     ('FONTSIZE', (1, 0), (2, -1), 9),
                     ('FONTNAME', (0, 0), (-1, -1), 'ArialBd'),
                     ('LINEBEFORE', (0, 0), (0, 4), 1, colors.black),
                     ('LINEBEFORE', (2, 0), (-1, -1), 1, colors.black),
                     ('LINEAFTER', (2, 0), (-1, -1), 1, colors.black),
                     ('LINEABOVE', (0, 5), (-1, 5), 1, colors.black),
                     ('LINEBELOW', (2, 5), (-1, 5), 1, colors.black),
                     ('ALIGN', (2, 0), (2, -1), "RIGHT"),
                      ])

                def header(canvas, pdf):
                   
                    # Draw heading
                     heading = Paragraph("VANCOUVER GLASS (1990) LTD.", style2)
                     heading.wrap(pdf.width, inch * 0.3)
                     heading.drawOn(canvas, pdf.leftMargin, pdf.height + inch)

                     
                     subheading = Paragraph("WORK ORDER", style3)
                     subheading.wrap(pdf.width, inch * 0.2)
                     subheading.drawOn(canvas, pdf.leftMargin, pdf.height + inch * 0.5)

                     addressnote = Paragraph("1706 E. HASTINGS, VAN, B.C. V5L 1S9 Phone (604)253-7707 Fax (604)253-8448",addstyle)
                     addressnote.wrap(pdf.width, inch)
                     addressnote.drawOn(canvas, pdf.leftMargin, 0.7 * inch)
                     
                    
                        #table line
                     canvas.line(1.7 * inch, 2.5 *inch, 1.7 *inch, 7.6 *inch)
                     canvas.line(2.26 * inch, 2.5 * inch, 2.26 * inch, 7.6 * inch)
                     canvas.line(6.41 * inch, 2.5 * inch, 6.41 * inch, 7.6 * inch)
                     canvas.line(7.19 * inch, 2.5 * inch, 7.19 * inch, 7.6 * inch)

                     

                     tablelist1 = [["Date: "+str(odate),"W/O#: "+str(owonum)],
                                  ["Salesman: "+str(osale), "Date Req: "+str(ordate)],
                                  ["GST#: 121989834RT"],
                                  ["Added-by: "+str(oadd),"Page: "+"%d " % doc.page]
                                  ] 
               
                     tablelist2 = [["Sold To:" , "Location: "+str(olstr)],
                                  [str(oname), "Ship To:"],
                                  [str(osadd1), str(olname)],
                                  [str(osadd2), str(oladd1)],
                                  [str(osadd3), str(oladd2)],
                                  [str(oscode), str(oladd3)],
                                  ["Phone: "+str(osphone), str(olcode)],
                                  ["PST Exempt# "+str(ospstno)]]

                     headerlist = [["Code", "Qty", "Description", "Unit Price", "Amount"]]

                     
                     footerlist = [["", "Sub-Total:", sumamt],
                                 ["", "GST", ogst],
                                 [""],
                                 [""],
                                 [""],
                                 ["", "Total Amount:", round(ototal,2)]]

                     table1 = Table(tablelist1, colWidths=[400,200], rowHeights=[10, 10, 10, 10],
                                   hAlign='CENTER', spaceBefore=5, style=tablestyle1)
                     table1.wrap(pdf.width, inch)
                     table1.drawOn(canvas, pdf.leftMargin, pdf.height - inch * 0.4)  
                        
                     table2 = Table(tablelist2, colWidths=[268, 268], rowHeights=[20, 15, 10, 10, 10, 10, 10, 20],
                                   hAlign='CENTER', spaceBefore=5, style=tablestyle2)
                     table2.wrap(pdf.width, inch)
                     table2.drawOn(canvas, pdf.leftMargin, pdf.height - inch * 2)

                     header = Table(headerlist,hAlign='LEFT', spaceBefore=5, repeatRows=1, style=headerstyle,
                                     colWidths=[86, 43, 301, 54, 53])
                     header.wrap(pdf.width, inch)
                     header.drawOn(canvas, pdf.leftMargin, pdf.height - inch * 2.35)
                     footertable = Table(footerlist, colWidths=[410, 71, 56], rowHeights=[20, 10, 10, 10, 10, 20],
                                        hAlign='LEFT', spaceBefore=5, style=footerstyle)
                     footertable.wrap(pdf.width, inch)
                     footertable.drawOn(canvas, pdf.leftMargin, 1.38 * inch)


                     
                        

                save_name = os.path.join("inv_pdf/", "w"+ str(owonum)+'.pdf')           
                doc = BaseDocTemplate(save_name, leftMargin=0.5 * inch, rightMargin=0.5 * inch)
                
                
                frame = Frame(

                    0.5 * inch,  # x
                    2.5 * inch,  # y at bottom
                    7.46 * inch,  # width
                    4.83 * inch,  # height
                    showBoundary=1
                )
              
                template = PageTemplate(id='all_pages',frames=frame,onPage=header)                             
                doc.addPageTemplates([template])

                tablestyle3 = TableStyle([
                    ('FONTSIZE', (0, 0), (-1, -1), 8),
                    ('FONTNAME', (0, 1), (-1, -1), 'Arial'),
                    ('ALIGN',(1,0),(1,-1), "RIGHT"),
                    ('ALIGN', (3, 0), (5, -1), "RIGHT"),
                ])
                table_style = TableStyle([
                    #('BACKGROUND', (1,1), (-2,-2), colors.green),
                    #('TEXTCOLOR', (0,0), (1,-1), colors.red),
                    ('BOX', (0,0), (-1,-1), 0.45, colors.black),
                    ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.blue),
                ])
                Elements = []
                table3 = Table(data, style=tablestyle3, hAlign='LEFT',repeatRows=1,colWidths=[80, 40, 300, 55, 50, 100], rowHeights=10)
                Elements.append(table3)
                
                doc.build(Elements)
              
            try:               
               df = pd.read_excel("Data1/prw_hd.xlsx", sheet_name="prw_hd")               
               df.loc[(df.PR_STATUS.isnull()), 'PR_STATUS'] = 'P'
               df.to_excel('Data1/prw_hd.xlsx',sheet_name = 'prw_hd', index=False)
               print('Status updated')                
            except Exception as e:
                print(e)
             

            end_time = time.time()
            execution_time = end_time - start_time
            self.time_label.setText("PDF created for " + str(execution_time) + " secs")
        except Exception as e:
            print(e)
            QMessageBox.information(self, "Error", "Failed to run script.")
        

