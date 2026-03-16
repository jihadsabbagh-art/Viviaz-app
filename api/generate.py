from http.server import BaseHTTPRequestHandler
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import json, io

CO = {'addr':'Zollergasse 9/26, 1070, Vienna - Austria','email':'viviane.zahra@zahra-interiors.com','phone':'+ 43 (0) 67682337441','vat':'ATU77022918','eori':'ATEOS1000136890','bank':'Raiffeisen Bank /  Vienna - Austria','iban':'AT47 3200 0000 1348 4274','bic':'RLNWATWW','owner':'Viviane Zahra','name':'Zahra Interiors'}
C0 = {'d':'404040','m':'606060','l':'999999','v':'CCCCCC','bg':'F7F7F7','w':'FFFFFF','h':'4A4A4A'}
TN = Side(style='thin', color=C0['v'])
MD = Side(style='medium', color=C0['d'])
BT = Border(top=TN, bottom=TN, left=TN, right=TN)
BH = Border(top=MD, bottom=MD, left=TN, right=TN)
BM = Border(top=MD)

def F(sz=9,b=False,c='d',n='Calibri Light'):
    return Font(name=n,size=sz,bold=b,color=C0.get(c,c))

def gen(data):
    wb=Workbook();ws=wb.active
    dt=data.get('docType','Invoice');ws.title=dt
    cur=data.get('currency','€')
    cf='#,##0.00 "€"' if cur=='€' else '"$" #,##0.00'
    ws.page_setup.paperSize=ws.PAPERSIZE_A4;ws.page_setup.orientation='portrait'
    ws.page_setup.fitToWidth=1;ws.page_setup.fitToHeight=1
    ws.sheet_properties.pageSetUpPr.fitToPage=True
    ws.page_margins.left=0.55;ws.page_margins.right=0.55;ws.page_margins.top=0.4;ws.page_margins.bottom=0.4
    for i,w in enumerate([15,36,4,4,8,15,5,15]):
        ws.column_dimensions[get_column_letter(i+1)].width=w

    def W(row,col,val,font=None,fill=None,align=None,border=None,nf=None,mg=None):
        c=ws.cell(row=row,column=col,value=val)
        if font:c.font=font
        if fill:c.fill=fill
        if align:c.alignment=align
        if border:c.border=border
        if nf:c.number_format=nf
        if mg:ws.merge_cells(start_row=row,start_column=col,end_row=row,end_column=mg)

    r=2
    W(r,1,dt,F(18,True));ws.row_dimensions[r].height=28
    r=3
    for c in range(1,9):W(r,c,None,border=BM)
    ws.row_dimensions[r].height=4

    r=4
    info=[]
    info.append((dt+' Nr.' if dt=='Invoice' else 'Quotation #',data.get('docNumber','')))
    info.append(('Date',data.get('date','')))
    info.append(('VAT Nr.',CO['vat']))
    if data.get('clientVat'):info.append(('Client VAT Nr.',data['clientVat']))
    info.append(('EORI #',CO['eori']))
    if data.get('workStart'):info.append(('Work Starting Date',data['workStart']))
    if data.get('workEnd'):info.append(('Work Compl. Date',data['workEnd']))
    for i,(l,v) in enumerate(info):
        W(r+i,6,l,F(8,False,'m'),align=Alignment(horizontal='right'))
        W(r+i,8,v,F(8,False,'d'))
        ws.row_dimensions[r+i].height=13

    W(r,1,'Client:',F(9,False,'m'))
    W(r+1,1,data.get('clientName',''),F(11,True,'d','Arial'),mg=4)
    ws.row_dimensions[r+1].height=18
    addr=[l for l in(data.get('clientAddress','')or'').replace('\\n','\n').split('\n')if l.strip()]
    for i,line in enumerate(addr):
        W(r+2+i,1,line,F(9,False,'l'),mg=4)
        ws.row_dimensions[r+2+i].height=13

    r=max(r+2+len(addr),r+len(info))+1
    W(r,1,'Project:',F(10,False,'m'))
    r+=1
    W(r,1,data.get('projectName',''),F(11,True,'d'),mg=5)
    ws.row_dimensions[r].height=18
    if data.get('location'):
        r+=1;W(r,1,'Location: '+data['location'],F(9,False,'m'),mg=8)
    r+=2

    items=data.get('items',[]);scope=[l for l in(data.get('scopeLines')or[])if isinstance(l,str)]

    if items and len(items)>0:
        hf=F(9,True,'w');hfi=PatternFill('solid',fgColor=C0['h'])
        ha=Alignment(horizontal='center',vertical='center')
        for c,h in enumerate(['Item Number','Item Description','','','Qty','Unit Price','','Total'],1):
            al=ha if c>=5 else Alignment(horizontal='left',vertical='center')
            W(r,c,h,hf,hfi,al,BH)
        ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=4)
        ws.row_dimensions[r].height=22;r+=1
        af=PatternFill('solid',fgColor=C0['bg']);wf=PatternFill('solid',fgColor=C0['w'])
        ra=Alignment(horizontal='right');ca=Alignment(horizontal='center')
        for idx,item in enumerate(items):
            fl=af if idx%2==1 else wf
            W(r,1,item.get('itemNumber',''),F(9,False,'d'),fl,border=BT)
            W(r,2,item.get('description',''),F(9,False,'d'),fl,border=BT,mg=4)
            for cc in range(3,5):W(r,cc,None,fill=fl,border=BT)
            W(r,5,item.get('qty',0),F(9,False,'d'),fl,ca,BT)
            W(r,6,item.get('price',0),F(9,False,'d'),fl,ra,BT,cf)
            W(r,7,None,fill=fl,border=BT)
            tv=(item.get('qty',0)or 0)*(item.get('price',0)or 0)
            W(r,8,tv,F(9,False,'d'),fl,ra,BT,cf)
            ws.row_dimensions[r].height=18;r+=1
        r+=1
    elif scope and len(scope)>0:
        W(r,1,'Scope Of Work',F(10,True,'d'),mg=8);r+=1
        for c in range(1,9):W(r,c,None,border=Border(top=TN))
        ws.row_dimensions[r].height=4;r+=1
        for line in scope:
            ip=line.strip().lower().startswith('phase')or line.strip().startswith('##')
            cl=line.replace('##','').strip()
            if ip:r+=1;W(r,1,cl,F(10,True,'d'),mg=8)
            elif cl:W(r,1,cl,F(9,False,'m'),mg=8)
            ws.row_dimensions[r].height=15;r+=1
        r+=1

    te=data.get('totalExclVat',0)or 0;vp=data.get('vatPercent',0)or 0
    va=te*vp/100;gr=te+va
    for c in range(5,9):W(r,c,None,border=BM)
    ws.row_dimensions[r].height=4;r+=1
    W(r,5,'Total excl. VAT',F(10,False,'d'),mg=7)
    W(r,8,te,F(10,True,'d'),align=Alignment(horizontal='right'),nf=cf)
    ws.row_dimensions[r].height=20;r+=1
    vl=f'VAT ({vp}%)' if vp>0 else 'VAT reverse charge'
    W(r,5,vl,F(9,False,'m'),mg=7)
    W(r,8,va,F(9,False,'m'),align=Alignment(horizontal='right'),nf=cf)
    ws.row_dimensions[r].height=18;r+=1
    gf=F(11,True,'d')
    for c in range(5,9):W(r,c,None,border=Border(top=MD))
    W(r,5,'Total Amount',gf,align=Alignment(vertical='center'),mg=7,border=Border(top=MD,bottom=MD))
    W(r,8,gr,gf,align=Alignment(horizontal='right',vertical='center'),nf=cf,border=Border(top=MD,bottom=MD))
    ws.row_dimensions[r].height=24;r+=2

    pt=data.get('paymentTerms',[])
    if pt and len(pt)>0:
        W(r,1,'Payments As Follows:',F(10,False,'d'),mg=5);r+=1
        for t in pt:
            if isinstance(t,str)and t.strip():
                W(r,1,'  '+t.strip(),F(9,False,'m'),mg=5)
                ws.row_dimensions[r].height=14;r+=1
        r+=1

    r+=1
    sb=Border(bottom=Side(style='thin',color=C0['d']))
    for c in range(1,4):W(r,c,None,border=sb)
    for c in range(5,9):W(r,c,None,border=sb)
    ws.row_dimensions[r].height=28;r+=1
    W(r,1,CO['owner'],F(9,False,'d'),mg=3)
    W(r,5,'Client Date & Signature',F(9,False,'l'),mg=8)
    r+=1;W(r,1,CO['name'],F(8,False,'l'),mg=3)

    fs=max(r+3,52)
    for c in range(1,9):W(fs,c,None,border=Border(top=Side(style='thin',color=C0['v'])))
    ff=F(7.5,False,'l');fa=Alignment(horizontal='center',vertical='center')
    fl=[CO['addr'],f"e-mail: {CO['email']}        Tel. {CO['phone']}",f"VAT# {CO['vat']}                EORI# {CO['eori']}",f"{CO['bank']}      IBAN: {CO['iban']}      BIC: {CO['bic']}"]
    for i,line in enumerate(fl):
        row=fs+1+i;W(row,1,line,ff,align=fa,mg=8);ws.row_dimensions[row].height=11
    ws.print_area=f'A1:H{fs+len(fl)+1}'

    buf=io.BytesIO();wb.save(buf);buf.seek(0)
    return buf.getvalue()

class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        body=self.rfile.read(int(self.headers.get('Content-Length',0)))
        data=json.loads(body)
        xls=gen(data)
        fn=data.get('docNumber','doc')+'.xlsx'
        self.send_response(200)
        self.send_header('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        self.send_header('Content-Disposition',f'attachment; filename="{fn}"')
        self.send_header('Access-Control-Allow-Origin','*')
        self.send_header('Access-Control-Allow-Methods','POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers','Content-Type')
        self.end_headers()
        self.wfile.write(xls)
    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin','*')
        self.send_header('Access-Control-Allow-Methods','POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers','Content-Type')
        self.end_headers()
