import openpyxl,sys
from openpyxl.styles import Border,Side,Alignment,Font
from openpyxl.utils import column_index_from_string as ci, get_column_letter as gl
p=sys.argv[1]
wb=openpyxl.load_workbook(p); ws=wb['年計']
T=Side(style='thin'); M=Side(style='medium'); N=Side(style=None)
months=['B','E','H','K','N','Q','T','W','Z','AC','AF','AI']
for L in months:
    dc,ec,cc=L,gl(ci(L)+1),gl(ci(L)+2)
    for day in range(1,32):
        r1=4+2*(day-1); r2=r1+1
        top=M if day==1 else T; bot=M if day==31 else T
        for r in (r1,r2):
            t=top if r==r1 else N; b=bot if r==r2 else N
            ws[f'{dc}{r}'].border=Border(left=M,right=T,top=t,bottom=b)
            ws[f'{ec}{r}'].border=Border(left=T,right=N,top=t,bottom=b)
            ws[f'{cc}{r}'].border=Border(left=N,right=M,top=t,bottom=b)
            ws[f'{dc}{r}'].alignment=Alignment(horizontal='center',vertical='center')
            ws[f'{ec}{r}'].alignment=Alignment(horizontal='left',vertical='center',shrink_to_fit=True)
            c=ws[f'{cc}{r}']
            multi=isinstance(c.value,str) and '\n' in c.value
            c.alignment=Alignment(horizontal='right',vertical='center',wrap_text=multi,shrink_to_fit=not multi)
            if multi:
                c.font=Font(name=c.font.name,size=min(c.font.sz or 9,7),color=c.font.color)
# 見出し行（3行目）と罫線の外側
for L in months:
    dc,ec,cc=L,gl(ci(L)+1),gl(ci(L)+2)
    ws[f'{dc}3'].border=Border(left=M,top=M,bottom=M)
    ws[f'{ec}3'].border=Border(top=M,bottom=M)
    ws[f'{cc}3'].border=Border(right=M,top=M,bottom=M)
# 表の外（AL列以降）の4〜65行の罫線を消す
for r in range(4,66):
    for c in range(ci('AL'),ci('AM')+1):
        ws.cell(r,c).border=Border()
for r in range(4,66): ws.row_dimensions[r].height=24
wb.save(p); print('ok')
