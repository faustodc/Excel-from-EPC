import sys, os, csv, shutil, time
import win32com.client as win32
from os.path import expanduser

today_date = time.strftime("%d/%m/%y")
today_date = today_date[:-2]+"20"+today_date[-2:]
time_hour = time.strftime("%H")
path = os.getcwd().replace('\'','\\') + '\\'
desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
result_folder = os.path.join(desktop, 'Ordini Ricambi')
home = expanduser("~")
excel = win32.gencache.EnsureDispatch('Excel.Application')

def menu():
    print("SCEGLIERE UNA OPZIONE\n")
    while True:
        print("1 - Generare Excel da EPC")
        print("2 - Uscire")
        choice = str(input("\nScelta: "))
        if choice not in ["1","2"]:
            os.system('cls')
            continue
        elif choice == "2":
            sys.exit()
        elif choice == "1":
            os.system('cls')
def getUserData():
    global date_final, am_pm, ol, wip, magazzino, riferimento, targa, operaio, nomepie
    date_final = input("Inserire una data (o premere Enter per inserire " + today_date + "): ")
    if date_final == "":
        date_final = today_date
    if int(time) < 12 and int(time) > 0:
        am_pm = "AM"
    else:
        am_pm = "PM"
    ol = input("Inserire OL TOP CAR: ")
    wip = input("Inserire WIP: ")
    magazzino = input("Inserire magazzino: ")
    magazzino = magazzino.upper()
    riferimento = input("Inserire persona di riferimento: ")
    targa = input("Inserire targa: ")
    operaio = input("Inserire Operaio: ")
    nomepie = input("Inserire un nome di pie di pagina: ")
    date_final = today_date
def writeData():
    global ws
    ws = wb.Worksheets("Pagina 1")
    ws.Range('A20:F41').Value = ""
    ws.Range('F17').Value = ""
    ws.Range('A43').Value = date_final
    ws.Range('B43').Value = am_pm
    ws.Range('H5').Value = ol
    ws.Range('B14').Value = wip
    ws.Range('C9').Value = magazzino
    ws.Range('A17').Value = riferimento
    ws.Range('B17').Value = targa
    ws.Range('C17').Value = operaio
    ws.Range('A52').Value = nomepie

if __name__ == '__main__':
    menu()
    getUserData()
    print(ol)
    print("Un momento...")
    shutil.copyfile('res\Ordine Ricambi - Pagina 1.xlsx', desktop + "\Ordine Ricambi - OL " + str(ol) + " WIP " + str(wip) + "- Pag 1.xlsx")
    wb = excel.Workbooks.Open(desktop + "\Ordine Ricambi - OL " + str(ol) + " WIP " + str(wip) + "- Pag 1.xlsx")
    epc_file = home+'\XFER\\'+str(os.listdir(home+'\XFER')[0])
    if os.path.exists(epc_file):
        page_count = 1
        writeData()
        
        with open(epc_file, 'rt') as csvfile:
            reader = csv.reader(csvfile, delimiter='|')
            count = 0
            reader = iter(reader)
            next(reader)
            for row in reader:
                cells_des = str('C'+str(20+count))
                cells_cat = str('A'+str(20+count))
                cells_qty = str('F'+str(20+count))
                
                if row[0][0] in ['Q']:
                    data_des = str('{} {: <1s} {: <1s} {: <1s} {: <1s}{: <1s}'.format(row[0][0], row[0][1:8], row[0][8:12], row[0][12:16], row[0][16:18], row[0][18:]))
                elif row[0][0] in ['N']:
                    data_des = str('{} {: <1s} {: <1s} {: <1s}'.format(row[0][0], row[0][1:7], row[0][7:14], row[0][14:]))
                else:
                    data_des = str('{} {: <1s} {: <1s} {: <1s} {: <1s} {: <1s}'.format(row[0][0], row[0][1:4], row[0][4:7], row[0][7:9], row[0][9:11], row[0][11:]))
                data_cat = str(row[1])
                data_qty = str(row[2])
                
                ws.Range(cells_des).Value = data_des
                ws.Range(cells_cat).Value = data_cat
                ws.Range(cells_qty).Value = data_qty
                ws.Range('F17').Value = row[3]
                count = count + 1
                if count >= 22:
                    
                    shutil.copyfile('res\Ordine Ricambi - Pagina 1.xlsx', desktop + "\Ordine Ricambi - OL " + str(ol) + " WIP " + str(wip) + "- Pag " + str(page_count+1) + ".xlsx")
                    wb = excel.Workbooks.Open(desktop + "\Ordine Ricambi - OL " + str(ol) + " WIP " + str(wip) + "- Pag " + str(page_count+1) + ".xlsx")
                    writeData()
                    ws.Name = 'Pagina ' + str(page_count+1)                    
                    page_count = page_count + 1
                    count = 0
                    
            print("Premere 'Si' su tutti gli avisi.")
            wb.SaveAs(desktop + "\Ordine Ricambi - OL " + str(ol) + " WIP " + str(wip) + "- Pag " + str(page_count) + ".xlsx")
            excel.Application.Quit()
    else:
        print("Non si ha trovato una lista di spesa!")