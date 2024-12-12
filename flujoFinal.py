import numpy as np
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from tkcalendar import DateEntry
from tqdm import tqdm 
import pandas as pd
from datetime import datetime
import openpyxl as xl
import shutil

pendSunat = None
ap = None
prov = None
oc = None
acumulado = None
leasing = None

def vermon(x):
    return str(x)

def get_file(n):
    global pendSunat, ap, prov, oc, acumulado,leasing  # Declarar como globales
    try:
        name = filedialog.askopenfilename(filetypes=[("All Files","*.xlsx *.xls *.xlsb")]) 
        #if name:    
        if n == 1:
            pendSunat = pd.read_excel(name)
        if n == 2:
            ap = pd.read_excel(name)
        if n == 3:
            prov = pd.read_excel(name)
        if n == 4:
            oc = pd.read_excel(name)
        if n == 5:
            leasing = pd.read_excel(name, sheet_name='Proveedores Lista Autorizada')
        if n == 6:
            #csv_name = name.replace('.xlsx', '.csv').replace('.xlsb', '.csv').replace('.xls', '.csv')
            #acumulado = pd.read_csv(csv_name)
            acumulado = pd.read_excel(name)
        update_label(name, n)
    except Exception as e:
        print(e)
        print("Error al cargar el archivo")

def update_label(file_path, n):
    label = tk.Label(second_frame, text=file_path, bg='#2C6ABF',padx=60,wraplength=200, fg='white', font=('Arial', 9),anchor='w',justify='left')
    if file_path:
        if n==6:
            label.grid(row=8, column=1, padx=20, pady=20)
        else:
            label.grid(row=n-1, column=1, padx=20, pady=20)

def on_closing():
    root.destroy()

def show_process_finished(save_nm):
    msg_label = tk.Label(second_frame, text=f"¡Proceso finalizado!,Puede cerrar la ventana y abir\n {save_nm}", bg='#32a852', fg='white', font=('Arial', 12),anchor='w',justify='left')
    msg_label.grid(row=15, column=1, padx=20, pady=20)
    
def show_buttons():    
    if var.get() == "Si":
        bt_examinar_2.grid(row=8, column=0, padx=20, pady=20)
        label_cal.grid(row=9, column=0, padx=20, pady=20)
        label_cal_ini.grid(sticky = 'w',row=10, column=0, padx=20, pady=20)
        label_cal_fin.grid(sticky = 'w',row=11, column=0, padx=20, pady=20)
        cal_ini.grid(sticky = 'e',row=10, column=0, padx=20, pady=20)
        cal_fin.grid(sticky = 'e',row=11, column=0, padx=20, pady=20)
        
    else:
        bt_examinar_2.grid_forget()
        cal_ini.grid_forget()
        cal_fin.grid_forget()
        label_cal.grid_forget()
        label_cal_ini.grid_forget()
        label_cal_fin.grid_forget()

def realizar_agregacion(df):
    
    res_monto = pd.pivot_table(df,index=['RUC'],columns=['Moneda'],
                                    values=['Monto Doc'],aggfunc='sum')

    res_pendientes = pd.crosstab(index=[df['RUC']],
                                 columns=df['Moneda'])

    resumen_por_ruc = pd.concat([res_pendientes, res_monto], axis=1)
    return resumen_por_ruc

def fusionar_resultados_acumulados(resultados_acum1, resultados_acum2, resultados_acum3):
    # Fusionar los DataFrames por la columna 'RUC'
    fusionados = pd.merge(resultados_acum1, resultados_acum2, on='RUC', how='outer')
    fusionados = pd.merge(fusionados, resultados_acum3, on='RUC', how='outer')

    return fusionados

def proceso_si(final_table):
    global acumulado
    start_date = pd.to_datetime(cal_ini.get_date())
    end_date = pd.to_datetime(cal_fin.get_date())

    print(acumulado.columns)
    acumulado_copy = acumulado.copy()
    # Convertir la columna 'Fecha Doc' a cadena (str) si no lo es ya
    acumulado_copy['Fecha Doc'] = acumulado_copy['Fecha Doc'].astype('str')
    date_mask = acumulado_copy['Fecha Doc'].str.contains('/')
    df_regular = acumulado_copy[date_mask].copy()
    df_excel = acumulado_copy[~date_mask].copy()
    df_regular['Fecha Doc'] = pd.to_datetime(df_regular['Fecha Doc'])

    df_excel['Fecha Doc'] = pd.TimedeltaIndex(df_excel['Fecha Doc'].astype(int), unit='d') + datetime(1899, 12, 30)
    acumulado = pd.concat([df_regular, df_excel])

    print(acumulado['Fecha Doc'].head(10))

    acumulado = acumulado[(acumulado['Fecha Doc'] >= start_date) & (acumulado['Fecha Doc'] <= end_date)]
    acumulado = acumulado[(acumulado['Fecha Doc'] >= start_date) & (acumulado['Fecha Doc'] <= end_date)]
    acumulado = acumulado.loc[acumulado['Estado'] != 'CANCELLED']
    acumulado = acumulado.loc[acumulado['Estado'] != 'CANCELADO']

    #return acum_filter

    mes1 = start_date.month
    mes2 = (start_date.month % 12) + 1
    mes3 = end_date.month

    acum1 = acumulado[acumulado['Fecha Doc'].dt.month == mes1]
    acum2 = acumulado[acumulado['Fecha Doc'].dt.month == mes2]
    acum3 = acumulado[acumulado['Fecha Doc'].dt.month == mes3]

    resultados_acum1 = realizar_agregacion(acum1)
    resultados_acum2 = realizar_agregacion(acum2)
    resultados_acum3 = realizar_agregacion(acum3)

    # Fusionar resultados con acum_filter
    resultados_fusionados = fusionar_resultados_acumulados(resultados_acum1, resultados_acum2, resultados_acum3)

    resultados_fusionados.reset_index(inplace=True)
    resultados_fusionados['Llave'] = ""
    resultados_fusionados['Llave'] = resultados_fusionados.apply(
                lambda x: vermon(x[resultados_fusionados.columns[0]]), axis=1)
    fin = final_table.merge(resultados_fusionados, on='Llave', how='left')
    return fin
    
def proccess_files():
    
    global pendSunat, ap, prov, oc, acumulado, leasing
    if any(df is None for df in [pendSunat, ap, prov, oc, leasing]):
        print("Error: Todos los archivos deben ser cargados.")
        return
    
    #el usuario selecciona donde guardar el archivo
    save_nm = filedialog.asksaveasfilename(filetypes=[("xlsm file", "*.xlsm")],
                                            defaultextension=".xlsm")
    if not save_nm:
        return print('Se cancelo la Generacion')
    print('Guardando en:', save_nm)
    
    """
    save_nm = filedialog.asksaveasfilename(
            filetypes=[("xlsx file", "*.xlsx")], defaultextension=".xlsx")
    if not save_nm:
        return print('Se cancelo la Generacion')"""



    save_nm_copy = save_nm

    #Preparacion de los archivos
    ap= ap.drop(index=range(8))
    ap = ap.reset_index(drop=True)
    nombres_columnas = ap.iloc[0]
    ap = ap[1:]
    ap.columns = nombres_columnas

    prov= prov.drop(index=range(1))
    prov = prov.reset_index(drop=True)
    nombres_columnas = prov.iloc[0]
    prov = prov[1:]
    prov.columns = nombres_columnas

    oc= oc.drop(index=range(6))
    oc = oc.reset_index(drop=True)
    nombres_columnas = oc.iloc[0]
    oc = oc[1:]
    oc.columns = nombres_columnas
    
    leasing=leasing.loc[:, ['RUC', 'PROVEEDOR']]
    #print(leasing.columns)
    #print(leasing)

    #1ER CRUCE  
    #Creamos la KEY para el archivo traido del portal de la SUNAT
    pendSunat['Número  Correlativo de CP'] = pendSunat['Número  Correlativo de CP'].astype(int).apply(lambda x: f'{x:08}')

    pendSunat.insert(2,'Correlativo', pendSunat['Número  documento de identidad del emisor'].astype(str) + pendSunat['Serie de CP'].astype(str) + '-' + pendSunat['Número  Correlativo de CP'].astype(str))

    #Creamos la KEY para el reporte de Sueldos AP
    ap = ap.loc[ap['Estado'] != 'CANCELLED']
    ap = ap.loc[ap['Estado'] != 'CANCELADO']
    ap.insert(2,'Correlativo', ap['RUC'].astype(str) + ap['Nº Doc'].astype(str))

    #1er cruce (SUNAT-Sueldos AP) , descartamos los que ya se tienen registrados
    pendSunat.insert(3,'Registrado en AP', pendSunat['Correlativo'].isin(ap['Correlativo']).map({True: 'Si', False: 'No'}))
    #FIN 1ER CRUCE

    df_crosstab = pd.crosstab(index=pendSunat['Tipo de moneda'], columns=pendSunat['Registrado en AP'])

    # Crear la tabla de resumen de importes
    df_pivot = pd.pivot_table(pendSunat,
                              values='Importe total  FE',
                            index='Tipo de moneda',
                            columns='Registrado en AP',
                            aggfunc='sum',
                            margins=True,
                            margins_name='Total Importe')

    # Concatenar ambas tablas
    resumen_importes = pd.concat([df_crosstab, df_pivot], axis=1, keys=['# Documentos', 'Total Importe'])

    # Redondear los valores
    resumen_importes = resumen_importes.round(2)

    #crear tabla de resumen por RUC de documentos pendientes
    resumen_por_ruc_moneda = pd.pivot_table(pendSunat[pendSunat['Registrado en AP'] == 'No'],
                                    index=['Número  documento de identidad del emisor','Razón social emisor'],columns=['Tipo de moneda'],
                                    values=['Importe total  FE'],aggfunc='sum')

    resumen_por_ruc_pendientes = pd.crosstab(index=[pendSunat[pendSunat['Registrado en AP'] == 'No']['Número  documento de identidad del emisor'],
                                                pendSunat[pendSunat['Registrado en AP'] == 'No']['Razón social emisor']],
                                            columns=pendSunat[pendSunat['Registrado en AP'] == 'No']['Tipo de moneda'])

    resumen_por_ruc = pd.concat([resumen_por_ruc_pendientes, resumen_por_ruc_moneda], axis=1)
    resumen_por_ruc.reset_index(inplace=True)
    resumen_por_ruc['Llave'] = ""
    resumen_por_ruc['Llave'] = resumen_por_ruc.apply(
                lambda x: vermon(x[resumen_por_ruc.columns[0]]), axis=1)
    
    #3er CRUCE 

    resumen_oc_por_ruc = pd.pivot_table(oc,
                                    index=['RUC','Proveedor'],columns=['MON.'],
                                    values=['Pendiente Facturar'],aggfunc='sum')
    resumen_oc_por_ruc.columns = ['_'.join(col) for col in resumen_oc_por_ruc.columns.values]
    sa=resumen_oc_por_ruc
    sa.reset_index(inplace=True)
    sa['Llave'] = ""
    sa['Llave'] = sa.apply(
                lambda x: vermon(x[sa.columns[0]]), axis=1)
    final_table = resumen_por_ruc.merge(sa, on='Llave', how='left')

    if var.get() == "Si":
        # Realizar el proceso específico cuando el radio button está en "Si"
        if acumulado is None:
            print("Error: Debe cargar el archivo acumulado de AP.")
            return
        final_table=proceso_si(final_table)
        col_del = ['Llave', 'RUC_x','RUC_y', 'Proveedor']
        final_table = final_table.drop(col_del, axis=1)

    if var.get() == "No":
        columnas_a_eliminar = ['Llave','RUC','Proveedor']
        final_table = final_table.drop(columnas_a_eliminar, axis=1)

    final_table.rename(columns={('Importe total  FE', 'Sol'): 'Importe FE Sol', ('Importe total  FE', 'US Dollar'): 'Importe FE US Dollar',
                                'Pendiente Facturar_PEN': 'OC Importe PEN','Pendiente Facturar_USD': 'OC Importe USD','Número  documento de identidad del emisor':'RUC'}, inplace=True)
    
    cols = final_table.columns.tolist()
    c1=cols.pop(6)
    c2=cols.pop(6)
    cols.append(c1)
    cols.append(c2)
    final_table = final_table[cols]

    final_table.columns = cols
    final_table['Comentario'] = ''
    final_table['Comentario'] = final_table['RUC'].astype(str).apply(
    lambda x: 'Proveedor de Leasing' if x in leasing['RUC'].astype(str).values
    else ('Proveedor no registrado' if x not in prov['Supplier Number'].astype(str).values else ''))

    if var.get() == "Si":
        start_date = pd.to_datetime(cal_ini.get_date())
        end_date = pd.to_datetime(cal_fin.get_date())
        nombres_meses = {1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril', 5: 'Mayo', 6: 'Junio',
                    7: 'Julio', 8: 'Agosto', 9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'}

        mes1 = start_date.month
        mes2 = (start_date.month % 12) + 1
        mes3 = end_date.month

        # Obtener nombres de mes según mes1, mes2 y mes3
        nombre_mes1 = nombres_meses[mes1]
        nombre_mes2 = nombres_meses[mes2]
        nombre_mes3 = nombres_meses[mes3]

        # Define nuevos nombres para las columnas
        new_column_names = [f'RUC', f'Razón social emisor', f'Sol', f'US Dollar', f'Importe FE Sol',
                        f'Importe FE US Dollar', f'PEN_{nombre_mes1}', f'USD_{nombre_mes1}', f'Monto Doc_PEN_{nombre_mes1}',
                        f'Monto Doc_USD_{nombre_mes1}', f'PEN_{nombre_mes2}', f'USD_{nombre_mes2}', f'Monto Doc_PEN_{nombre_mes2}',
                        f'Monto Doc_USD_{nombre_mes2}', f'PEN_{nombre_mes3}', f'USD_{nombre_mes3}', f'Monto Doc_PEN_{nombre_mes3}',
                        f'Monto Doc_USD_{nombre_mes3}', f'OC Importe PEN', f'OC Importe USD', f'Comentario']


        # Asigna los nuevos nombres a las columnas
        final_table.columns = new_column_names

        nuevo_orden_columnas = [f'RUC', f'Razón social emisor', f'Sol',f'Importe FE Sol', f'US Dollar', 
                        f'Importe FE US Dollar', f'PEN_{nombre_mes1}', f'Monto Doc_PEN_{nombre_mes1}', f'USD_{nombre_mes1}',
                        f'Monto Doc_USD_{nombre_mes1}', f'PEN_{nombre_mes2}', f'Monto Doc_PEN_{nombre_mes2}', f'USD_{nombre_mes2}',
                        f'Monto Doc_USD_{nombre_mes2}', f'PEN_{nombre_mes3}', f'Monto Doc_PEN_{nombre_mes3}', f'USD_{nombre_mes3}',
                        f'Monto Doc_USD_{nombre_mes3}', f'OC Importe PEN', f'OC Importe USD', f'Comentario']

        final_table = final_table[nuevo_orden_columnas]

    if var.get() == "No":
        # Define nuevos nombres para las columnas
        new_column_names = [f'RUC', f'Razón social emisor', f'Sol', f'US Dollar', f'Importe FE Sol',
                        f'Importe FE US Dollar', f'OC Importe PEN', f'OC Importe USD', f'Comentario']

        # Asigna los nuevos nombres a las columnas
        final_table.columns = new_column_names

        nuevo_orden_columnas = [f'RUC', f'Razón social emisor', f'Sol',f'Importe FE Sol', f'US Dollar', 
                        f'Importe FE US Dollar',f'OC Importe PEN', f'OC Importe USD', f'Comentario']

        final_table = final_table[nuevo_orden_columnas]

    # Escribir a Excel
    #with pd.ExcelWriter(save_nm) as writer:
    #    resumen_importes.to_excel(writer, sheet_name='Resumen General', index=True)
    #    final_table.to_excel(writer, sheet_name='DocPendientes',startrow=4, index=True,float_format="%.2f")
    #    pendSunat.to_excel(writer, sheet_name='Comprobantes', index=False)
        
    _file = r"files\Template_.xlsm"
    shutil.copy(_file, save_nm)

    book = xl.load_workbook(save_nm, read_only=False,keep_vba=True)

    print(save_nm)
    with pd.ExcelWriter(save_nm, engine='openpyxl',mode='a') as wrt:
        wrt.workbook = book
        wrt.worksheets = dict((ws.title, ws) for ws in book.worksheets)
        #wrt.vba_archive = book.vba_archive
        resumen_importes.to_excel(wrt, sheet_name='Resumen General', index=True)
        final_table.to_excel(wrt, sheet_name='DocPendientes',startrow=4, index=True,float_format="%.2f")
        pendSunat.to_excel(wrt, sheet_name='Comprobantes', index=False)
        #wrt.save(save_nm)
    book.close()
    show_process_finished(save_nm_copy)


#interfaz grafica
root = tk.Tk()
root.title("Seleccionar archivos")
root.protocol("WM_DELETE_WINDOW", on_closing)
root.geometry("700x600")
title_ = tk.Label(root, text="Seleccionar archivos",bg='#32a852',fg='white',font=('Arial', 14))
title_.pack(fill=tk.X)
#main frame
main_frame = tk.Frame(root, bg='#FFFFFF')
main_frame.pack(fill=tk.BOTH, expand=True)   

#canvas
my_canvas = tk.Canvas(main_frame, bg='#FFFFFF')
my_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

#add scrollbar to canvas
my_scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=my_canvas.yview)
my_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

#configure canvas
my_canvas.configure(yscrollcommand=my_scrollbar.set)

#create another frame inside canvas
second_frame = tk.Frame(my_canvas, bg='#FFFFFF')
second_frame.bind('<Configure>', lambda e: my_canvas.configure(scrollregion = my_canvas.bbox("all")))

#add new frame to a window in the canvas
my_canvas.create_window((0,0), window=second_frame, anchor="nw")

#Buttons
wrap_len = 250
button_width = 25
bt_ps = tk.Button(second_frame, text="Examinar archivo de Pendientes Sunat",width= button_width,wraplength=wrap_len,padx=20, pady=10,
                  command=lambda: get_file(1), bg='#242140', fg='white', font=('Arial', 12))
bt_ps.grid(row=0, column=0, padx=20, pady=20)

bt_ap = tk.Button(second_frame, text="Examinar reporte de saldos AP",width= button_width, wraplength=wrap_len, padx=20, pady=10,
                  command=lambda: get_file(2), bg='#242140', fg='white', font=('Arial', 12))
bt_ap.grid(row=1, column=0, padx=20, pady=20)

bt_prov = tk.Button(second_frame, text="Examinar reporte de Proveedores",width= button_width, wraplength=wrap_len, padx=20, pady=10,
                    command=lambda: get_file(3), bg='#242140', fg='white', font=('Arial', 12))
bt_prov.grid(row=2, column=0, padx=20, pady=20)

bt_oc = tk.Button(second_frame, text="Examinar reporte de OCs Recepcionadas sin facturar",width= button_width, wraplength=wrap_len, padx=20, pady=10,
                  command=lambda: get_file(4), bg='#242140', fg='white', font=('Arial', 12))
bt_oc.grid(row=3, column=0, padx=20, pady=20)

bt_ls = tk.Button(second_frame, text="Examinar reporte de Proveedores Leasing",width= button_width, wraplength=wrap_len, padx=20, pady=10,
                  command=lambda: get_file(5), bg='#242140', fg='white', font=('Arial', 12))
bt_ls.grid(row=4, column=0, padx=20, pady=20)

#RADIO BUTTON ACUMULADO AP
var = tk.StringVar()
radio_si = tk.Radiobutton(second_frame, text="Sí", variable=var, value="Si", font=('Arial', 10), 
                          command=show_buttons,justify="left", anchor="w",bg='#FFFFFF')
radio_no = tk.Radiobutton(second_frame, text="No", variable=var, value="No",font=('Arial', 10), 
                          command=show_buttons,justify="left", anchor="w",bg='#FFFFFF')
label_reporte_acumulado = tk.Label(second_frame, text="Reporte Acumulado de AP", fg='black', font=('Arial', 11),bg='#FFFFFF')
label_reporte_acumulado.grid(sticky = 'w',row=5, column=0, padx=25, pady=20)
radio_si.grid(sticky = 'w', row=6, column=0,padx = 25)
radio_no.grid(sticky = 'w', row=7, column=0,padx = 25)    

bt_examinar_2 = tk.Button(second_frame, text="Examinar Acumulado de AP", padx=10, pady=10,
                          command=lambda: get_file(6), bg='#242140', fg='white', font=('Arial', 12))

label_cal = tk.Label(second_frame, text="Seleccione un rango de fechas de 3 meses",bg='#FFFFFF', fg='black', font=('Arial', 11))

label_cal_ini = tk.Label(second_frame, text="Fecha de inicio:",bg='#FFFFFF', fg='black', font=('Arial', 11))
label_cal_fin= tk.Label(second_frame, text="Fecha de fin:",bg='#FFFFFF', fg='black', font=('Arial', 11))

cal_ini = DateEntry(second_frame, background='darkblue',locale="es", foreground='white', borderwidth=2,font=('Arial',8)) 

cal_fin = DateEntry(second_frame,background='darkblue',locale="es", foreground='white', borderwidth=2,font=('Arial',8)) 

bt_process = tk.Button(second_frame, text="Procesar", padx=20, pady=10,
                       command=proccess_files, bg='#98CCF7', fg='white', font=('Arial', 12))
bt_process.grid(row=15, column=0, padx=20, pady=20)
root.mainloop()