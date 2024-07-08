import threading
from tkinter import *
from tkinter import messagebox
from speedtest import Speedtest
import statistics
from ping3 import ping
from datetime import datetime
import asyncio

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.oauth2 import service_account



#Llamada al excel sheets y uso de la key propocionada por la API
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
KEY = 'key.json'
SPREADSHEET_ID = '1r99GE71SYGZ0L8-U8c5ePdTBh1YJca-bTyVJn_VfMTM'

creds = None
creds = service_account.Credentials.from_service_account_file(KEY,scopes=SCOPES)

service = build('sheets', 'v4', credentials=creds)
sheet = service.spreadsheets()


# Host para el ping de la prueba de paquetes perdidos
host = "google.com"
ping_count = 100
ping_timeout = 0.1

#definir propiedades de la interfaz donde se ejecutará el código
root = Tk()
root.title("Internet Speed Test")
root.geometry("1024x660")
root.resizable(False, False)
root.configure(bg="#1a212d")

#definir variables de la encuesta Qoe
var = IntVar()
var1 = IntVar()
var2 = IntVar()
var3 = IntVar()
sel=0
sel1=0
sel2=0
sel3=0
encuesta=""
download_speed=0 
upload_speed=0
ping_value=0
jitter_value=0
packet_loss=0
e_model_value=0
fecha_de_ejec=0

# Función para medir el ancho de banda
def medir_ancho_de_banda():
    st = Speedtest()
    st.get_best_server()
    download_speed = st.download() / 1_000_000  # Conversión a Mbps
    upload_speed = st.upload() / 1_000_000  # Conversión a Mbps
    results = st.results.dict()
    return download_speed, upload_speed, results

# Función para medir el ping
def medir_ping():
    st = Speedtest()
    st.get_best_server()
    ping_value = st.results.ping
    return ping_value

# Función para calcular el jitter
def calcular_jitter(pings):
    if len(pings) < 2: #si son menos de dos pings se elimina porque no se ejecuta una conexión estable
        return 0
    return statistics.stdev(pings)

# Función para medir la pérdida de paquetes usando ping3 de forma asincrónica
async def async_ping(host, timeout=1):
    return ping(host, timeout=timeout)

async def perdida_paquetes(host, count=50, timeout=1):
    tasks = [async_ping(host, timeout) for _ in range(count)]
    results = await asyncio.gather(*tasks)
    lost_packets = results.count(None)
    packet_loss = (lost_packets / count) * 100
    return packet_loss

# Función para calcular el valor E-Model
def calcular_e_model(ping_value, jitter_value, packet_loss):
    R0 = 93.2
    Is = 0
    d = ping_value + jitter_value 
    # Calcular el factor de deterioro por retardo
    if d <= 177.3:
        Id = 0.024 * d
    else:
        Id = 0.024 * d + 0.11 * (d - 177.3)
    I_e = (95) * packet_loss / (packet_loss + 4.3)
    A = 0
    # Calcular R
    R = round(R0 - Is - Id - I_e + A,2)
    return R

#Función para generar la fecha de ejcución
def fecha_ejec():
    now = datetime.now()
    fecha_hora = now.strftime("%Y-%m-%d %H:%M:%S")
    return fecha_hora



#Funcion para crear el archivo de encuesta
def verificar ():
        global sel
        global sel1
        global sel2
        global sel3
        sel=var.get()  
        sel1=var1.get()
        sel2=var2.get()
        sel3=var3.get()
        if (sel!=0 and sel1!=0 and sel2!=0 and sel3!=0):
            button_enviar.config(state=NORMAL)
            button_enviar.place(x=700, y=580)
            
def enviar():
        global sel
        global sel1
        global sel2
        global sel3
        global var 
        global var1 
        global var2
        global var3
        global encuesta
        seltotal=sel1+sel2+sel3+sel
        if (seltotal<=4):
            encuesta="Servicio deficiente"
        elif (seltotal<=9):
              encuesta="Servicio no cumplidor" 
        elif (seltotal<=11):
              encuesta="Buen servicio" 
        elif (seltotal==12):
              encuesta="Excelente servicio"
        label_logo.place(x=550,y=180)
        sel=0
        sel1=0
        sel2=0
        sel3=0
        values = [[fecha_de_ejec,download_speed,upload_speed,ping_value,jitter_value,packet_loss,e_model_value,encuesta]]

        result=sheet.values().append(spreadsheetId=SPREADSHEET_ID,
                             range='A1',
                             valueInputOption='USER_ENTERED',
                             body={'values':values}).execute()

        var.set(-1)
        var1.set(-1)
        var2.set(-1)
        var3.set(-1)
        pr1.place_forget()
        pr2.place_forget()
        pr3.place_forget()
        pr4.place_forget()
        f1.place_forget()
        f2.place_forget()
        f3.place_forget()
        f4.place_forget()
        button_enviar.place_forget()
        etiqueta_VD.config(text=f"{00:.2f}")
        etiqueta_VS.config(text=f"{00:.2f}")
        etiqueta_ping.config(text=f"{00:.2f}")
        etiqueta_jitter.config(text=f"{00:.2f}")
        etiqueta_jitter.config(text=f"{00:.2f}")
        etiqueta_packet.config(text=f"{00:.2f}")
        etiqueta_e_model.config(text=f"{00:.2f}")
#Función para activar la encuesta qoe
def encuesta_qoe ():  
    #Variables a usar en los botones
    pr1.place(x=500, y=100)
    pr2.place(x=500, y=220)
    pr3.place(x=500, y=340)
    pr4.place(x=500, y=460) 
    #frames para los botones
    f1.place(x=540, y=150)
    f2.place(x=540, y=270)
    f3.place(x=540, y=390)
    f4.place(x=540, y=510)
    # Función para ejecutar todas las pruebas
    R1.pack(anchor=W, side=LEFT, ipadx=40)
    R2.pack(anchor=W, side=LEFT, ipadx=40)
    R3.pack(anchor=W, side=LEFT,ipadx=40)
    R12.pack(anchor=W, side=LEFT, ipadx=40)
    R22.pack(anchor=W, side=LEFT, ipadx=40)
    R32.pack(anchor=W, side=LEFT,ipadx=40)
    R13.pack(anchor=W, side=LEFT, ipadx=40)
    R23.pack(anchor=W, side=LEFT, ipadx=40)
    R33.pack(anchor=W, side=LEFT,ipadx=40)
    R14.pack(anchor=W, side=LEFT, ipadx=40)
    R24.pack(anchor=W, side=LEFT, ipadx=40)
    R34.pack(anchor=W, side=LEFT,ipadx=40)

#Crea un msg box para saber que se esta obteniendo resultados
def mostrar_ventana_carga():
    global ventana_carga
    ventana_carga = Toplevel(root)
    ventana_carga.title("Obteniendo resultados")
    ventana_carga.geometry("350x115")
    ventana_carga.configure(bg="#1a212d")
    
    # Cargar la imagen del icono
    imagen_icono = PhotoImage(file=r"icono.png")
    
    # Etiqueta para mostrar el mensaje
    label = Label(ventana_carga,
        text="Obteniendo resultados, por favor espere...",
        font=("Arial", 13, "bold"),
        bg="#1a212d",
        fg="white"
    )
    label.pack(expand=True)
    
    # Etiqueta para la imagen
    Label(ventana_carga, image=imagen_icono, bg="#1a212d")
    
    ventana_carga.iconphoto(False, imagen_icono)

#cierra el messagebox creado 
def cerrar_ventana_carga():
    ventana_carga.destroy()
    
# Función para ejecutar todas las pruebas
def ejecutar_tests():
    global download_speed, upload_speed, ping_value, jitter_value, packet_loss, e_model_value, fecha_de_ejec
    #muestra el msg box
    mostrar_ventana_carga()
    # Medir ancho de banda
    download_speed, upload_speed, results = medir_ancho_de_banda()
    etiqueta_VD.config(text=f"{download_speed:.2f}")
    etiqueta_VS.config(text=f"{upload_speed:.2f}")

    # Medir ping
    ping_value = medir_ping()
    etiqueta_ping.config(text=f"{ping_value:.2f}")

    # Calcular y mostrar jitter
    num_tests = 10
    pings = [medir_ping() for _ in range(num_tests)]
    jitter_value = calcular_jitter(pings)
    etiqueta_jitter.config(text=f"{jitter_value:.2f}")

    # Medir y mostrar pérdida de paquetes
    packet_loss = asyncio.run(perdida_paquetes(host, ping_count, ping_timeout))
    etiqueta_packet.config(text=f"{packet_loss:.2f}")

    # Calcular y mostrar E-Model
    e_model_value = calcular_e_model(ping_value, jitter_value, packet_loss)
    etiqueta_e_model.config(text=f"{e_model_value:.2f}")

    #Generar la fecha de ejecución
    fecha_de_ejec = fecha_ejec()
    
    cerrar_ventana_carga()
    
    # Cerrar el logo
    label_logo.place_forget() 
    encuesta_qoe()

def ejecutar_tests_en_hilo():
    threading.Thread(target=ejecutar_tests).start()
# Icono
image_icon = PhotoImage(file=r"icono.png")
root.iconphoto(False, image_icon)
label_image_icon=Label(root, image=image_icon, bg="#1a212d") 

# Imágenes
Top = PhotoImage(file=r"topp.png")
labelTop = Label(root, image=Top, bg="#1a212d")
labelTop.place(x=7,y=0)

Main = PhotoImage(file=r"velocimetrogigante2.png")
Label(root, image=Main, bg="#1a212d").place(x=10,y=140)

button = PhotoImage(file=r"botonok.png")
Button(root, image=button, bg="#1a212d", bd=0, activebackground="#1a212d", cursor="hand2", command=ejecutar_tests_en_hilo).place(x=185,y=480)

logo = PhotoImage(file=r"logo.png")
label_logo=Label(root, image=logo, bg="#1a212d")
label_logo.place(x=550,y=180)

# Etiquetas
etiqueta_jitter = Label(root, text="00", font="arial 13 bold", bg="#1a212d", fg="white")
etiqueta_jitter.place(x=82, y=60, anchor="center")

etiqueta_VD = Label(root, text="00", font="arial 13 bold", bg="#1a212d", fg="white")
etiqueta_VD.place(x=190, y=60, anchor="center")

etiqueta_VS = Label(root, text="00", font="arial 13 bold", bg="#1a212d", fg="white")
etiqueta_VS.place(x=300, y=60, anchor="center")

etiqueta_packet = Label(root, text="0.0", font="arial 13 bold", bg="#1a212d", fg="white")
etiqueta_packet.place(x=408, y=60, anchor="center")

etiqueta_ping = Label(root, text="00", font="arial 40 bold", bg="#1a212d", fg="white")
etiqueta_ping.place(x=245, y=330, anchor="center")

etiqueta_e_model = Label(root, text="00", font="arial 25 bold", bg="#1a212d", fg="white")
etiqueta_e_model.place(x=70, y=550, anchor="center")  
# Labels de título
Label(root, text="JITTER", font="arial 10 bold", bg="#1a212d", fg="white").place(x=55, y=0)
Label(root, text="DESCARGA", font="arial 10 bold", bg="#1a212d", fg="white").place(x=150, y=0)
Label(root, text="CARGA", font="arial 10 bold", bg="#1a212d", fg="white").place(x=270, y=0)
Label(root, text="PACKET LOSS", font="arial 10 bold", bg="#1a212d", fg="white").place(x=355, y=0)
Label(root, text="E-MODEL", font="arial 10 bold", bg="#1a212d", fg="white").place(x=40, y=500)  # Ajuste de coordenadas

Label(root, text="MS", font="arial 7 bold", bg="#1a212d", fg="white").place(x=72, y=80)
Label(root, text="Mbps", font="arial 7 bold", bg="#1a212d", fg="white").place(x=175, y=80)
Label(root, text="Mbps", font="arial 7 bold", bg="#1a212d", fg="white").place(x=285, y=80)
Label(root, text="%", font="arial 7 bold", bg="#1a212d", fg="white").place(x=400, y=80)

Label(root, text="PING", font="arial 25 bold", bg="#1a212d", fg="white").place(x=205, y=220)
Label(root, text="MS", font="arial 25 bold", bg="#1a212d", fg="white").place(x=215, y=390)
#labels de preguntas
pr1=Label(root, text="¿Con qué frecuencia tiene pérdidas de su internet?", font="arial 15 ", bg="#1a212d", fg="white")
pr2=Label(root, text="¿Cómo calificaría la transmisión de contenido multimedia?", font="arial 15 ", bg="#1a212d", fg="white")
pr3=Label(root, text="¿Cómo calificaría la calidad llamadas virtuales? ", font="arial 15 ", bg="#1a212d", fg="white")
pr4=Label(root, text="¿Cómo calificaría la fluidez de navegación web?", font="arial 15 ", bg="#1a212d", fg="white")
button_enviar = Button(root, text="Enviar", command=enviar, state=DISABLED)
button_enviar.place_forget()
f1=Frame(root, width=50, bd=10,bg="#1a212d")
f1.grid(column=1,row=2, sticky=W)
f2=Frame(root, width=50, bd=10,bg="#1a212d")
f2.grid(column=1,row=2, sticky=W)
f3=Frame(root, width=50, bd=10,bg="#1a212d")
f3.grid(column=1,row=2, sticky=W)
f4=Frame(root, width=50, bd=10,bg="#1a212d")
f4.grid(column=1,row=2, sticky=W)
pr1.place_forget()
pr2.place_forget()
pr3.place_forget()
pr4.place_forget()
f1.place_forget()
f2.place_forget()
f3.place_forget()
f4.place_forget()
#botones
showfont="Arial 11"
#botones para la primera pregunta
R1 = Radiobutton(f1, anchor=W, text="Nunca", variable=var, value=3, command=verificar,font=showfont, width=3,bg="#1a212d",foreground="white",selectcolor="#1a212d")
R2 = Radiobutton(f1, anchor=W, text="A veces", variable=var, value=2, font=showfont,command=verificar, width=3,bg="#1a212d",foreground="white",selectcolor="#1a212d")
R3 = Radiobutton(f1, anchor=W, text="Muy seguido", variable=var, value=1,font=showfont,command=verificar, width=3,bg="#1a212d",foreground="white",selectcolor="#1a212d")
#botones para la segunda pregunta
R12 = Radiobutton(f2, anchor=W, text="Mala", variable=var1, value=1,font=showfont, width=3,command=verificar,bg="#1a212d",foreground="white",selectcolor="#1a212d")
R22 = Radiobutton(f2, anchor=W, text="Buena", variable=var1, value=2, font=showfont, width=3,command=verificar,bg="#1a212d",foreground="white",selectcolor="#1a212d")
R32 = Radiobutton(f2, anchor=W, text="Excelente", variable=var1, value=3, font=showfont, width=3,command=verificar,bg="#1a212d",foreground="white",selectcolor="#1a212d")
#botones para la tercera pregunta
R13 = Radiobutton(f3, anchor=W, text="Mala", variable=var2, value=1, font=showfont, width=3,command=verificar,bg="#1a212d",foreground="white",selectcolor="#1a212d")
R23 = Radiobutton(f3, anchor=W, text="Buena", variable=var2, value=2, font=showfont, width=3,command=verificar,bg="#1a212d",foreground="white",selectcolor="#1a212d")
R33 = Radiobutton(f3, anchor=W, text="Excelente", variable=var2, value=3, font=showfont, width=3,command=verificar,bg="#1a212d",foreground="white",selectcolor="#1a212d")
#botones para la cuarta pregunta
R14 = Radiobutton(f4, anchor=W, text="Mala", variable=var3, value=1, font=showfont, width=3,command=verificar,bg="#1a212d",foreground="white",selectcolor="#1a212d")
R24 = Radiobutton(f4, anchor=W, text="Buena", variable=var3, value=2, font=showfont, width=3,command=verificar,bg="#1a212d",foreground="white",selectcolor="#1a212d")
R34 = Radiobutton(f4, anchor=W, text="Excelente", variable=var3, value=3, font=showfont, width=3,command=verificar,bg="#1a212d",foreground="white",selectcolor="#1a212d")
#olvidar botones

root.mainloop()