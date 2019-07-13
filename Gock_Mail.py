# import GUI
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox

# imports para enviar correos y adjuntar archivos 
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

# Variables Globales
adju = ""
ruta_arch = []
nombre_arch = []
num_arch = 0

# Prepara el correo para ser envíado
def enviarMail(use,passwd,des,asu,mes):
    global num_arch, ruta_arch , nombre_arch
    # Para cada ruta de los archivos adjuntados guardara su nombre en una lista
    for path in ruta_arch:
        #print("Preparando el envío...")
        ruta , nombre_adjunto = os.path.split(path)
        nombre_arch.append(nombre_adjunto)
        
    # Creamos el objeto mensaje
    mensaje = MIMEMultipart()
    # Establecemos los atributos del mensaje
    mensaje['From'] = use
    mensaje['To'] = ", ".join(des)
    mensaje['Subject'] = asu
    # Agregamos el cuerpo del mensaje como objeto MIME de tipo texto
    mensaje.attach(MIMEText(mes, 'plain'))

    # Para cada archivo lo agregamos al mesnaje del correo
    i = 0
    for path in ruta_arch:
        # Abrimos el archivo que vamos a adjuntar
        try:
            archivo_adjunto = open(path, 'rb')
            # Creamos un objeto MIME base
            adjunto_MIME = MIMEBase('application', 'octet-stream')
            # Y le cargamos el archivo adjunto
            adjunto_MIME.set_payload((archivo_adjunto).read())
            # Codificamos el objeto en BASE64
            encoders.encode_base64(adjunto_MIME)
            # Agregamos una cabecera al objeto
            adjunto_MIME.add_header('Content-Disposition', "attachment; filename= %s" % nombre_arch[i])
            i += 1
            # Y finalmente lo agregamos al mensaje
            mensaje.attach(adjunto_MIME)
        except:
            pass
            #print("Archivo no Adjunto")

    smtpserver = smtplib.SMTP("smtp.gmail.com", 587) # 587 es el puerto estandar para smtp
    smtpserver.ehlo()
    smtpserver.starttls()
    # Tratamos la Excepcion por si no se llega a logear
    try: 
        smtpserver.login(use, passwd)
    except:
        JoptioPane = messagebox.showerror("Error de logeo","Usuario o contraseña incorrecta." )
        return
    # Mensaje para mostrar si se envio correctamente el correo
    header = "From: " + use + "\n"
    header += "To: " + des + "\n"
    header += "Subject: " + asu + "\n"
    header += "Files: " + str(num_arch) + "\n"
    msg = mensaje.as_string()

    # Tratamos la Excepcion por si no se envia el correo
    try:
        smtpserver.sendmail(use, des, msg)
        JoptioPane = messagebox.showinfo("Mensaje Enviado",header)
    except:
        JoptioPane = messagebox.showerror("Mensaje NO Enviado","Fallo al enviar\n" +header)
    smtpserver.close() # Cerramos conexion con el servidor
    # Restauramos todo la GUI y borramos los datos de los archivos adjuntos 
    num_arch = 0
    con.configure(text="Archivos: " + str(num_arch))
    entrada1.delete(0,END)
    entrada2.delete(0,END)
    entrada3.delete(0,END)
    entrada4.delete(0,END)
    t.delete("1.0", "end-1c")
    ruta_arch.clear()
    nombre_arch.clear()

# Funcion que guarda la ruta de los archivos a Adjuntar
def adjuntarAr():
    global adju, num_arch, ruta_arch
    # Guarda la ruta de un archivo del File Choose de tkinter
    # Tembien lo configuramos esteticamente
    adju =  filedialog.askopenfilename(initialdir = "/",title = "Elige tu archivo",filetypes = (("all files","*.*"),("PDF files","*.pdf"),("Word files","*.doc;*.docx"),("PowerPoint files","*.pptx;*.potx;*.ppsx;*.pptm;*.ppsom;*.ppam"),("Excel files","*.xltx;*.xlsx;*.xlsm;*.xltm;*.csv"),("TXT files","*.txt"),("HTML files","*.html"),("R files","*.R"),("Python files","*.py"),("Java files","*.java"),("C files","*.c;*.cpp;*.h"),("Png files","*.png"),("Jpeg files","*.jpg"),("Zip files","*.zip"),("Rar files","*.rar")))
    ruta_arch.append(adju) # Guardamos la ruta en una lista
    num_arch+=1 # Llevamos conteo de archivos adjuntados
    con.configure(text="Archivos: " + str(num_arch)) # Mostramos cantidad de archivos en la GUI

# Funcion que muestra mis datos al presionar "Ctrl + m"
def yo(event):
    messagebox.showwarning("Gock","Mario Vásquez\nEstudiante de Ing. en Computación UNAM\nBecario Gen.37 PROTECO\nAnd the moon is shinning...")

    
###################### FUNCION QUE CENTRA LA VENTANA #######################
def center(win):
    win.update_idletasks()
    width = win.winfo_width()
    height = win.winfo_height()
    x = (win.winfo_screenwidth() // 2) - (width // 2)
    y = (win.winfo_screenheight() // 2) - (height // 2)
    win.geometry('{}x{}+{}+{}'.format(width, height, x, y))

################################## GUI #####################################
Ventana = Tk()
Ventana.title("Gock's Mail")
Ventana.geometry("490x310")
Ventana.resizable(width=False,height=False)
center(Ventana)
titulo = Label(Ventana,text="Bienvenido",font=("Arial",24)).place(x=40,y=15)
instruc = Label(Ventana,text="Favor de llenar los datos correspondientes.",font=("Arial",14)).place(x=40,y=55)
consi = Label(Ventana,text="Usuario: ").place(x=40,y=95)
entrada1 = Entry(Ventana) #Esta posicionado por separado para que funcione el .get() o cualquier otro metodo
#entrada1.insert(END,"algun correo") # Por si se quiere tener un correo por defecto
entrada1.place(x=100,y=95) #Esta posicionado por separado para que funcione el .get() o cualquier otro metodo
consi1 = Label(Ventana,text="Password: ").place(x=40,y=125)
entrada2 = Entry(Ventana,width=20,show='*')
entrada2.place(x=100,y=125)
consi2 = Label(Ventana,text="Destinatario: ").place(x=250,y=95)
entrada3 = Entry(Ventana,width=20)
entrada3.place(x=330,y=95)
consi2 = Label(Ventana,text="Asunto: ").place(x=250,y=125)
entrada4 = Entry(Ventana,width=20)
entrada4.place(x=330,y=125)
consi3 = Label(Ventana,text="Mensaje: ").place(x=40,y=155)
t = Text(Ventana, height=7, width=40)
t.place(x=40,y=175)
con = Label(Ventana,text="Archivos: " + str(num_arch))
con.place(x=400,y=160)
limit = Label(Ventana,text="25 Mb max.").place(x=400,y=180)
adj = Button(Ventana,text="Adjuntar",command = adjuntarAr).place(x=400,y=210)
cre = Button(Ventana,text="Enviar",command=lambda: enviarMail(entrada1.get(),entrada2.get(),entrada3.get(),entrada4.get(),t.get("1.0", "end-1c"))).place(x=400,y=250)
Ventana.bind('<Control-m>', yo)

Ventana.mainloop()
