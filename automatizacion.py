"""
Autor: Jyron Cedeño
Fecha: 18-12-2020

INDICACIONES: Deben tener el archivo excel en el mismo directorio del archivo .py 


Debes iniciar sesión con una cuenta de gmail, no olvidar que previo a eso debes haber configurado
tú cuenta de gmail para permitir acceso a aplicaciones no seguras, esto con el fin de que los emails
no se vayan a spam

en en el siguiente link puedes ver como configurar gmail: https://docs.rocketbot.co/?p=1567 

"""




import openpyxl
import smtplib, ssl
import getpass
from socket import gaierror
from email.mime.text import MIMEText
from email.header import Header
from email.mime.multipart import MIMEMultipart


class AutomatizacionExcel:
    def leerExcel(self):
        libro = openpyxl.load_workbook('datos.xlsx', data_only=True)
        hoja = libro.active
        celdas = hoja['A2':'G6']
        return(celdas)

class EnviarCorreo(AutomatizacionExcel):
    def __init__(self, usuario, contrasena):
        self.usuario = usuario 
        self.contrasena = contrasena 

    def conectarEmail(self):    

        try:
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
                server.login(self.usuario, self.contrasena)
                EnviarCorreo.sesionIniciada(self)
                print("\n")
                lista_empleado = list()
                celdas = AutomatizacionExcel.leerExcel(self)
                for filas in celdas:
                    empleado = [celda.value  for celda in filas]
                    #print(empleado)
                    lista_empleado.append(empleado)
                for empleado in lista_empleado:  
                    mensaje = MIMEMultipart("alternative")
                    mensaje["Subject"] = "Información de empleados"
                    mensaje["From"] = self.usuario
                    mensaje["To"] = empleado[6]
                    html = f"""
                        
                                <!DOCTYPE html>
                                <html lang="en">
                                <head>
                                    <meta charset="UTF-8">
                                    <meta name="viewport" content="width=device-width, initial-scale=1.0">
                                    <!-- Font Awesome -->
                                    <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.8.2/css/all.css">
                                    <!-- Google Fonts -->
                                    <link rel="stylesheet" href="https://fonts.googleapis.com/css?family=Roboto:300,400,500,700&display=swap">
                                    <!-- Bootstrap core CSS -->
                                    <link href="https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/4.5.0/css/bootstrap.min.css" rel="stylesheet">
                                    <!-- Material Design Bootstrap -->
                                    <link href="https://cdnjs.cloudflare.com/ajax/libs/mdbootstrap/4.19.0/css/mdb.min.css" rel="stylesheet">	
                                    <title>Document</title>
                                </head>
                                <body>
                                <!-- Card deck -->
                                <div class="card-deck p-5 m-5">

                                <!-- Card -->
                                <div class="card mb-4">

                                    <!--Card image-->
                                    <div class="view overlay">
                                    <img class="card-img-top" src="https://mdbootstrap.com/img/Photos/Others/images/16.jpg"
                                        alt="Card image cap">
                                    <a href="#!">
                                        <div class="mask rgba-white-slight"></div>
                                    </a>
                                    </div>

                                    <!--Card content-->
                                    <div class="card-body">

                                    <!--Title-->
                                    <h4 class="card-title">Hola, {empleado[0]}  {empleado[1]}</h4>
                                    <!--Text-->
                                    <p class="card-text">Usted tiene el cargo de {empleado[3]} y su salario es de: ${empleado[4]}</p>
                                    <p class="bold">Gracias por ser parte de esta grandiosa empresa.</p>
                                    <!-- Provides extra visual weight and identifies the primary action in a set of buttons -->
                                    <button type="button" class="btn btn-light-blue btn-md">Read more</button>

                                    </div>

                                </div>
                                <!-- Card -->

                                </div>
                                <!-- Card deck -->


                                    <!-- JQuery -->
                                    <script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/jquery/3.5.1/jquery.min.js"></script>
                                    <!-- Bootstrap tooltips -->
                                    <script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.14.4/umd/popper.min.js"></script>
                                    <!-- Bootstrap core JavaScript -->
                                    <script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/4.5.0/js/bootstrap.min.js"></script>
                                    <!-- MDB core JavaScript -->
                                    <script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/mdbootstrap/4.19.0/js/mdb.min.js"></script>
                                    
                                </body>
                                </html>
                                """
                    parte_html = MIMEText(html, 'html')
                    mensaje.attach(parte_html)

                    if empleado[6]:
                        try:
                            server.sendmail(self.usuario, empleado[6], mensaje.as_string())
                            EnviarCorreo.mailEnviado(self)
                        except (gaierror, ConnectionRefusedError):
                            EnviarCorreo.mailNoEnviado(self) 
                    else:
                        try:
                            print('El empleado {}, no tiene email registrado, por lo que no se notificará.'.format(empleado[0]+' '+empleado[1])) 
                        except (gaierror, ConnectionRefusedError):
                            EnviarCorreo.mailNoEnviado(self)                             
                           

        except (gaierror, ConnectionRefusedError):
            print('Error al conectar con el servidor. ¿Mala configuración de conexión?')
        except smtplib.SMTPServerDisconnected:
            print('Error al conectar con el servidor. ')
            EnviarCorreo.sesionFallida(self)
        except smtplib.SMTPException as e:
            print('Ocurrió un error de SMTP, es posible que el Usuario y contraseña sean incorrectos, o su configuración de gmail no permita acceder: ' + str(e))

    def sesionIniciada(self):
        print("Sesión iniciada")
    def sesionFallida(self): 
        print("Error la sesión no se ha podido iniciar sesión, intente nuevamente")  
    def mailEnviado(self):
        print('El mail se envió correctamente')
    def mailNoEnviado(self):
        print('El correo no se pudo entregar')


objAutomatizacion = EnviarCorreo(input("Ingrese su correo de logueo gmail: "), getpass.getpass("Ingrese su contaseña: "))
objAutomatizacion.conectarEmail()