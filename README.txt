Guía de Código:
Librerias:
openpyxl -> Es un módilo d ePython que permite leer, crear y modificar ficheros Excel.

# leemos el fichero
libro = load_workbook('fichero.xlsx')

# obtenemos la pestaña/hoja activa (nada mas abrir es la primera)
hoja = libro.active

# mostramos la celda B1 = 1:2
print(hoja.cell(row=1, column=2).value)
print(hoja['B1'].value)

Getpass:
Solicita al usuario una contraseña sin hacer eco.

ssl:
Este módulo brinda acceso a la seguridad de la capa de transporte (a menudo conocida como "capa de sockets seguros") y las funciones de autenticación de pares para los sockets de red, tanto del lado del cliente como del lado del servidor. Este módulo utiliza la biblioteca OpenSSL.

create_default_context()->devuelven un nuevo contexto con configuraciones predeterminadas seguras.



smtplib:
El módulo smtplib define un objeto de sesión de cliente SMTP que se puede utilizar para enviar correo a cualquier máquina de Internet con un demonio de escucha SMTP o ESMTP.

smtplib.SMTP_SSL(host='', port=0, local_hostname=None, keyfile=None, certfile=None, [timeout, ]context=None, source_address=None)
Una instancia de SMTP_SSL se comporta exactamente igual que las instancias de SMTP. SMTP_SSL debe usarse para situaciones donde se requiere SSL desde el comienzo de la conexión y el uso starttls() no es apropiado. Si no se especifica host, se utiliza el host local. Si port es cero, se utiliza el puerto estándar SMTP sobre SSL (465). Los argumentos opcionales local_hostname, timeout y source_address tienen el mismo significado que en la clase SMTP. context, también opcional, puede contener una SSLContext y permite configurar varios aspectos de la conexión segura

class email.mime.multipart.MIMEMultipart(_subtype='mixed', boundary=None, _subparts=None, *, policy=compat32, **_params)
Módulo: email.mime.multipart

Una subclase de MIMEBase, se trata de una clase base intermedia para los mensajes MIME que son multipart. El valor predeterminado opcional de _subtype es mixed, pero se puede utilizar para especificar el subtipo del mensaje. Se agregará un encabezado Content-Type de multipart/_subtype al objeto del mensaje. También se agregará un encabezado MIME-Version.


