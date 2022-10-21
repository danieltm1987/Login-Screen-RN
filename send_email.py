# importa los paquetes necesarios
#import pyodbc
import openpyxl
import pymssql
import pandas as pd
import time
import smtplib
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

#timestr = time.strftime("%Y%m%d")
timestr = time.strftime("20220411")

filename_cons = 'FN100_3033_Supergiros_'+timestr+'.txt'
sql_cons =  """
            SELECT CAST(data.Cedula AS INT) as Cedula
                    , CONVERT(varchar,data.Fecha,112) as Fecha       
                    , CAST(SUM(data.ValorVenta) AS INT) as Valor
                    , COUNT(1) as Cantidad
                    , 516 as Comision 
                FROM (
                SELECT ts.Descripcion as TipoServicio
                    , e.Nombre as Establecimiento
                    , p.Codigo as PuntoVenta
                    , d.serial as Datafono
                    , u.Nombre as usuario
                    , u.Documento as Cedula
                    , i.CodigoContrato as Documento
                    , c.Nombre as NomCliente
                    , i.CodigoInstalacion
                    , i.CodigoContrato
                    , i.Ciudad
                    , i.Barrio
                    , v.NumeroTransaccion
                    , v.Sesion
                    , v.ValorVenta
                    , v.ValorAplicado
                    , v.Fecha 
                FROM TokEnergyAire.dbo.Ventas v
                    ,TokEnergyAire.dbo.Establecimientos e
                    ,TokEnergyAire.dbo.PuntosDeVenta p
                    ,TokEnergyAire.dbo.dispositivos d     
                    ,TokEnergyAire.dbo.Usuarios u       
                    ,TokEnergyAire.dbo.Instalaciones i
                    ,TokEnergyAire.dbo.Clientes c
                    ,TokEnergyAire.dbo.TiposServicios ts
                WHERE e.IdEstablecimiento = 1
                AND p.IdEstablecimiento = e.IdEstablecimiento
                AND d.iddispositivo = p.IdDispositivo   
                AND u.IdUsuario = v.IdUsuario
                AND d.IdDispositivo = v.IdDispositivo
                AND v.IdInstalacion = i.IdInstalacion
                AND v.Fecha >= CONVERT(VARCHAR,SYSDATETIME(),110)                
                --AND v.Fecha BETWEEN CONVERT(SMALLDATETIME,'20220411 00:00:00') AND CONVERT(SMALLDATETIME,'20220411 23:59:59')
                AND c.IdCliente = i.IdCliente
                AND ts.IdTipoServicio = v.IdTipoServicio ) data
                GROUP BY CAST(data.Cedula AS INT)
                        , CONVERT(varchar,data.Fecha,112)
          ;
            """

filename_det = 'FN100_3033_Supergiros_Detallado_'+timestr+'.xlsx'
sql_cdet =  """
            SELECT ts.Descripcion as TipoServicio
                    , e.Nombre as Establecimiento
                    , p.Codigo as PuntoVenta
                    , d.serial as Datafono
                    , u.Nombre as usuario
                    , u.Documento as Cedula
                    , i.CodigoContrato as Documento
                    , c.Nombre as NomCliente
                    , i.CodigoInstalacion
                    , i.CodigoContrato
                    , i.Ciudad
                    , i.Barrio
                    , v.NumeroTransaccion
                    , v.Sesion
                    , v.ValorVenta
                    , v.ValorAplicado
                    , FORMAT(v.Fecha,'dd/MM/yyyy hh:mm:ss tt') as Fecha
                FROM TokEnergyAire.dbo.Ventas v
                    ,TokEnergyAire.dbo.Establecimientos e
                    ,TokEnergyAire.dbo.PuntosDeVenta p
                    ,TokEnergyAire.dbo.dispositivos d     
                    ,TokEnergyAire.dbo.Usuarios u       
                    ,TokEnergyAire.dbo.Instalaciones i
                    ,TokEnergyAire.dbo.Clientes c
                    ,TokEnergyAire.dbo.TiposServicios ts
                WHERE e.IdEstablecimiento = 1
                AND p.IdEstablecimiento = e.IdEstablecimiento
                AND d.iddispositivo = p.IdDispositivo   
                AND u.IdUsuario = v.IdUsuario
                AND d.IdDispositivo = v.IdDispositivo
                AND v.IdInstalacion = i.IdInstalacion
                AND v.Fecha >= CONVERT(VARCHAR,SYSDATETIME(),110)                
                --AND v.Fecha BETWEEN CONVERT(SMALLDATETIME,'20220411 00:00:00') AND CONVERT(SMALLDATETIME,'20220411 23:59:59')
                AND c.IdCliente = i.IdCliente
                AND ts.IdTipoServicio = v.IdTipoServicio
                ;
            """

sender = 'no-responder@air-e.com'
to = ['rigoberto.hernandez@supergirosatlantico.co']
cc = ['luis.leyva@supergirosatlantico.co','roberto.castellar@supergirosatlantico.co','jleivag@air-e.com','alvaro.logreira@air-e.com','paola.nunez@air-e.com','freyser.velasquez@supergiros.com.co','ingrith.franco@supergiros.com.co','melquicide.sanchez@supergirosatlantico.co','elida.gomez@supergirosatlantico.co','saul.gonzalez@supergirosatlantico.co','daniel.torres@air-e.com']
#to = ['daniel.torres@air-e.com']
#cc = ['dtorresm@energiacsc.co']
subject = 'Informe ventas energía prepagada - Super Giros ' + timestr
body = 'Este es un correo generado de forma automatica, favor no responder.'
files = [filename_cons,filename_det]

def create_file(filename,sql,p_header):
    p_server = '10.20.11.101'
    p_user = 'TokEnergyQuery'
    p_pass = 'Cm8rc1564l'
    p_db = 'TokEnergyAire'

    # crea la cadena de conexion
    #conn = pyodbc.connect('Driver={SQL Server};Server='+p_server+';UID='+p_user+';PWD='+p_pass+';Database='+p_db+';')
    conn = pymssql.connect(p_server, p_user, p_pass, p_db, charset='UTF-8' )

    df = pd.read_sql(sql, conn)
    df.to_csv(filename, header=p_header, index=False, sep=';', mode='a')
    conn.close()

def create_file_excel(filename,sql,p_header):
    p_server = '10.20.11.101'
    p_user = 'TokEnergyQuery'
    p_pass = 'Cm8rc1564l'
    p_db = 'TokEnergyAire'

    # crea la cadena de conexion
    #conn = pyodbc.connect('Driver={SQL Server};Server='+p_server+';UID='+p_user+';PWD='+p_pass+';Database='+p_db+';')
    conn = pymssql.connect(p_server, p_user, p_pass, p_db, charset='UTF-8' )

    df = pd.read_sql(sql, conn)
    #df.to_csv(filename, header=p_header, index=False, sep=';', mode='a')
    df.to_excel(filename, sheet_name=timestr, index=False,  engine='openpyxl')    
    conn.close()

def send_email(p_sender,p_to,p_cc,p_subject,p_body,files):
    # parameters smtp
    server = '10.251.139.2'
    port = 25

    # parameters who to send email
    sender = p_sender
    to = p_to
    cc = p_cc
    recipents = to + cc
    subject = p_subject
    body = p_body

    msg = MIMEMultipart()
    msg['From'] = sender
    msg['To'] = ", ".join(to)
    msg['Cc'] = ", ".join(cc)
    msg['Subject'] = subject
    msg.attach(MIMEText(body,'plain'))


    for file in files:
        attachment = open(file, "rb")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload((attachment).read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % file)
        msg.attach(part)

    # sending email
    with smtplib.SMTP(server, port) as server:
        server.sendmail(sender, recipents, msg.as_string())
        server.quit()

def delete_file(files):
    for file in files:
        os.remove(file)

def main():
    create_file(filename_cons,sql_cons,False)
    #create_file(filename_det,sql_cdet,True)
    create_file_excel(filename_det,sql_cdet,True)
    send_email(sender,to,cc,subject,body,files)
    delete_file(files)

main()
