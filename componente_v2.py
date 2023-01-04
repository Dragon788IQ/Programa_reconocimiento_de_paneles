from matplotlib import pyplot as plt
from win10toast import ToastNotifier
import pandas as pd
import numpy as np
import cv2
import os
import easyocr
import random
import time
import datetime
import win32api



#Declaracion de las variables de inciación
drawing = False
point1 = ()
point2 = ()
camera_selection = True
camera_crop = True
registro_T= []
registro_R =[]
registro_time = []
global_aux = 0
global_time = 0
inciador_lector = False
registrador_excel = 0
ID = str(random.randint(1000,2000))
fecha = datetime.datetime.now()
fecha = fecha.strftime("%Y_%m_%d")

#Declaracion de la funcion que permite seleccionar el area a leer
def mouse_drawing(event, x, y, flags, params):
    global point1, point2, drawing

    if event == cv2.EVENT_LBUTTONDOWN:
        if drawing is False:
            drawing = True
            point1 = (x, y)
        else:
            drawing = False

    elif event == cv2.EVENT_MOUSEMOVE:
        if drawing is True:
            point2 = (x, y)


#Declaracion de la funcion que permite hacer y guardar las graficas de las lecturas
def graficador(x, y, title, x_label, y_label, style):
    try:
        plt.style.use(style)
        plt.grid()
        plt.scatter(x,y, color="red")
        #plt.ylim([min(y)-(min(y)*0.1), (max(y)+(max(y)*0.1))]) 
        plt.xlabel(x_label)
        plt.ylabel(y_label)
        plt.title(title)
        plt.show(block=False)
        plt.pause(2)
        nombre_grafica = "Grafica_experimental_" + fecha +"_"+ID+".png"
        plt.savefig(nombre_grafica)
        plt.close()
    except:
        pass

def Excel_saver(fecha, ID, registro_time, registro_T):
        nombre_documento = "Datos_Experimento_" + fecha + "_" + ID +".xlsx"
        ruta_resultado = "Resultados/" + nombre_documento
        if not os.path.exists("Resultados"):
            os.makedirs("Resultados")
        Experiment_data = pd.DataFrame({"Tiempo": registro_time, "Temperatura": registro_T})
        writer = pd.ExcelWriter(ruta_resultado, engine='xlsxwriter')
        Experiment_data.to_excel(writer, sheet_name='Sheet1', index='False')
        writer.save()
        print("Datos salvados en el excel")

#Desplegue del mensaje de las instrucciones
win32api.MessageBox(0, 
    """
    Paso 1: Dibujar el rectangulo donde será el area de lectura.
    
    Paso 2: Una vez seleccionada el area presionar la tecla Q para pasar al modo de trabajo.

    Paso 3: Una vez que se comprobo que la imagen es la optima para la lectura precionar la tecla C para comenzar las lecturas

    Paso 4: De forma automatica se registraran los datos en un excel y la grafica en un .PNG

    Extra: Si se requiere se puede pausar el sistema de lectura presionando la tecla P y precionando de nuevo la misma tecla se reanuda.

    Nota: Mantener una distancia correcta entre la camara y el panel cuidando de tener una buena illiminación para detectar de forma correcta los datos
    """
, 'Manual de uso')

win32api.MessageBox(0, 
    """
    Notas a conciderar:
    -La incialización de la camara puede tardar unos segundos.
    -La frecuencia de la toma de datos puede variar un poco dependiendo de las especificaciones del equipo de computo.
    """
, 'Manual de uso')


#Inicializacion de la camara
try:
    cap = cv2.VideoCapture(1) #Intentamos iniciar la camara USB
except:
    cap = cv2.VideoCapture(0) #Si falla la camara usamos la de la lap

#Contruccion de la ventana de la camara
cv2.namedWindow("Frame")
cv2.setMouseCallback("Frame", mouse_drawing) #Inicializamos la escucha de los eventos

#Loop encargado de permitirnos dibujar la seleccion a leer y registrarla
while camera_selection == True:
    _, frame = cap.read()

    if point1 and point2:
        cv2.rectangle(frame, point1, point2, (0, 255, 0))

    cv2.imshow("Frame", frame)
    key = cv2.waitKey(1)

    #Aqui declaramos que cerraremos la ventana con la tecla "Q"
    if cv2.waitKey(1) & 0xFF == ord('q'):
        camera_selection= False

cv2.destroyAllWindows()

#Main loop donde funciona la lectura del programa
while camera_crop == True:
    #Lectura de la camara con solo la seleccion elegida previamente
    _, frame = cap.read()
    frame = np.array(frame) #transformacion de la imagen a una matriz
    frame = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY) #Aplicacion de un filtro blanco y negro
    frame = 255-frame #Aplicacion de un filtro negativo

    #Lanzamiento de la imagen de la camara
    cv2.imshow("Cropped", frame[point1[1]:point2[1], point1[0]:point2[0]])

    #Aumento de unidad para la variable que nos da un time span
    global_aux += 1
    print(global_aux)

    #Aqui declaramos que comenzaremos con el escaneo presionando la tecla "C"
    if cv2.waitKey(1) & 0xFF == ord('c'):
        toast = ToastNotifier()
        inciador_lector = True
        toast.show_toast(
            "Programa de lectura automatica",
            "Se ah iniciado el registro automatico, por favor de no mover la camara",
            duration = 5,
            threaded = True,
        )

    #Aqui declaramos que pausaremos el escaneo presionando la tecla "P"
    if cv2.waitKey(1) & 0xFF == ord('p'):
        if inciador_lector == True:
            toast = ToastNotifier()
            toast.show_toast(
                "Programa de lectura automatica",
                "Se ha detenido la lectura automatica para activar de nuevo precione de nuevo la tecla P",
                duration = 5,
                threaded = True,
            )
            inciador_lector = not inciador_lector
        else:
            inciador_lector = not inciador_lector
            toast = ToastNotifier()
            toast.show_toast(
                "Programa de lectura automatica",
                "Se ha reanudado el sistema de lectura automatica",
                duration = 5,
                threaded = True,
            )
        Excel_saver(fecha, ID, registro_time, registro_T)



    while inciador_lector == True and global_aux >= 125:
        global_time = time.time()

        try:
            global_aux = 0
            frame_to_read = frame[point1[1]:point2[1], point1[0]:point2[0]]
            time_start = time.time()
            Reader = easyocr.Reader(["en"], gpu=False)
            result = Reader.readtext(frame_to_read, detail=0,  allowlist ='0123456789')
            print(f"El resultado de Temperatura es: {float(result[0])/10} °C")
            registro_T.append(float(result[0])/10)
            #print(f"A la referencia medida es : {float(result[1])/10}")
            #registro_R.append(result[1])
            deltaTime = time.time() - time_start
            print(f"Tiempo de procesamiento= {deltaTime}")
        except:
            registro_T.append(0)


        if len(registro_time)==0:
            registro_time.append(time.time() - global_time)
        elif len(registro_time) ==1:
            registro_time.append(registro_time[-1]+(time.time() - global_time))
        else:
            registrador_excel = len(registro_T)
            registro_time.append(registro_time[-1]+(time.time() - global_time))
            #registro_time.append(time.time() - global_time)
            #registro_time.append(registro_time[-1] + deltaTime)
            graficador(registro_time, registro_T,"Datos experimentales","Tiempo [s]","Temperatura [°C]","ggplot")


    if (registrador_excel%5) == 0:
        registrador_excel += 1
        print("Los resultados de Temperatura guardados son:")
        print(registro_T)
        print("Los resultados de Tiempo guardados son:")
        print(registro_time)
        Excel_saver(fecha, ID, registro_time, registro_T)


    key = cv2.waitKey(1)
    if cv2.waitKey(1) & 0xFF == ord('q'):
        camera_crop = False