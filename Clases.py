from PyQt5 import QtWidgets
from PyQt5.QtCore import *

from Gui2 import Ui_Actis
from Rutinas_proceso import procesamiento_general


class SignalsProc(QObject):
    """
    En esta clase se definen las señales multithreading que se usarán para indicar eventos en el
    procesado de las ofertas
    """

    informacion = pyqtSignal(str)   # Envía mensajes informativos a la pantalla principal
    completar_libro = pyqtSignal()  # Notifica que se ha terminado el procesamiento
    error = pyqtSignal(str)         # Notifica mensajes de error que aparecen en pantallas separadas


class ProcesarFichero(QRunnable):
    """
    Clase usada para correr en multithreading el procedimiento principal procesamiento_general
    """

    def __init__(self, *args):
        super(ProcesarFichero, self).__init__()
        self.nombre_fichero = args[0]
        self.signals = SignalsProc()

    @pyqtSlot()
    def run(self):
        procesamiento_general(self.nombre_fichero, self)


class GuiPrincipal(QtWidgets.QDialog, Ui_Actis):
    """
    Esta clase es la que gestiona la pantalla principal, y por tanto, la aplicación en su conjunto
    """

    def __init__(self):
        super(self.__class__, self).__init__()
        self.setupUi(self)

        self.boton_Cancel.clicked.connect(self.close)
        self.boton_OK.clicked.connect(self.procesar_oferta)
        self.botonBrowse.clicked.connect(self.buscar_oferta)
        self.avisos.hide()
        self.threadpool = QThreadPool()

    def procesar_oferta(self):
        """
        Procedimiento que efectúa el procesado del fichero de ofertas cuando se pulsa el botón OK
        :return: Devuelve un fichero resultado de nombre Renovaciones.xls
        """
        archivo_entrada = self.fichero_Excel.text()
        self.avisos.show()

        trabajo = ProcesarFichero(archivo_entrada)
        trabajo.signals.informacion.connect(self.reportar)
        trabajo.signals.completar_libro.connect(self.completar_libro)
        trabajo.signals.error.connect(self.mensajes_error)

        self.threadpool.start(trabajo)  # Se envía la tarea principal a una thread separada

    def buscar_oferta(self):  # Lee el nombre del fichero de entrada seleccionado por el usuario
        archivo = QtWidgets.QFileDialog.getOpenFileName(self, "Elegir archivo de oferta")
        self.fichero_Excel.setText(archivo[0])
        print(archivo[0])

    def reportar(self, informacion):
        self.avisos.setText(informacion)  # Escribe la información recibida en señales 'informacion'
        self.avisos.show()

    def completar_libro(self):  # Tareas a efectuar una vez acabado el procesamiento (señal 'completar_libro')

        self.avisos.setText('')
        self.avisos.hide()
        QtWidgets.QMessageBox.information(self, "Información", "Todo correcto")

    def mensajes_error(self, mensaje):  # Presenta mensaje de error (de la señal 'error') en pantalla separada
        self.avisos.setText('')
        self.avisos.hide()
        QtWidgets.QMessageBox.critical(self, "Error", mensaje)
