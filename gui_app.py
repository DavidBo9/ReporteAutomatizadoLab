import sys
import os
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QFileDialog, QLabel, QMessageBox, 
                             QLineEdit, QDateEdit, QGroupBox, QFormLayout)
from PyQt5.QtCore import QDir, QDate
from reporte_logica import generar_reporte_completo 

class ReportGeneratorApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Generador de Reportes de Inventario Lab (Datos Dinámicos)')
        self.archivo_excel = None
        self.ruta_plantilla = os.path.join('plantilla', 'Reporte_Plantilla.docx') 
        self.initUI()

    def initUI(self):
        vbox = QVBoxLayout()

        self.title_label = QLabel("Generación Automática de Reporte Mensual")
        self.title_label.setStyleSheet("font-size: 16pt; font-weight: bold;")
        vbox.addWidget(self.title_label)
        
        # --- Grupo 1: Configuración de Fechas ---
        date_group = QGroupBox("📅 1. Periodo de Control")
        date_layout = QFormLayout()
        
        self.input_mes = QLineEdit()
        self.input_mes.setPlaceholderText(QDate.currentDate().toString('MMMM yyyy'))
        date_layout.addRow("Mes del Reporte (Ej: Mayo 2026):", self.input_mes)
        
        self.input_fecha_inicio = QDateEdit(calendarPopup=True)
        self.input_fecha_inicio.setDate(QDate.currentDate().addDays(-30))
        date_layout.addRow("Fecha de Inicio de Conteo:", self.input_fecha_inicio)
        
        self.input_fecha_fin = QDateEdit(calendarPopup=True)
        self.input_fecha_fin.setDate(QDate.currentDate())
        date_layout.addRow("Fecha de Fin de Conteo:", self.input_fecha_fin)
        
        date_group.setLayout(date_layout)
        vbox.addWidget(date_group)
        
        # --- Grupo 2: Responsabilidades ---
        resp_group = QGroupBox("🧑‍🔬 2. Responsables del Documento")
        resp_layout = QFormLayout()
        
        self.input_dir = QLineEdit()
        self.input_dir.setPlaceholderText("Nombre del Director/Supervisor")
        resp_layout.addRow("Direccional/Área Supervisora:", self.input_dir)
        
        self.input_resp = QLineEdit()
        self.input_resp.setPlaceholderText("Nombre del Ejecutor de Inventario")
        resp_layout.addRow("Responsable de Inventario:", self.input_resp)
        
        self.input_verif = QLineEdit()
        self.input_verif.setPlaceholderText("Nombre del Verificador (Calidad)")
        resp_layout.addRow("Verificado por:", self.input_verif)
        
        resp_group.setLayout(resp_layout)
        vbox.addWidget(resp_group)
        
        # --- Grupo 3: Archivos y Ejecución ---
        file_group = QGroupBox("🗃️ 3. Archivo y Generación")
        file_layout = QVBoxLayout()

        self.btn_select = QPushButton('Seleccionar Archivo de Datos Excel (.xlsx)')
        self.btn_select.clicked.connect(self.select_excel_file)
        file_layout.addWidget(self.btn_select)

        self.label_path = QLabel("Archivo seleccionado: Ninguno")
        file_layout.addWidget(self.label_path)

        self.btn_generate = QPushButton('GENERAR Y GUARDAR REPORTE WORD')
        self.btn_generate.clicked.connect(self.generate_report)
        self.btn_generate.setEnabled(False) 
        self.btn_generate.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        file_layout.addWidget(self.btn_generate)
        
        file_group.setLayout(file_layout)
        vbox.addWidget(file_group)

        # Etiqueta de Estado
        self.status_label = QLabel("Estado: Esperando selección de archivo.")
        vbox.addWidget(self.status_label)
        
        self.setLayout(vbox)

    def select_excel_file(self):
        start_dir = os.path.join(QDir.currentPath(), 'inventario')
        
        fileName, _ = QFileDialog.getOpenFileName(self, "Seleccionar Archivo de Inventario", start_dir,
                                                  "Archivos Excel (*.xlsx);;Todos los Archivos (*)")
        
        if fileName:
            self.archivo_excel = fileName
            self.label_path.setText(f"Archivo: {os.path.basename(fileName)}")
            self.btn_generate.setEnabled(True)
            self.status_label.setText("Listo. Presione el botón para generar el Word.")
        else:
            self.archivo_excel = None
            self.label_path.setText("Archivo seleccionado: Ninguno")
            self.btn_generate.setEnabled(False)
            self.status_label.setText("Debe seleccionar un archivo Excel.")

    def generate_report(self):
        if not self.archivo_excel:
            QMessageBox.warning(self, "Advertencia", "Por favor, seleccione un archivo Excel primero.")
            return
            
        # 1. Capturar todos los datos de la GUI
        datos_gui = {
            'mes': self.input_mes.text() or self.input_mes.placeholderText(),
            'fecha_inicio': self.input_fecha_inicio.date().toString("yyyy-MM-dd"),
            'fecha_fin': self.input_fecha_fin.date().toString("yyyy-MM-dd"),
            'dir_nombre': self.input_dir.text() or self.input_dir.placeholderText(),
            'resp_nombre': self.input_resp.text() or self.input_resp.placeholderText(),
            'verif_nombre': self.input_verif.text() or self.input_verif.placeholderText(),
        }

        self.status_label.setText("Estado: Generando reporte... NO CIERRE EL PROGRAMA.")
        QApplication.processEvents() 

        # 2. LLAMADA A LA LÓGICA CON LOS NUEVOS PARÁMETROS
        success, result = generar_reporte_completo(
            self.archivo_excel, self.ruta_plantilla, datos_gui
        )
        
        if success:
            QMessageBox.information(self, "Éxito", f"Reporte generado exitosamente:\nGuardado en: {result}")
            self.status_label.setText(f"Estado: Reporte finalizado. Guardado en {os.path.dirname(result)}")
        else:
            QMessageBox.critical(self, "Error", f"Ocurrió un error: {result}")
            self.status_label.setText("Estado: ¡ERROR! Verifique los datos o el Excel.")

# Bloque main omitido por ser el mismo que en el archivo main.py