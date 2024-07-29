from PyQt6 import *
from PyQt6.QtWidgets import *
from PyQt6 import uic
from PyQt6.QtCore import QThread, pyqtSignal
import pyvisa
import time
import serial.tools.list_ports
import threading
import pandas as pd
import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg
import os





class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow,self).__init__()
        uic.loadUi("untitled.ui",self)
        ### Statements for mainwindow
        self.progressBar.setValue(0)
        #Run Worker
        self.df_toSave=pd.DataFrame()
        self.Run.clicked.connect(self.run)
        self.Stopbutton.clicked.connect(self.clear_plots)
        self.save_button.clicked.connect(self.Save)


    def run(self):
        self.clear_plots()
        try:
            self.df=pd.read_excel("Measurements_2.xlsx")
            x_axis = self.df["measCurrOut"]
            y_axis = self.df["n"] * 100
            y_axis = [round(y, 2) for y in y_axis]
            y_axis_points = list(range(0, 101, 5))

            # Create a new plot
            fig, ax = plt.subplots()
            ax.plot(x_axis, y_axis, marker="o", linestyle="--")
            ax.set_xlabel("Output Current [A]")
            ax.set_ylabel("Efficiency [%]")
            ax.set_title("Efficiency of DC-DC Converter")
            ax.set_yticks(y_axis_points)
            ax.grid(True)

            # Create a canvas for the plot
            canvas = FigureCanvasQTAgg(fig)

            # Get or create the layout for the PlotFrame
            layout = self.PlotFrame.layout()
            if layout is None:
                layout = QVBoxLayout(self.PlotFrame)
                self.PlotFrame.setLayout(layout)

            # Add the canvas to the layout
            layout.addWidget(canvas)
        except Exception as e:
            print(e)

    def clear_plots(self):
        try:
            layout = self.PlotFrame.layout()
            if layout:
                while layout.count():
                    item = layout.takeAt(0)
                    widget = item.widget()
                    if widget:
                        widget.deleteLater()
        except Exception as e:
            print(e)

    def Save(self):
        try:
            file_name, _ = QFileDialog.getSaveFileName(self, "Save Excel File", "", "Excel Files (*.xlsx)")
            if file_name:
                self.df.to_excel(file_name,index=False)
        except Exception as e:
            print(e)
            self.message("Error saving file!")

    #Error message pop-up
    def message(self, errortext):
        error_message = QMessageBox()
        error_message.setIcon(QMessageBox.Icon.Critical)
        error_message.setWindowTitle("Error")
        error_message.setText(errortext)
        error_message.setStandardButtons(QMessageBox.StandardButton.Ok)
        error_message.exec()


if __name__=="__main__":
    import sys
    app=QApplication(sys.argv)
    window=MainWindow()
    window.show()
    sys.exit(app.exec())