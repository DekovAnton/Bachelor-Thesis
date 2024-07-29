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







#Worker multithreading
class Worker(QThread):
    update = pyqtSignal(int)
    dictupdate=pyqtSignal(dict)
    # Init method
    def __init__(self):
        super().__init__()
        self.vdcaddr=""
        self.dummyaddr=""
        self.vin=""
        self.startcurr=""
        self.endcurr=""
        self.step=""
        self.delay=""
        self.name=""
        self.cin=""
        self.DC_Load_reset=""
        self.VDC_reset=""
        self.Discharge_cap_on_exit=""

    def variables(self,vdcaddr,dummyaddr,vin,cin,startcurr,endcurr,step,delay,name,DC_Load_Reset,vdc_reset,capacitordis):
        try:
            self.vdcaddr = vdcaddr
            self.dummyaddr = dummyaddr
            self.vin = vin
            self.startcurr = startcurr
            self.endcurr = endcurr
            self.step = step
            self.delay = delay
            self.name = name
            self.cin = cin
            self.DC_Load_reset = DC_Load_Reset
            self.VDC_reset = vdc_reset
            self.Discharge_cap_on_exit=capacitordis

            # calculate variables
            self.steps = (float(self.endcurr) - float(self.startcurr)) / float(self.step)
            self.sleeptime = float(delay) / float(1000)

            difer = float(self.endcurr) - float(self.startcurr)
            self.steps = int(float(difer) / float(step))
            self.curvalues = [round(float(startcurr) + float(step) * float(i), 3) for i in range(self.steps + 1)]
            self.progress = 0

        except Exception as e:
            print(e)


    def stop(self):
        #method do change stop request ==True
        self.StopRequest= True

    def run(self):
        #method to assign stoprequest==False and run simulation.
        self.StopRequest= False
        try:
            self.simulate()
        except Exception as e:
            print(e)

    def simulate(self):
        try:
            point = 100 / len(self.curvalues)
            # VDC init
            self.rm = pyvisa.ResourceManager()
            toolvdc = self.rm.open_resource(self.vdcaddr)

            toolvdc.write("*RST")
            toolvdc.write("APPLy " + str(self.vin) + "," + str(self.cin))
            toolvdc.write("DISPlay:MENU:NAME 3")

            # dc load init
            tooldummy = self.rm.open_resource(self.dummyaddr)
            tooldummy.write("*RST")
            CurrInList = []
            PowerInList = []
            VoltInList = []
            CurrOutList = []
            PowerOutList = []
            VoltOutList = []
            timestamp = []
            step = []
            # cycle for measuring
            for i in self.curvalues:
                # stop when testing
                if self.StopRequest:
                    return

                # turn Dummy and then DC
                tooldummy.write(":CURRent " + str(i) + "A")
                tooldummy.write(":INPut ON")
                toolvdc.write("OUTP:STAT:IMM ON")
                # time sleep for dc-dc
                time.sleep(self.sleeptime)

                # measure current,voltage, power from VDC
                measCurrIn = toolvdc.query("MEASure:SCALar:CURRent:DC?")
                CurrInList.append(str(measCurrIn).replace("+","").strip())
                print("INPUT CURRENT: " + str(measCurrIn))
                measVoltIn = toolvdc.query("MEASure:SCALar:VOLTage:DC?")
                VoltInList.append(str(measVoltIn).replace("+","").strip())
                print("INPUT VOLTAGE: " + str(measVoltIn))
                measPowerIn = toolvdc.query("MEASure:SCALar:POWer:DC?")
                PowerInList.append(str(measPowerIn).replace("+","").strip())
                print("INPUT POWER: " + str(measPowerIn))

                # measure current,voltage,power from DC Load
                measCurrOut = tooldummy.query(":MEASure:CURRent?")
                CurrOutList.append(str(measCurrOut).replace("A","").strip())
                print("OUTPUT CURRENT: " + str(measCurrOut))
                measVoltOut = tooldummy.query(":MEASure:VOLTage?")
                VoltOutList.append(str(measVoltOut).replace("V","").strip())
                print("OUTPUT VOLTAGE: " + str(measVoltOut))
                measPowerOut = tooldummy.query(":MEASure:POWer?")
                PowerOutList.append(str(measPowerOut).replace("W","").strip())
                print("OUTPUT POWER: " + str(measPowerOut))
                timestamp.append(datetime.datetime.today())
                step.append(i)

                # progress and time.sleep
                self.progress = self.progress + round(point)
                self.update.emit(int(self.progress))

                # Turn VDC off and Dummy
                if self.VDC_reset:
                    toolvdc.write("OUTP:STAT:IMM OFF")
                if self.DC_Load_reset:
                    tooldummy.write(":INPut OFF")
                if self.VDC_reset or self.DC_Load_reset:
                    time.sleep(self.sleeptime)

            # stop VDC and close connection
            toolvdc.write("OUTP:STAT:IMM OFF")
            toolvdc.close()
            if self.Discharge_cap_on_exit:
                time.sleep(self.sleeptime)
            tooldummy.write(":INPut OFF")
            tooldummy.close()

            self.measure_dict={""}
            self.progress = 100
            self.update.emit(int(self.progress))
            measure_dict={"TimeStamp":timestamp,"Step":step,"measCurrIn":CurrInList,"measPowerIn":PowerInList,"measVoltIn":VoltInList,"measCurrOut":CurrOutList,"measPowerOut":PowerOutList,"measVoltOut":VoltOutList}
            self.dictupdate.emit(measure_dict)
        except Exception as e:
            print(e)







#Main Window code
class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow,self).__init__()
        uic.loadUi("untitled.ui",self)
        ### Statements for mainwindow
        self.progressBar.setValue(0)
        self.COMSettings.clicked.connect(self.runqdialog)
        self.Emergencystop.clicked.connect(self.disconnectVDC)

        #Run Worker
        self.worker=Worker()
        self.worker.update.connect(self.progress)
        self.worker.dictupdate.connect(self.process_Data)
        self.Run.clicked.connect(self.run)
        self.Stopbutton.clicked.connect(self.stop)
        self.df_toSave=pd.DataFrame()

    def stop(self):
        self.worker.stop()

        try:
            inst = self.rm.open_resource(self.VDC[0])
            inst.write("OUTP:STAT:IMM OFF")
            inst.write("*RST")
            inst.close()
            dummy=self.rm.open_resource((self.dummy[0]))
            dummy.write(":INPut OFF")
            dummy.close()
            self.message("Measurements aborted!")
        except Exception as e:
            print(e)
        finally:
            self.progressBar.setValue(0)

    def run(self):
        try:
            #self.Convertername.setText("Levski")
            name=self.Convertername.text()
            vdcaddr=self.VDC[0]
            dummyaddr=self.dummy[0]
            vin=self.Vin.text()
            cin=self.cstart.text()
            startcurr=self.CurStart.text()

            endcurr=self.Curend.text()
            step=self.CurStep.text()
            delay=self.Delay.text()
            DC_Load_Reset=self.Load_reset.isChecked()
            vdc_reset=self.vdc_reset.isChecked()
            capacitordis=self.Discharge_cap.isChecked()
            if vdcaddr=="":
                raise
            elif dummyaddr=="":
                raise
            elif name=="":
                raise
            floatvalues=[vin,cin,startcurr,endcurr,step,delay,DC_Load_Reset,vdc_reset,capacitordis]
            for value in floatvalues:
                float(value)
        except Exception as e:
            print(e)
            self.message("Incorrect parameter input!")
            return
        try:
            self.worker.variables(vdcaddr,dummyaddr,vin,cin,startcurr,endcurr,step,delay,name,DC_Load_Reset,vdc_reset,capacitordis)
            self.worker.start()
        except Exception as e:
            print(e)

    def progress(self, var):
        self.progressBar.setValue(var)

    #qdialog for COMports
    def runqdialog(self):
        # initialize qdialog window
        dialog=QDialog(self)
        uic.loadUi("untitled2.ui",dialog)

        # Find Comports
        try:
            self.rm = pyvisa.ResourceManager()
            self.SerialComports = list(self.rm.list_resources())
        except:
            self.message("Serial Comports not available!")

        # Check Comports available
        if len(self.SerialComports)<1:
            self.message("No tools connected!")

        ##get IDN from serials
        self.SerialComportsidn=[self.rm.open_resource(port).query("*IDN?").strip() for port in self.SerialComports]
        self.resources={y: self.SerialComportsidn[x] for x,y in enumerate(self.SerialComports)}
        self.resources[None]=None

        #Set Comports in dropdown menu
        dialog.VDCdrop.addItems(list(self.resources.values()))
        dialog.Dummydrop.addItems(list(self.resources.values()))

        ## statements for dialog
        dialog.Canceldialog.clicked.connect(dialog.reject)
        dialog.Okdialog.clicked.connect(self.getCOMS)
        self.runqdialog=dialog
        dialog.exec()

    def process_Data(self,measure_dict):
        try:
            self.clear_plots()
            print(measure_dict)
            for key,value in measure_dict.items():
                try:
                    measure_dict[key]=[float(x) for x in value]
                except Exception as e:
                    print(e)
            df=pd.DataFrame.from_dict(measure_dict)
            df["Calculated Pin"]=df["measCurrIn"]*df["measVoltIn"]
            df["Calculated Pout"]=df["measCurrOut"]*df["measVoltOut"]
            df["n"]=df["Calculated Pout"]/df["Calculated Pin"]
            df["Converter Name"]=self.Convertername.text()


            x_axis = df["measCurrOut"]
            y_axis = df["n"] * 100
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

    #Get COMports and write to locked lines
    def getCOMS(self):
        #get text from drop down menu
        vdccom=self.runqdialog.VDCdrop.currentText()
        self.VDC=[key for key,value in self.resources.items() if value==vdccom]
        dummycom=self.runqdialog.Dummydrop.currentText()
        self.dummy=[key for key,value in self.resources.items() if value==dummycom]

        #set text on main window editline and close Qdialog
        self.VDClockedline.setText(vdccom)
        self.Dummylockedline.setText(dummycom)
        self.runqdialog.accept()


    #Emergency Stop VDC
    def disconnectVDC(self):
        try:
            inst = self.rm.open_resource(self.VDC[0])
            inst.write("OUTP:STAT:IMM OFF")
            inst.write("*RST")
            inst.close()
            tooldummy=self.rm.open_resource((self.dummy[0]))
            tooldummy.write(":INPut OFF")
            tooldummy.close()
        except Exception as e:
            print(e)
            self.message("Can't stop VDC Supply. Please stop supply manually!")
        finally:
            self.stop()
            return


    def clear_plots(self):
        try:
            layout = self.PlotFrame.layout()
            if layout:
                while layout.count():
                    item = layout.takeAt(0)
                    widget = item.widget()
                    if widget:
                        widget.deleteLater()
            time.sleep(1)
        except Exception as e:
            print(e)

    #Error message pop-up
    def message(self, errortext):
        error_message = QMessageBox()
        error_message.setIcon(QMessageBox.Icon.Critical)
        error_message.setWindowTitle("Error")
        error_message.setText(errortext)
        error_message.setStandardButtons(QMessageBox.StandardButton.Ok)
        error_message.exec()

    def Save(self):
        options=QFileDialog.Options()

if __name__=="__main__":
    import sys
    app=QApplication(sys.argv)
    window=MainWindow()
    window.show()
    sys.exit(app.exec())