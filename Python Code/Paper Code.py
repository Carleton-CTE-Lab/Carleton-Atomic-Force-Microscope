import time
import clr
clr.AddReference("C:\\Program Files\\Thorlabs\\Kinesis\\Thorlabs.MotionControl.DeviceManagerCLI.dll")
clr.AddReference("C:\\Program Files\\Thorlabs\\Kinesis\\Thorlabs.MotionControl.GenericMotorCLI.dll")
clr.AddReference("C:\\Program Files\\Thorlabs\\Kinesis\\ThorLabs.MotionControl.KCube.PiezoCLI.dll")
clr.AddReference("C:\\Program Files\\Thorlabs\\Kinesis\\ThorLabs.MotionControl.KCube.StrainGaugeCLI.dll")
clr.AddReference("C:\\Program Files\\Thorlabs\\Kinesis\\Thorlabs.MotionControl.KCube.PositionAlignerCLI.dll")
from Thorlabs.MotionControl.DeviceManagerCLI import *
from Thorlabs.MotionControl.GenericMotorCLI import *
from Thorlabs.MotionControl.GenericPiezoCLI import *
from Thorlabs.MotionControl.KCube.PiezoCLI import *
from Thorlabs.MotionControl.KCube.StrainGaugeCLI import *
from Thorlabs.MotionControl.KCube.PositionAlignerCLI import *
from System import Decimal  # necessary for real world units #
import nidaqmx
import matplotlib.pyplot as plt
import tkinter as tk
from PIL import ImageTk, Image
from tkinter import messagebox
from datetime import datetime

# create new device
serial_P = "29250983"
serial_S = "59000880"
serial_A = "69252146"

def main():

    try:
        DeviceManagerCLI.BuildDeviceList()

        # Connect, begin polling, and enable
        Piezo = KCubePiezo.CreateKCubePiezo(serial_P)
        Strain = KCubeStrainGauge.CreateKCubeStrainGauge(serial_S)
        Aligner = KCubePositionAligner.CreateKCubePositionAligner(serial_A)

        Piezo.Connect(serial_P)
        Strain.Connect(serial_S)
        Aligner.Connect(serial_A)

        # Get Device Information and display description
        Piezo_info = Piezo.GetDeviceInfo()
        Strain_info = Strain.GetDeviceInfo()
        Aligner_info = Aligner.GetDeviceInfo()
                                                                #ensure successful connection of:
        print(f'Opening {Piezo_info.Description}, {serial_P}')  #Piezo
        print("----------------------------------------")
        print(f'Opening {Strain_info.Description}, {serial_S}')  #Strain Gauge
        print("----------------------------------------")
        print(f'Opening {Aligner_info.Description}, {serial_A}')  #Position Aligner
        print("----------------------------------------")

        # Start polling and enable for Piezo
        Piezo.StartPolling(250)  #250ms polling rate
        time.sleep(0.25)
        Piezo.EnableDevice()
        time.sleep(0.5)  # Wait for device to enable

        if not Piezo.IsSettingsInitialized():  # checks if piezo initializes
            Piezo.WaitForSettingsInitialized(10000)  # 10 second timeout
            assert Piezo.IsSettingsInitialized() is True

        # Start polling and enable for Position Aligner
        # 250ms polling rate
        Aligner.StartPolling(250)
        time.sleep(0.25)
        Aligner.EnableDevice()
        time.sleep(0.5)  # Wait for device to enable

        if not Aligner.IsSettingsInitialized():  # checks if Position Aligner initializes
             Aligner.WaitForSettingsInitialized(10000)
             assert Aligner.IsSettingsInitialized() is True

        # Start polling and enable for Strain Gauge
        # 250ms polling rate
        Strain.StartPolling(250)
        time.sleep(0.25)
        Strain.EnableDevice()
        time.sleep(0.5)  # Wait for device to enable

        if not Strain.IsSettingsInitialized():  # checks if Strain Gauge initializes
            Strain.WaitForSettingsInitialized(10000)
            assert Strain.IsSettingsInitialized() is True

        # Load the device configuration
        Piezo_config = Piezo.GetPiezoConfiguration(serial_P)
        Aligner_config = Aligner.GetPositionAlignerConfiguration(serial_A)
        Strain_config = Strain.GetStrainGaugeConfiguration(serial_S)

        print("Cubes Initialized")

        #initialization of data from K-cubes
        straindata = []
        Xdiff = []
        Ydiff = []
        Sum = []
        straindatarev = []
        Xdiffrev = []
        Ydiffrev = []
        Sumrev = []
        y = []
        i = 0
        XdiffFwd = "XDiff-Forward-Indentation.gsac"
        YdiffFwd = "YDiff-Forward-Indentation.gsac"
        SUMFwd = "Sum-Forward-Indentation.gsac"
        XdiffRev = "XDiff-Reverse-Indentation.gsac"
        YdiffRev = "YDiff-Reverse-Indentation.gsac"
        SUMRev = "Sum-Reverse-Indentation.gsac"

        window = tk.Tk()
        window.title("K-Cube Controller")
        window.minsize(400, 250)
        window.resizable(width=False, height=False)

        for i in range(3):
            window.columnconfigure(i, weight=1, minsize=125)
            window.rowconfigure(i, weight=1, minsize=75)
        for j in range(0, 3):
            frame = tk.Frame(master=window)
            frame.grid(row=i, column=j, padx=10, pady=10)

        def quitApp():
            # Stop Polling and Disconnect
            print("Closing the Program")
            Piezo.SetZero()
            Piezo.StopPolling()
            Piezo.Disconnect()
            Strain.StopPolling()
            Strain.Disconnect()
            Aligner.StopPolling()
            Aligner.Disconnect()
            window.destroy()
            exit()

        def Kcube():
            StartTime = time.time()
            Mats = Material.get()
            Velocity = radio.get()
            MaxAlignerSetpoint = Displacement.get()

            if Mats.strip() == "" and Velocity == 0 and MaxAlignerSetpoint.strip() == "":
                messagebox.showinfo("Error", "Please enter values into all three parameters")

            elif Mats.strip() != "" and Velocity != 0 and MaxAlignerSetpoint.strip() != "":
                Aligner.SetOperatingMode(PositionAlignerStatus.OperatingModes.Monitor, True)
                Xdiff.clear()
                Ydiff.clear()
                Sum.clear()
                straindata.clear()
                Xdiffrev.clear()
                Ydiffrev.clear()
                Sumrev.clear()
                straindatarev.clear()
                figure, axis = plt.subplots(3, 2, sharey='row')

                channels = ["Dev1/ai0", "Dev1/ai1", "Dev1/ai2", "Dev1/ai3"]
                # Initializing and Setting Inputs for DAQ
                with nidaqmx.Task() as task:
                    for channel in channels:
                        task.ai_channels.add_ai_voltage_chan(
                            channel)  # adds XDiff, YDiff, Sum, and Strain Gauge Monitor to the NI DAQ Initialization

                    increment = 0
                    y = []

                    # MAIN APPROACH LOOP

                    while True:
                        if (Velocity == 1):
                            increment += 0.1
                        elif (Velocity == 2):
                            increment += 0.05
                        else:
                            increment +=0.02
                        Piezo.SetOutputVoltage(Decimal(increment))
                        with nidaqmx.Task() as task:
                            for channel in channels:
                                task.ai_channels.add_ai_voltage_chan(
                                    channel)  # adds XDiff, YDiff, Sum, and Strain Gauge Monitor to the NI DAQ Initialization

                            data = task.read()
                        Xdiff.append(data[1])
                        Ydiff.append(data[0])
                        Sum.append(data[2])
                        straindata.append(2.146478922 * data[3])
                        y.append(abs(data[0]))
                        time.sleep(0.0001)
                        if abs(y[-1]) >= float(MaxAlignerSetpoint):
                            break
                        if increment >= 74:
                            break
                        window.update()

                    # MAIN RETRACTION LOOP
                    y = []
                    while True:
                        if (Velocity == 1):
                            increment -= 0.1
                        elif (Velocity == 2):
                            increment -= 0.05
                        else:
                            increment -= 0.02
                        Piezo.SetOutputVoltage(Decimal(increment))
                        with nidaqmx.Task() as task:
                            for channel in channels:
                                task.ai_channels.add_ai_voltage_chan(
                                    channel)  # adds XDiff, YDiff, Sum, and Strain Gauge Monitor to the NI DAQ Initialization

                            data = task.read()
                        Xdiffrev.append(data[1])
                        Ydiffrev.append(data[0])
                        Sumrev.append(data[2])
                        straindatarev.append(2.146478922 * data[3])
                        y.append(data[0])
                        time.sleep(0.0001)
                        if increment <= 0.1:
                            break
                        window.update()

                    EndTime = time.time() - StartTime
                    print(EndTime)
                    # plotting data into different plots
                    axis[0, 0].plot(straindata, Xdiff)
                    axis[0, 0].set_ylabel('Xdiff Voltage')
                    axis[0, 0].set_title('Indentation')
                    axis[1, 0].plot(straindata, Ydiff)
                    axis[1, 0].set_ylabel('Ydiff Voltage')
                    axis[2, 0].plot(straindata, Sum)
                    axis[2, 0].set_xlabel('Strain Gauge Displacement (um)')
                    axis[2, 0].set_ylabel('Sum Voltage')
                    axis[0, 1].plot(straindatarev[::-1], Xdiffrev[::-1])
                    axis[0, 1].invert_xaxis()
                    axis[0, 1].set_title('Retraction')
                    axis[1, 1].plot(straindatarev[::-1], Ydiffrev[::-1])
                    axis[1, 1].invert_xaxis()
                    axis[2, 1].plot(straindatarev[::-1], Sumrev[::-1])
                    axis[2, 1].invert_xaxis()
                    axis[2, 1].set_xlabel('Strain Gauge Displacement (um)')
                    plt.subplots_adjust(wspace=0)
                    plt.show()
            else:
                messagebox.showinfo("Error", "Please enter a values into the missing parameters ")

        def Setzero():
            time.sleep(5)
            Piezo.SetZero()
            Strain.SetZero()
            #time.sleep(14)

        def SaveTxt():
            Mats = Material.get()
            time = datetime.now().strftime("%H:%M:%S // %d-%m-%Y")

            with open(Mats + ", " + XdiffFwd, "w") as file:
                file.write(time + '\n')
                file.write("# Strain Data" + "\t" + "\t" + "\t" + "XDiff Data" + "\n")
                i = 0
                while i < len(straindata):
                    file.write(str(straindata[i]) + " " + str(Xdiff[i]) + "\n")
                    i += 1
            file.close()

            with open(Mats + ", " + YdiffFwd, "w") as file:
                file.write(time + '\n')
                file.write("# Strain Data" + "\t" + "\t" + "\t" + "YDiff Data" + "\n")
                i = 0
                while i < len(straindata):
                    file.write(str(straindata[i]) + " " + str(Ydiff[i]) + "\n")
                    i += 1
            file.close()

            with open(Mats + ", " + SUMFwd, "w") as file:
                file.write(time + '\n')
                file.write("# Strain Data" + "\t" + "\t" + "\t" + "Sum Data" + "\n")
                i = 0
                while i < len(straindata):
                    file.write(str(straindata[i]) + " " + str(Sum[i]) + "\n")
                    i += 1
            file.close()

            with open(Mats + ", " + XdiffRev, "w") as file:
                file.write(time + '\n')
                file.write("# Strain Data" + "\t" + "\t" + "\t" + "XDiff Data" + "\n")
                i = len(straindatarev) - 1
                while i >=0:
                    file.write(str(straindatarev[i]) + " " + str(Xdiffrev[i]) + "\n")
                    i -= 1
            file.close()

            with open(Mats + ", " + YdiffRev, "w") as file:
                file.write(time + '\n')
                file.write("# Strain Data" + "\t" + "\t" + "\t" + "YDiff Data" + "\n")
                i = len(straindatarev) - 1
                while i >=0:
                    file.write(str(straindatarev[i]) + " " + str(Ydiffrev[i]) + "\n")
                    i -= 1
            file.close()

            with open(Mats + ", " + SUMRev, "w") as file:
                file.write(time + '\n')
                file.write("# Strain Data" + "\t" + "\t" + "\t" + "Sum Data" + "\n")
                i = len(straindatarev) - 1
                while i >= 0:
                    file.write(str(straindatarev[i]) + " " + str(Sumrev[i]) + "\n")
                    i -= 1
            file.close()

            with open(Mats + ", Indentation.gsac", "w") as file:
                file.write(time + '\n')
                file.write("# Strain Data" + "\t" + "\t" + "\t" + "\t" + "XDiff Data"  + "\t" + "\t" + "\t" + "\t" + "YDiff Data"  + "\t" + "\t" + "\t" + "\t" + "Sum Data" + "\n")
                i = len(straindata) - 1
                while i >= 0:
                    file.write(str(straindata[i]) + "\t" + "\t" + str(Xdiff[i]) + "\t" + "\t" + str(Ydiff[i]) + "\t" + "\t" + str(Sum[i]) + "\n")
                    i -= 1
            file.close()

            with open(Mats + ", Retraction.gsac", "w") as file:
                file.write(time + '\n')
                file.write("# Strain Data" + "\t" + "\t" + "\t" + "\t" + "XDiff Data" + "\t" + "\t" + "\t" + "\t" + "YDiff Data"  + "\t" + "\t" + "\t" + "\t" + "Sum Data" + "\n")
                i = len(straindatarev) - 1
                while i >= 0:
                    file.write(str(straindatarev[i]) + "\t" + "\t" + str(Xdiffrev[i]) + "\t" + "\t" + str(Ydiffrev[i]) + "\t" + "\t" + str(Sumrev[i]) + "\n")
                    i -= 1
            file.close()

        '''
        def Conversion():
            i = 0
            ConversionTest = "CorrectedVoltage.gsac"
            ConversionFactor = "CorrectionFactor.txt"

            with open(ConversionFactor, "r") as file:
                incrementalconversion = [float(line.strip()) for line in file]
            file.close()

            while i < len(incrementalconversion) and i < len(straindata):
                ConversionData.append(straindata[int(i)] * incrementalconversion[int(i)])
                i += 1

            with open(ConversionTest, "w") as file:
                file.write("# Strain Data" + "\t" + "XDiff Data" + "\n")
                i = 0
                while i < len(ConversionData):
                    file.write(str(ConversionData[i]) + "\n")
                    i += 1
            file.close()
            ConversionData.clear()
        '''
        #Speed = tk.Entry(window, bd=6, width=15)
        #Speed.grid(column=0, row=0, padx=10, pady=10)
        #tk.Label(window, text=": Set Maximum Speed" + '\n' + "(0.0001-0.1)").grid(column=1, row=0, padx=10, pady=10)
        radio = tk.IntVar()

        VelocityFast = tk.Radiobutton(window, bd=5, text="0.1v/ increment", font=("Arial", 9), variable=radio, value=1)
        VelocityFast.grid(column=0, row=0, columnspan=2, rowspan=2, sticky='NW')

        VelocitySlow = tk.Radiobutton(window, bd=5, text="0.05v/ increment", font=("Arial", 9), variable=radio, value=2)
        VelocitySlow.grid(column=0, row=0, columnspan=2, rowspan=2, sticky='NE')

        VelocitySlower = tk.Radiobutton(window, bd=5, text="0.02v/ increment", font=("Arial", 9), variable=radio, value=3)
        VelocitySlower.grid(column=0, row=0, columnspan=2, sticky='S')

        Displacement = tk.Entry(window, bd=5, width=15)
        Displacement.grid(column=0, row=1, sticky='N')
        tk.Label(window, text=": Set Position Setpoint (0-1)").grid(column=1, row=1, sticky='N')

        Material = tk.Entry(window, bd=5, width=15)
        Material.grid(column=0, row=1, rowspan=2)
        tk.Label(window, text=": Material Test" + '\n' + "(File Name Save)").grid(column=1, row=1, rowspan=2)

        SubmitButton = tk.Button(window, bd=5, bg="#90EE90", text="Start", height=2, width=15, command=Kcube)
        SubmitButton.grid(column=2, row=0, padx=4, pady=4, sticky='N')

        Zero = tk.Button(window, bd=5, bg="#E0B0FF", text="Zero", height=1, width=15, command=Setzero)
        Zero.grid(column=2, row=0, rowspan=2, padx=4, pady=4)

        Save = tk.Button(window, bd=5, bg="#E0B0FF", text="Save to Txt File", height=1, width=15, command=SaveTxt)
        Save.grid(column=2, row=1, padx=4, pady=4)

        QuitButton = tk.Button(window, bd=5, bg="#ff0000", text="Quit", height=1, width=15, command=quitApp)
        QuitButton.grid(column=2, row=1, padx=4, pady=4, rowspan=2)

        path = "CTELogo.jpg"
        img = ImageTk.PhotoImage(Image.open(path))
        Logo = tk.Label(window, image=img)
        Logo.grid(column=0, row=2, columnspan=3, sticky='S')

        '''
        Convert = tk.Button(window, bd=5, bg="#abdbd9", text="V > nN", height=1, width=10, command=Conversion)
        Convert.grid(column=0, row=3, padx=10, pady=10)
        tk.Label(window, text=": Conversion of Voltage to Force").grid(column=1, row=3, padx=10, pady=10)
        '''
        window.mainloop()

    except Exception as e:
        print(e)

if __name__ == "__main__":
    main()