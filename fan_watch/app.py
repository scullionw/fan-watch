from win10toast import ToastNotifier
import sys
import time
import clr
import ctypes
import os

N_SAMPLES = 3
FAN_ID = 5
FAN_MIN_THRESHOLD = 400
FAN_MAX_THRESHOLD = 2500
FAN_COUNT = 3
SAMPLE_DELAY = 0.5


def main():
    if not is_admin():
        sys.exit("Must run as administrator!")

    monitor = init_librehardwaremonitor()

    samples = []
    for _ in range(N_SAMPLES):
        sensors = read_sensors(monitor)
        status = fan_status(sensors)
        samples.append(status)
        time.sleep(SAMPLE_DELAY)

    if not all(samples):
        alert("Erratic fan reading, restart PC to protect GPU")
        os.system("shutdown -t 0 -f")
    else:
        alert("Fans ok!")


def alert(message):
    ToastNotifier().show_toast(
        "Fan status", message, threaded=False, icon_path="fan.ico", duration=2,
    )


def init_librehardwaremonitor():
    clr.AddReference(r"librehardwaremonitor\LibreHardwareMonitorLib")
    from LibreHardwareMonitor import Hardware

    handle = Hardware.Computer()
    handle.IsMotherboardEnabled = True
    handle.Open()

    return handle


def read_sensors(monitor):
    sensorlist = []

    for i in monitor.Hardware:
        i.Update()
        for sensor in i.Sensors:
            sensorlist.append(sensor)
        for j in i.SubHardware:
            j.Update()
            for subsensor in j.Sensors:
                sensorlist.append(subsensor)

    return sensorlist


def fan_status(sensors):
    fan_speeds = [sensor.Value for sensor in sensors if sensor.SensorType == FAN_ID]

    print(fan_speeds)

    return len(fan_speeds) == FAN_COUNT and all(
        [rpm > FAN_MIN_THRESHOLD and rpm < FAN_MAX_THRESHOLD for rpm in fan_speeds]
    )


def is_admin():
    try:
        is_admin = os.getuid() == 0
    except AttributeError:
        is_admin = ctypes.windll.shell32.IsUserAnAdmin() != 0
    return is_admin


if __name__ == "__main__":
    sys.exit(main())
