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


def main():
    if not is_admin():
        sys.exit("Must run as administrator!")

    toaster = ToastNotifier()
    monitor = init_librehardwaremonitor()

    samples = []

    for _ in range(N_SAMPLES):
        sensors = read_sensors(monitor)
        status = fan_status(sensors)

        samples.append(status)

        time.sleep(0.5)

    if not all(samples):
        toaster.show_toast(
            "Fan warning!",
            f"Erratic fan reading, restart PC to protect GPU",
            threaded=True,
            icon_path="fan.ico",
            duration=3,
        )

    while toaster.notification_active():
        time.sleep(0.1)


def init_librehardwaremonitor():
    clr.AddReference("librehardwaremonitor\LibreHardwareMonitorLib")
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
    fans = [sensor for sensor in sensors if sensor.SensorType == FAN_ID]

    return all(
        [
            fan.Value > FAN_MIN_THRESHOLD and fan.Value < FAN_MAX_THRESHOLD
            for fan in fans
        ]
    )


def is_admin():
    try:
        is_admin = os.getuid() == 0
    except AttributeError:
        is_admin = ctypes.windll.shell32.IsUserAnAdmin() != 0
    return is_admin


if __name__ == "__main__":
    sys.exit(main())
