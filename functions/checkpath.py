import os

def check_path(concat):
    if os.path.isdir(concat + "/" + "Pump-Station"):
        print("we in pump")
        pump = 1
    else:
        print("pump folder not found")
        pump = 0

    if os.path.isdir(concat + "/" + "Pressurized-Pipe"):
        print("we in pressure")
        pressure = 1
    else:
        print("pressure folder not found")
        pressure = 0

    if os.path.isdir(concat + "/" + "Wastewater"):
        print("we in gravity")
        gravity = 1
    else:
        print("gravity folder not found")
        gravity = 0

    if os.path.isdir(concat + "/" + "Excel"):
        print("we in Excel")
        excel = 1
    else:
        print("Excel not found")
        excel = 0

    return pump, pressure, gravity, excel
