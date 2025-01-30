import os

def check_path(concat):
    try:
        pump = 1 if os.path.isdir(os.path.join(concat, "Pump-Station")) else 0
        pressure = 1 if os.path.isdir(os.path.join(concat, "Pressurized-Pipe")) else 0
        gravity = 1 if os.path.isdir(os.path.join(concat, "Wastewater")) else 0
        excel = 1 if os.path.isdir(os.path.join(concat, "Excel")) else 0

        # Debugging prints
        print("Pump-Station:", "Found" if pump else "Not Found")
        print("Pressurized-Pipe:", "Found" if pressure else "Not Found")
        print("Wastewater:", "Found" if gravity else "Not Found")
        print("Excel:", "Found" if excel else "Not Found")

        return pump, pressure, gravity, excel

    except Exception as e:
        print(f"Error checking path: {concat}. Exception: {str(e)}")
        return 0, 0, 0, 0  # Default return in case of an error
