import wmi

c = wmi.WMI()

# CPU information
try:
    cpu = c.Win32_Processor()[0]
    print(f"CPU model: {cpu.Name}")
    print(f"CPU frequency: {cpu.CurrentClockSpeed} MHz")
    print(f"CPU cores: {cpu.NumberOfCores}")
    print(f"CPU threads: {cpu.NumberOfLogicalProcessors}")
except:
    print("CPU information not available.")

# NVIDIA GPU information
try:
    import pynvml
    pynvml.nvmlInit()
    handle = pynvml.nvmlDeviceGetHandleByIndex(0)
    name = pynvml.nvmlDeviceGetName(handle)
    gpu_temp = pynvml.nvmlDeviceGetTemperature(handle, pynvml.NVML_TEMPERATURE_GPU)
    gpu_clock_speed = pynvml.nvmlDeviceGetClockInfo(handle, pynvml.NVML_CLOCK_GRAPHICS)
    print(f"Graphic Card name: {name}")
    print(f"Graphic Card temperature: {gpu_temp} C")
    print(f"Graphic Card clock speed: {gpu_clock_speed} MHz")
    pynvml.nvmlShutdown()
except:
    print("NVIDIA GPU not found.")

# RAM information
try:
    total_capacity = sum(int(ram.Capacity) for ram in c.Win32_PhysicalMemory())
    total_capacity_gb = total_capacity / (1024 ** 3)
    print(f"Total RAM capacity: {total_capacity_gb:.2f} GB")
    print(f"RAM frequency: {c.Win32_PhysicalMemory()[0].Speed} MHz")
except:
    print("RAM information not available.")

# Storage information
try:
    for disk in c.Win32_LogicalDisk():
        name = disk.Caption
        capacity = int(disk.Size) / (1024 ** 3)
        free_space = int(disk.FreeSpace) / (1024 ** 3)
        print(f"Storage device {name}")
        print(f"---- Capacity: {capacity:.2f} GB")
        print(f"---- Free Space: {free_space:.2f} GB")
except:
    print("Storage information not available.")

# BIOS version
try:
    bios = c.Win32_BIOS()[0]
    print(f"BIOS version: {bios.Version}")
except:
    print("BIOS version not available.")

print("Press enter to exit program...")
input()