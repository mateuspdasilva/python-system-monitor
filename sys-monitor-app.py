import wmi
import pynvml

# Global instance
wmi_instance = wmi.WMI()
pynvml.nvmlInit()

# CPU information

print(f"--------------------------")
print(f"- Processor")
print(f"--------------------------")
try:
    processor = wmi_instance.Win32_Processor()[0]

    try:
        processor_model = processor.Name
        print(f"Processor model: {processor_model}")
    except:
        print("Processor model not available.")

    try:
        processor_brand = processor.Manufacturer
        print(f"Processor brand: {processor_brand}")
    except:
        print("Processor brand not available.")

    try:
        processor_frequency = (processor.MaxClockSpeed / 1000)
        print(f"Processor max frequency: {processor_frequency:.1f}GHz")
    except:
        print("Processor max frequency not available.")

    try:
        processor_cores = processor.NumberOfCores
        print(f"Processor cores: {processor_cores} cores")
    except:
        print("Processor cores not available.")

    try:
        processor_threads = processor.NumberOfLogicalProcessors
        print(f"Processor threads: {processor_threads} threads")
    except:
        print("Processor threads not available.")

    try:
        processor_cache_memory = (processor.L2CacheSize + processor.L3CacheSize) / 1000
        print(f"Processor cache memory: {processor_cache_memory}MB")
    except:
        print("Processor cache memory not available.")

    try:
        processor_capacity = processor.DataWidth
        print(f"Processor Capacity: {processor_capacity} bits")
    except:
        print("Processor capacity not available.")

except:
    print("CPU information not available.")

# NVIDIA GPU information
print(f"--------------------------")
print(f"- Graphic Card")
print(f"--------------------------")
try:
    pynvml.nvmlInit()
    handle = pynvml.nvmlDeviceGetHandleByIndex(0)
    try:
        gpu_name = pynvml.nvmlDeviceGetName(handle).decode('utf-8')
        print(f"Graphic Card name: {gpu_name}")
    except:
        print("Could not get GPU name information.")

    try:
        memory_info = pynvml.nvmlDeviceGetMemoryInfo(handle)
        memory_size = memory_info.total / (1024 ** 3)
        print(f"Graphic Card memory size: {memory_size:.0f}GB")
    except:
        print("Could not get GPU memory size information.")

    try:
        memory_clock = pynvml.nvmlDeviceGetClockInfo(handle, pynvml.NVML_CLOCK_MEM)
        print(f"Graphic Card memory clock: {memory_clock}MHz")
    except:
        print("Could not get GPU memory clock information.")

    try:
        gpu_clock_speed = pynvml.nvmlDeviceGetClockInfo(handle, pynvml.NVML_CLOCK_GRAPHICS)
        print(f"Graphic Card base clock: {gpu_clock_speed}MHz")
    except:
        print("Could not get GPU base clock information.")

    try:
        gpu_boost_clock = pynvml.nvmlDeviceGetMaxClockInfo(handle, pynvml.NVML_CLOCK_GRAPHICS)
        print(f"Graphic Card boost clock: {gpu_boost_clock}MHz")
    except:
        print("Could not get GPU boost clock information.")

    try:
        pcie_width = pynvml.nvmlDeviceGetMaxPcieLinkWidth(handle)
        pcie_gen = pynvml.nvmlDeviceGetMaxPcieLinkGeneration(handle)
        print(f"Graphic Card Bus Interface: PCIe x{pcie_width} {pcie_gen}.0")
    except:
        print("Could not get GPU bus interface information.")

    try:
        driver_version = pynvml.nvmlSystemGetDriverVersion().decode('utf-8')
        print(f"Graphic Card driver version: {driver_version}")
    except:
        print("Could not get Graphic Card driver version information.")
    
    try:
        max_tdp = pynvml.nvmlDeviceGetPowerManagementLimit(handle) / 1000
        print(f"Max power consumption: {max_tdp:.0f}W")
    except:
        print("Could not get max power consumption information.")

    try:
        current_tdp = pynvml.nvmlDeviceGetPowerUsage(handle) / 1000
        print(f"Current power consumption: {current_tdp:.1f}W")
    except:
        print("Could not get current power consumption information.")

    try:
        gpu_temp = pynvml.nvmlDeviceGetTemperature(handle, pynvml.NVML_TEMPERATURE_GPU)
        print(f"Graphic Card temperature: {gpu_temp}°C")
    except:
        print("Could not get GPU temperature information.")

    try:
        fan_speed = pynvml.nvmlDeviceGetFanSpeed(handle)
        print(f"Graphic Card fan speed: {fan_speed}%")
    except:
        print("Could not get Fan Speed information.")

except:
    pynvml.nvmlShutdown()
    print("Could not get Graphic Card information.")
    

# RAM information
print(f"--------------------------")
print(f"- RAM")
print(f"--------------------------")
try: 
    memory_list = wmi_instance.Win32_PhysicalMemory()
    total_capacity = 0
    for ram in memory_list:
        try:
            ram_brand = ram.Manufacturer
            ram_capacity = int(ram.Capacity) / (1024 ** 3)
            ram_speed = ram.Speed
            print(f"Ram device:")
            print(f"---> Brand: {ram_brand}")
            print(f"---> Capacity: {ram_capacity:.0f}GB")
            print(f"---> Frequency: {ram_speed}MHz")
            total_capacity += ram_capacity
        except:
            print("Error while getting information for Ram device.")

    print(f"Total RAM capacity: {total_capacity:.0f}GB")
except:
    print("RAM information not available.")

# Storage information
print(f"--------------------------")
print(f"- Storage")
print(f"--------------------------")
try:
    for disk in wmi_instance.Win32_LogicalDisk():
        try:
            name = disk.Caption
            capacity = int(disk.Size) / (1024 ** 3)
            free_space = int(disk.FreeSpace) / (1024 ** 3)
            print(f"Storage device {name}")
            print(f"---> Capacity: {capacity:.2f}GB")
            print(f"---> Free Space: {free_space:.2f}GB")
        except:
            print(f"Error while getting information for storage device {name}")
except:
    print(f"Error while getting storage information.")

# BIOS version
print(f"--------------------------")
print(f"- BIOS")
print(f"--------------------------")
try:
    bios = wmi_instance.Win32_BIOS()[0]
    print(f"BIOS version: {bios.Version}")
except:
    print("BIOS version not available.")

print("Press enter to exit program...")
input()