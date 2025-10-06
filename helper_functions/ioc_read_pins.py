import re
import pandas as pd
import warnings
from helper_functions.write_excel_file import write_excel_file


def parse_pin_file(file_path):
    if not isinstance(file_path, str):
        raise TypeError("The file path must be a string.")
    if not file_path.endswith(".ioc"):
        raise ValueError("The file must have a .ioc extension.")
    # Additional info
    additional_info = {"McuName": "", "McuPackage": "", "McuCPN": ""}

    # Define column names
    columns = ["PinNumber", "PinName", "PinFunction", "Signal", "Label", "Mode"]

    # Create an empty DataFrame with the specified columns
    pin_data = pd.DataFrame(columns=columns)

    with open(file_path, "r") as file:
        lines = file.readlines()

    # First pass: basic pin info
    for line in lines:
        if line.startswith("Mcu.Name"):
            additional_info["McuName"] = (line.split("=")[1]).strip()
        elif line.startswith("Mcu.Package"):
            additional_info["McuPackage"] = (line.split("=")[1]).strip()
        elif line.startswith("Mcu.CPN"):
            additional_info["McuCPN"] = (line.split("=")[1]).strip()

        elif line.startswith("Mcu.Pin"):
            pin_match = re.match(r"Mcu\.Pin(\d+)=(\w+)", line)
            if pin_match:
                current_pin_number, current_pin_name = pin_match.groups()
                if current_pin_name[0] == "P":  # Only real pins
                    current_pin_number = int(current_pin_number)
                    new_data = {
                        "PinNumber": current_pin_number,
                        "PinName": current_pin_name,
                        "PinFunction": "",
                        "Signal": "",
                        "Label": "",
                        "Mode": "",
                    }
                    pin_data = pd.concat(
                        [pin_data, pd.DataFrame([new_data])], ignore_index=True
                    )

        # Handle special RCC pins
        elif any(
            key in line
            for key in ["RCC_OSC_IN", "RCC_OSC_OUT", "RCC_OSC32_IN", "RCC_OSC32_OUT"]
        ):
            pin_name = line.split("-")[0]
            signal = line.split("=")[-1].strip() if "=" in line else line.strip()
            new_data = {
                "PinNumber": "",
                "PinName": pin_name,
                "PinFunction": "",
                "Signal": signal,
                "Label": "",
                "Mode": "",
            }
            pin_data = pd.concat(
                [pin_data, pd.DataFrame([new_data])], ignore_index=True
            )

    # Second pass: capture Mode, Signal, Label, PinFunction
    for line in lines:
        for pin_name in pin_data["PinName"]:
            # Match with optional (Function)
            mode_match = re.match(
                r"{}(?:\(([^)]+)\))?\.Mode=(\S+)".format(pin_name), line
            )
            signal_match = re.match(
                r"{}(?:\(([^)]+)\))?\.Signal=(\S+)".format(pin_name), line
            )
            label_match = re.match(
                r"{}(?:\(([^)]+)\))?\.GPIO_Label=(\S+)".format(pin_name), line
            )

            if not (mode_match or signal_match or label_match):
                continue

            i = pin_data[pin_data["PinName"] == pin_name].index[0]

            if mode_match:
                pin_data.loc[i, "Mode"] = mode_match.group(2)
                if mode_match.group(1):
                    pin_data.loc[i, "PinFunction"] = mode_match.group(1)
            elif signal_match:
                pin_data.loc[i, "Signal"] = signal_match.group(2)
                if signal_match.group(1):
                    pin_data.loc[i, "PinFunction"] = signal_match.group(1)
            elif label_match:
                pin_data.loc[i, "Label"] = label_match.group(2)
                if label_match.group(1):
                    pin_data.loc[i, "PinFunction"] = label_match.group(1)

    # Combine RCC pins if needed
    pin_data = (
        pin_data.groupby("PinName")
        .agg(
            {
                "PinNumber": "first",
                "PinFunction": "first",
                "Label": "first",
                "Mode": "first",
                "Signal": "".join,
            }
        )
        .reset_index()
    )

    # Final formatting
    pin_data["Mode/Label"] = pin_data["Mode"] + pin_data["Label"]
    pin_data = pin_data[["PinNumber", "PinName", "PinFunction", "Signal", "Mode/Label"]]

    # Check duplicates for EXTI
    duplicates = pin_data[
        pin_data["Signal"].str.contains("GPXTI1", na=False)
        & pin_data.duplicated(subset="Signal", keep=False)
    ]
    duplicate_EXTI_error = False
    if not duplicates.empty:
        duplicate_EXTI_error = True
        warnings.warn("Warning: Duplicate values found in the 'Signal' column.")
        print(duplicates)

    return pin_data, additional_info, duplicate_EXTI_error


def parse_peripherals(file_path):
    """Parse peripheral configuration from the .ioc file."""
    columns = [
        "Peripheral",
        "Mode",
        "Channel",
        "BaudRatePrescaler",
        "CalculatedBaudRate",
    ]
    periph_data = pd.DataFrame(columns=columns)

    with open(file_path, "r") as file:
        lines = file.readlines()

    current = {}

    for line in lines:
        line = line.strip()
        if "=" not in line:
            continue

        # Match peripherals (allow dashes, spaces, backslashes in param)
        periph_match = re.match(
            r"^(UART\d*|USART\d*|I2C\d*|SPI\d*|TIM\d*|CAN\d*|I2S\d*|SDIO\d*|SDMMC\d*|USB\d*|RCC.*|SYS.*)\.([\w\-\\ ]+)=(.+)",
            line,
        )

        # Match shared peripherals (SH.*)
        sh_match = re.match(r"^(SH\.\S+)\.([\w\-\\ ]+)=(.+)", line)

        if periph_match:
            periph, param, value = periph_match.groups()
            value = value.strip()
            current.setdefault("Peripheral", periph)

            if param in [
                "Mode",
                "VirtualType",
                "VirtualMode",
                "VirtualMode-Asynchronous",
            ]:
                current["Mode"] = value
            elif param.startswith("Channel-"):
                current["Channel"] = value
            elif param == "BaudRatePrescaler":
                current["BaudRatePrescaler"] = value
            elif param in ["CalculateBaudRate", "CalculatedBaudRate"]:
                current["CalculatedBaudRate"] = value

        elif sh_match:
            periph, param, value = sh_match.groups()
            value = value.strip()

            if param.isdigit():
                channel = value.split(",")[0]  # TIM8_CH1
                base_periph = channel.split("_")[0]  # TIM8
                current["Peripheral"] = base_periph
                current["Channel"] = channel

        # Commit row if meaningful
        if "Peripheral" in current and any(
            k in current
            for k in ["Mode", "Channel", "BaudRatePrescaler", "CalculatedBaudRate"]
        ):
            new_row = {col: current.get(col, "") for col in columns}
            periph_data = pd.concat(
                [periph_data, pd.DataFrame([new_row])], ignore_index=True
            )
            current = {}

    # Deduplicate raw entries
    periph_data = periph_data.drop_duplicates()

    # Merge by Peripheral: concatenate multiple values into one cell separated by '-'
    def merge_values(series):
        vals = [v for v in series if v]  # remove blanks
        return "-".join(sorted(set(vals))) if vals else ""

    periph_data = periph_data.groupby("Peripheral", as_index=False).agg(
        {
            "Mode": merge_values,
            "Channel": merge_values,
            "BaudRatePrescaler": merge_values,
            "CalculatedBaudRate": merge_values,
        }
    )

    return periph_data


if __name__ == "__main__":
    file_path = "STM32H562VGT6_config.ioc"  # Replace with your file path

    # Parse pins and peripherals
    pin_data, additional_info, duplicate_EXTI_error = parse_pin_file(file_path)
    periph_data = parse_peripherals(file_path)

    # Build output file path
    excel_file_path = file_path[:-3] + "xlsx"

    # Write everything to Excel
    write_excel_file(
        excel_file_path,
        pin_data,
        additional_info,
        duplicate_EXTI_error,
        peripherals=periph_data,
    )
