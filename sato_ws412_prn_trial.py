import win32print


def send_prn_to_printer(printer_name: str, prn_filepath: str) -> None:
    """Sends a raw PRN file directly to a specified printer on Windows."""
    try:
        # Open a connection to the printer
        h_printer = win32print.OpenPrinter(printer_name)

        try:
            # Start a print job in RAW mode
            # RAW tells Windows NOT to process or alter the content
            job_info = ("SATO PRN Test Job", None, "RAW")
            job_id = win32print.StartDocPrinter(h_printer, 1, job_info)

            win32print.StartPagePrinter(h_printer)

            # Read the bytes from the PRN file and write directly to the printer
            with open(prn_filepath, "rb") as prn_file:
                raw_data = prn_file.read()
                win32print.WritePrinter(h_printer, raw_data)

            win32print.EndPagePrinter(h_printer)
            win32print.EndDocPrinter(h_printer)

            print(
                f"Successfully sent '{prn_filepath}' to '{printer_name}' (Job ID: {job_id})"
            )

        finally:
            
            win32print.ClosePrinter(h_printer)

    except Exception as e:
        print(f"Error printing PRN file: {e}")


def list_installed_printers():
    """Lists all printer names installed on this computer."""
    printers = win32print.EnumPrinters(
        win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    )
    print("Available Printers:")
    for _, _, name, _ in printers:
        print(f"  - {name}")


if __name__ == "__main__":
    
    #list_installed_printers()

    
    PRINTER_NAME = "SATO WS412TT"  
    FILE_PATH = "Thvim_rack_naming_qr.prn"  

    send_prn_to_printer(PRINTER_NAME, FILE_PATH)
