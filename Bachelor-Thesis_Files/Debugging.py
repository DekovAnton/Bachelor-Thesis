import pyvisa

resource_manager= pyvisa.ResourceManager()
Serial_COM_ports=resource_manager.list_resources()
Instrument=resource_manager.open_resource("Name of instrument that will be used for data acquistion")
Instrument.query("MEASure:VOLTage:DC?")
