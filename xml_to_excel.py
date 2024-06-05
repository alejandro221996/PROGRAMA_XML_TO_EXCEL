import os
import xml.etree.ElementTree as ET
import pandas as pd

# Get the directory of the script
script_dir = os.path.dirname(__file__)

# Directory containing XML files (assuming it's in the same folder as the script)
xml_folder = os.path.join(script_dir, '')

# Excel file path
excel_file_path = os.path.join(script_dir, 'nomina_excel.xlsx')

# Function to process XML files
def process_xml_files():
    # Initialize an empty DataFrame
    df_all = pd.DataFrame()

    # Iterate through each XML file in the folder
    count = 0
    errors = 0
    for filename in os.listdir(xml_folder):
        if filename.endswith('.xml'):
            try:
                xml_path = os.path.join(xml_folder, filename)

                # Parse the XML file
                tree = ET.parse(xml_path)
                root = tree.getroot()

                # Extract information from Emisor and Receptor
                emisor = root.find('.//cfdi:Emisor', namespaces={'cfdi': 'http://www.sat.gob.mx/cfd/3'})
                nombre_emisor = emisor.get('Nombre')
                rfc_emisor = emisor.get('Rfc')

                receptor = root.find('.//cfdi:Receptor', namespaces={'cfdi': 'http://www.sat.gob.mx/cfd/3'})
                nombre_receptor = receptor.get('Nombre')
                rfc_receptor = receptor.get('Rfc')
                
                # Extract UUID
                timbre_fiscal_digital = root.find('.//tfd:TimbreFiscalDigital', namespaces={'tfd': 'http://www.sat.gob.mx/TimbreFiscalDigital'})
                uuid = timbre_fiscal_digital.get('UUID')
                
                #Fecha de emision
                fecha_emision = timbre_fiscal_digital.get('FechaTimbrado')

                #Nomina
                nomina = root.find('.//nomina12:Nomina', namespaces={'nomina12': 'http://www.sat.gob.mx/nomina12'})
                fecha_pago = nomina.get('FechaPago')

                # Extract TotalPercepciones
                total_percepciones = nomina.get('TotalPercepciones')
                
                #Total Sueldos
                total_sueldos = nomina.find('.//nomina12:Percepciones', namespaces={'nomina12': 'http://www.sat.gob.mx/nomina12'}).get('TotalSueldos')
                
                # Extract Concepto="PREMIOS DE ASISTENCIA"
                concepto_asistencia = nomina.find('.//nomina12:Percepciones/nomina12:Percepcion[@Concepto="PREMIOS DE ASISTENCIA"]', namespaces={'nomina12': 'http://www.sat.gob.mx/nomina12'})
                importe_asistencia = concepto_asistencia.get('ImporteGravado') if concepto_asistencia is not None else None
                
                # Extract Concepto="PREMIOS DE PUNTUALIDAD"
                concepto_puntualidad = nomina.find('.//nomina12:Percepciones/nomina12:Percepcion[@Concepto="PREMIOS DE PUNTUALIDAD"]', namespaces={'nomina12': 'http://www.sat.gob.mx/nomina12'})
                importe_puntualidad = concepto_puntualidad.get('ImporteGravado') if concepto_puntualidad is not None else None
                
                # Extract SubsidioCausado from OtroPago with Clave="D100" and Concepto="SUBSIDIO PARA EL EMPLEO"
                subsidio_causado = nomina.find('.//nomina12:OtrosPagos/nomina12:OtroPago[nomina12:SubsidioAlEmpleo]', namespaces={'nomina12': 'http://www.sat.gob.mx/nomina12'})
                subsidio_causado_value = subsidio_causado.find('.//nomina12:SubsidioAlEmpleo', namespaces={'nomina12': 'http://www.sat.gob.mx/nomina12'}).get('SubsidioCausado') if subsidio_causado is not None else None

                # Extract Importe from Concepto="IMSS" under Deducciones
                concepto_imss = nomina.find('.//nomina12:Deducciones/nomina12:Deduccion[@Concepto="IMSS"]', namespaces={'nomina12': 'http://www.sat.gob.mx/nomina12'})
                importe_imss = concepto_imss.get('Importe') if concepto_imss is not None else None
                
                # Extract Importe from Concepto="ISR" under Deducciones
                concepto_isr = nomina.find('.//nomina12:Deducciones/nomina12:Deduccion[@Concepto="ISR"]', namespaces={'nomina12': 'http://www.sat.gob.mx/nomina12'})
                importe_isr = concepto_isr.get('Importe') if concepto_isr is not None else None
                
                # Extract TotalDeducciones
                total_deducciones = nomina.get('TotalDeducciones') 
                
                # Extract information from Comprobante using the default namespace
                total_comprobante = root.get('Total')
                
                # Extract Importe from Concepto="CREDITO INFONAVIT" under Deducciones
                concepto_credito_infonavit = nomina.find('.//nomina12:Deducciones/nomina12:Deduccion[@Concepto="CREDITO INFONAVIT"]', namespaces={'nomina12': 'http://www.sat.gob.mx/nomina12'})
                importe_credito_infonavit = concepto_credito_infonavit.get('Importe') if concepto_credito_infonavit is not None else None
                
                # Create a DataFrame with the information
                data = {
                    'UUID': [uuid],
                    'RFC Receptor': [rfc_receptor],
                    'Nombre Receptor': [nombre_receptor],
                    'RFC Emisor': [rfc_emisor],
                    'Nombre Emisor': [nombre_emisor],
                    'Fecha de emision': [fecha_emision],
                    'Fecha de pago': [fecha_pago],
                    'Total Sueldos': [total_sueldos],
                    'Importe Asistencia': [importe_asistencia],
                    'Aguinaldo Gravado' : "",
                    'Aguinaldo Exento' : "",
                    'Importe Puntualidad': [importe_puntualidad],
                    'Subsidio Causado': [subsidio_causado_value],
                    'Total Percepciones': [total_percepciones],
                    'Importe IMSS': [importe_imss],
                    'Importe ISR': [importe_isr],
                    'Credito Infonavit': [importe_credito_infonavit],
                    'Total Deducciones': [total_deducciones],
                    'Neto Pagado': [total_comprobante]
                }

                df = pd.DataFrame(data)

                # Concatenate the current DataFrame with the overall DataFrame
                print(f'Procesando xml: {filename}')
                df_all = pd.concat([df_all, df], ignore_index=True)
                count += 1
            except Exception as e:
                print(str(e))
                print(f'Error procesando xml: {filename}')
                errors += 1
    
    # Save the DataFrame to an Excel file
    print("Archivos procesados correctamente: ", count)
    print("Archivos con errores: ", errors)
    df_all.to_excel(os.path.join(script_dir, 'nomina_excel.xlsx'), index=False)
    print('Archivo Excel creado con éxito')

# Ask user for input to start the script
input("Press Enter to start the script...")

# Call the function to process XML files
process_xml_files()

# Ask user for input to close the script
input("Press Enter to close the script...")