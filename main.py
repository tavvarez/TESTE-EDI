import os
import xml.etree.ElementTree as ET
from openpyxl import Workbook

def parse_cte_xml(file_path):
    tree = ET.parse(file_path)
    root = tree.getroot()

    # ajuste caso o XML possua outros namespaces
    ns = {'ns': 'http://www.portalfiscal.inf.br/cte'}

    
    data = {}

    # mapeamento de chaves
    data['versao'] = root.attrib.get('versao')
    data['EmitCNPJ'] = root.findtext('.//ns:emit/ns:CNPJ', namespaces=ns)
    data['EmitxNome'] = root.findtext('.//ns:emit/ns:xNome', namespaces=ns)
    data['TomadorCNPJ'] = root.findtext('.//ns:rem/ns:CNPJ', namespaces=ns)
    data['TomadorxNome'] = root.findtext('.//ns:rem/ns:xNome', namespaces=ns)
    data['vTPrest'] = root.findtext('.//ns:vPrest/ns:vTPrest', namespaces=ns)
    data['vRec'] = root.findtext('.//ns:vPrest/ns:vRec', namespaces=ns)
    data['serie'] = root.findtext('.//ns:ide/ns:serie', namespaces=ns)
    data['nCT'] = root.findtext('.//ns:ide/ns:nCT', namespaces=ns)
    data['dhEmi'] = root.findtext('.//ns:ide/ns:dhEmi', namespaces=ns)
    data['chave'] = root.findtext('.//ns:infNFe/ns:chave', namespaces=ns)
    
    return data

 

def save_to_excel(data_list, output_file, xml_file):
    cte_data = parse_cte_xml(xml_file)
    cte_number = cte_data['nCT']
    wb = Workbook()
    ws = wb.active
    ws.title = f"CTE_{cte_number}"

    # adicionar colunas
    headers = list(data_list[0].keys())
    ws.append(headers)

    # persistir dados
    for data in data_list:
        row = [data.get(key, '') for key in headers]
        ws.append(row)

    # Salvar o arquivo exvel
    wb.save(output_file)
    print(f"Arquivo Excel salvo como '{output_file}'.")

def process_directory(directory_path, output_file):
    data_list = []
    xml_file = None
    
    # Buscar XML no diretório XML padrão
    for filename in os.listdir(directory_path):
        if filename.endswith('.xml'):
            file_path = os.path.join(directory_path, filename)
            if xml_file is None:
                xml_file = file_path
            try:
                data = parse_cte_xml(file_path)
                data_list.append(data)
            except Exception as e:
                print(f"Erro ao processar '{filename}': {e}")

    # Salvar os dados no Excel
    if data_list and xml_file:
        save_to_excel(data_list, output_file, xml_file)
    else:
        print("Nenhum dado extraído dos arquivos XML.")


directory_path = './XML'
output_file = f'CTe_Data.xlsx'
process_directory(directory_path, output_file)
