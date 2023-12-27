import xml.etree.ElementTree as ET, pandas as pd, os, re, sys, argparse

def extract_data(element, path=""):
    """Extrai dados de elementos aninhados em um arquivo XML em um único dicionário."""
    data = {}
    if list(element):  # Se o elemento tiver filhos
        new_path = f"{path}/{element.tag}" if path else element.tag
        for child in element:
            child_data = extract_data(child, new_path)
            data.update(child_data)
    else:
        data_key = path + '/' + element.tag if path else element.tag
        data[data_key] = element.text
        # Verifica e extrai a propriedade ID
        if 'ID' in element.attrib:
            id_key = data_key + '/@ID'
            data[id_key] = element.attrib['ID']
    return data


def process_xml_to_df(diretorio, file_path):
    """Converte um arquivo XML em um DataFrame do pandas com uma única linha."""
    tree = ET.parse(os.path.join(diretorio, file_path))
    root = tree.getroot()

    # Removendo namespaces
    for elem in root.iter():
        if '}' in elem.tag:
            elem.tag = elem.tag.split('}', 1)[1]

    data = extract_data(root)
    return pd.DataFrame([data])

def group_files_by_number(xml_files):
    """Agrupa arquivos XML por número no final do nome do arquivo."""
    grouped_files = {}
    for file in xml_files:
        match = re.search(r'(\d+)\.xml$', file)
        if match:
            number = match.group(1)
            grouped_files.setdefault(number, []).append(file)
    return grouped_files

def main(diretorio, excel_file):

    # Listando e filtrando arquivos XML
    xml_files = [f for f in os.listdir(diretorio) if os.path.isfile(os.path.join(diretorio, f)) and f.endswith('.xml')]

    # Agrupando os arquivos por número
    grouped_xml_files = group_files_by_number(xml_files)

    # Convertendo e escrevendo no Excel
    with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
        for number, files in grouped_xml_files.items():
            dfs = [process_xml_to_df(diretorio, file) for file in files]
            combined_df = pd.concat(dfs, ignore_index=True)
            sheet_name = f'{number}'
            combined_df.to_excel(writer, sheet_name=sheet_name, index=False)


if __name__ == "__main__":

    parser = argparse.ArgumentParser(description='Converte arquivos XML em um arquivo Excel.')
    parser.add_argument('diretorio', type=str, help='Diretório contendo os arquivos XML.')
    parser.add_argument('excel_file', type=str, help='Nome do arquivo Excel de saída.')
    args = parser.parse_args()

    main(args.diretorio, args.excel_file)