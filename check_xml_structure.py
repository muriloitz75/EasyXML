import xmltodict
import os
import logging
import sys

# Configuração de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("xml_check.log"),
        logging.StreamHandler(sys.stdout)
    ]
)

def check_xml_structure(xml_file_path):
    try:
        logging.info(f"Analisando estrutura do arquivo: {xml_file_path}")
        
        with open(xml_file_path, "rb") as xml_file:
            logging.info("Arquivo aberto, iniciando parse XML")
            xml_dict = xmltodict.parse(xml_file)
            logging.info("Parse XML concluído")
            
            # Imprime as chaves de primeiro nível
            root_keys = list(xml_dict.keys())
            logging.info(f"Chaves de primeiro nível: {root_keys}")
            
            # Verifica a estrutura específica
            if "ConsultarNfseServicoPrestadoResposta" in xml_dict:
                logging.info("Estrutura identificada: ConsultarNfseServicoPrestadoResposta")
                
                # Verifica se há ListaNfse
                if "ListaNfse" in xml_dict["ConsultarNfseServicoPrestadoResposta"]:
                    logging.info("Contém ListaNfse")
                    
                    # Verifica se há CompNfse
                    if "CompNfse" in xml_dict["ConsultarNfseServicoPrestadoResposta"]["ListaNfse"]:
                        comp_nfse = xml_dict["ConsultarNfseServicoPrestadoResposta"]["ListaNfse"]["CompNfse"]
                        
                        # Verifica se CompNfse é uma lista ou um único item
                        if isinstance(comp_nfse, list):
                            logging.info(f"CompNfse é uma lista com {len(comp_nfse)} itens")
                            
                            # Analisa o primeiro item da lista
                            first_item = comp_nfse[0]
                            if "Nfse" in first_item:
                                logging.info("Primeiro item contém Nfse")
                                if "InfNfse" in first_item["Nfse"]:
                                    logging.info("Primeiro item contém InfNfse")
                                    
                                    # Mostra algumas informações da primeira nota
                                    inf_nfse = first_item["Nfse"]["InfNfse"]
                                    numero = inf_nfse.get("Numero", "Não encontrado")
                                    logging.info(f"Número da primeira nota: {numero}")
                                    
                                    # Verifica se há PrestadorServico
                                    if "PrestadorServico" in inf_nfse:
                                        prestador = inf_nfse["PrestadorServico"]
                                        razao_social = prestador.get("RazaoSocial", "Não encontrado")
                                        logging.info(f"Razão social do prestador: {razao_social}")
                        else:
                            logging.info("CompNfse é um único item")
                            # Analisa o item único
                            if "Nfse" in comp_nfse:
                                logging.info("Item contém Nfse")
                                if "InfNfse" in comp_nfse["Nfse"]:
                                    logging.info("Item contém InfNfse")
            elif "NFe" in xml_dict:
                logging.info("Estrutura identificada: NFe")
            elif "nfeProc" in xml_dict:
                logging.info("Estrutura identificada: nfeProc")
            else:
                logging.warning("Estrutura não reconhecida")
                
    except Exception as e:
        logging.error(f"Erro ao analisar o arquivo: {str(e)}")

if __name__ == "__main__":
    xml_dir = "nfs"
    
    if not os.path.exists(xml_dir):
        logging.error(f"Diretório {xml_dir} não encontrado")
        sys.exit(1)
        
    xml_files = os.listdir(xml_dir)
    
    if not xml_files:
        logging.warning(f"Nenhum arquivo encontrado no diretório {xml_dir}")
        sys.exit(0)
        
    for xml_file in xml_files:
        if xml_file.lower().endswith('.xml'):
            xml_path = os.path.join(xml_dir, xml_file)
            check_xml_structure(xml_path)
