import xmltodict
import os
import logging
import sys
import json

# Configuração de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("xml_analysis.log"),
        logging.StreamHandler(sys.stdout)
    ]
)

def analyze_xml_structure(xml_file_path):
    try:
        logging.info(f"Analisando estrutura do arquivo: {xml_file_path}")
        
        with open(xml_file_path, "rb") as xml_file:
            logging.info("Arquivo aberto, iniciando parse XML")
            xml_dict = xmltodict.parse(xml_file)
            logging.info("Parse XML concluído")
            
            # Verifica se é uma NFSe de Serviço Prestado
            if "ConsultarNfseServicoPrestadoResposta" in xml_dict:
                logging.info("Estrutura identificada: ConsultarNfseServicoPrestadoResposta")
                
                # Obtém a primeira nota para análise
                comp_nfse = xml_dict["ConsultarNfseServicoPrestadoResposta"]["ListaNfse"]["CompNfse"]
                
                if isinstance(comp_nfse, list) and len(comp_nfse) > 0:
                    first_note = comp_nfse[0]
                else:
                    first_note = comp_nfse
                
                # Analisa a estrutura da primeira nota
                inf_nfse = first_note["Nfse"]["InfNfse"]
                
                # Lista todos os campos disponíveis na nota
                logging.info("Campos disponíveis na nota:")
                for key in inf_nfse:
                    logging.info(f"- {key}")
                
                # Analisa detalhes específicos
                if "ValoresNfse" in inf_nfse:
                    logging.info("Campos em ValoresNfse:")
                    for key in inf_nfse["ValoresNfse"]:
                        logging.info(f"  - {key}")
                
                if "PrestadorServico" in inf_nfse:
                    logging.info("Campos em PrestadorServico:")
                    for key in inf_nfse["PrestadorServico"]:
                        logging.info(f"  - {key}")
                
                if "DeclaracaoPrestacaoServico" in inf_nfse:
                    decl = inf_nfse["DeclaracaoPrestacaoServico"]
                    if "InfDeclaracaoPrestacaoServico" in decl:
                        inf_decl = decl["InfDeclaracaoPrestacaoServico"]
                        logging.info("Campos em InfDeclaracaoPrestacaoServico:")
                        for key in inf_decl:
                            logging.info(f"  - {key}")
                        
                        if "Servico" in inf_decl:
                            logging.info("Campos em Servico:")
                            for key in inf_decl["Servico"]:
                                logging.info(f"    - {key}")
                        
                        if "Tomador" in inf_decl:
                            logging.info("Campos em Tomador:")
                            for key in inf_decl["Tomador"]:
                                logging.info(f"    - {key}")
                
                # Salva a estrutura completa da primeira nota em um arquivo JSON para análise
                with open("estrutura_nota.json", "w", encoding="utf-8") as json_file:
                    json.dump(inf_nfse, json_file, indent=2, ensure_ascii=False)
                logging.info("Estrutura da nota salva em 'estrutura_nota.json'")
                
            else:
                logging.warning("Estrutura não reconhecida como NFSe de Serviço Prestado")
                
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
            analyze_xml_structure(xml_path)
            break  # Analisa apenas o primeiro arquivo
