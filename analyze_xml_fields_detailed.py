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
        logging.FileHandler("xml_analysis_detailed.log"),
        logging.StreamHandler(sys.stdout)
    ]
)

def analyze_xml_structure(xml_file_path):
    try:
        logging.info(f"Analisando estrutura detalhada do arquivo: {xml_file_path}")
        
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
                
                # Analisa detalhes do serviço
                if "DeclaracaoPrestacaoServico" in inf_nfse:
                    decl = inf_nfse["DeclaracaoPrestacaoServico"]
                    if "InfDeclaracaoPrestacaoServico" in decl:
                        inf_decl = decl["InfDeclaracaoPrestacaoServico"]
                        
                        if "Servico" in inf_decl:
                            servico = inf_decl["Servico"]
                            logging.info("Campos detalhados em Servico:")
                            for key, value in servico.items():
                                if isinstance(value, dict):
                                    logging.info(f"  - {key}:")
                                    for subkey, subvalue in value.items():
                                        logging.info(f"    - {subkey}: {subvalue}")
                                else:
                                    logging.info(f"  - {key}: {value}")
                
                # Verifica se há itens de serviço
                if "DeclaracaoPrestacaoServico" in inf_nfse:
                    decl = inf_nfse["DeclaracaoPrestacaoServico"]
                    if "InfDeclaracaoPrestacaoServico" in decl:
                        inf_decl = decl["InfDeclaracaoPrestacaoServico"]
                        
                        if "Servico" in inf_decl:
                            servico = inf_decl["Servico"]
                            
                            # Verifica se há itens
                            if "ItensServico" in servico:
                                itens = servico["ItensServico"]
                                logging.info("Itens de Serviço encontrados:")
                                
                                if isinstance(itens, list):
                                    for i, item in enumerate(itens):
                                        logging.info(f"Item {i+1}:")
                                        for key, value in item.items():
                                            logging.info(f"  - {key}: {value}")
                                else:
                                    logging.info("Item único:")
                                    for key, value in itens.items():
                                        logging.info(f"  - {key}: {value}")
                            else:
                                logging.info("Não foram encontrados itens de serviço detalhados")
                                
                            # Verifica códigos de serviço
                            logging.info("Códigos de serviço encontrados:")
                            if "ItemListaServico" in servico:
                                logging.info(f"  - ItemListaServico: {servico['ItemListaServico']}")
                            if "CodigoCnae" in servico:
                                logging.info(f"  - CodigoCnae: {servico['CodigoCnae']}")
                            if "CodigoTributacaoMunicipio" in servico:
                                logging.info(f"  - CodigoTributacaoMunicipio: {servico['CodigoTributacaoMunicipio']}")
                
                # Salva a estrutura completa da primeira nota em um arquivo JSON para análise
                with open("estrutura_nota_detalhada.json", "w", encoding="utf-8") as json_file:
                    json.dump(inf_nfse, json_file, indent=2, ensure_ascii=False)
                logging.info("Estrutura detalhada da nota salva em 'estrutura_nota_detalhada.json'")
                
                # Verifica se há mais notas com estruturas diferentes
                if isinstance(comp_nfse, list) and len(comp_nfse) > 1:
                    # Verifica algumas notas adicionais para ver se há diferenças estruturais
                    sample_indices = [min(100, len(comp_nfse)-1), min(500, len(comp_nfse)-1), min(1000, len(comp_nfse)-1)]
                    for idx in sample_indices:
                        if idx > 0:  # Pula o primeiro que já analisamos
                            sample_note = comp_nfse[idx]
                            sample_inf_nfse = sample_note["Nfse"]["InfNfse"]
                            
                            # Verifica se há campos adicionais nesta nota
                            if "DeclaracaoPrestacaoServico" in sample_inf_nfse:
                                sample_decl = sample_inf_nfse["DeclaracaoPrestacaoServico"]
                                if "InfDeclaracaoPrestacaoServico" in sample_decl:
                                    sample_inf_decl = sample_decl["InfDeclaracaoPrestacaoServico"]
                                    
                                    if "Servico" in sample_inf_decl:
                                        sample_servico = sample_inf_decl["Servico"]
                                        
                                        # Verifica se há itens nesta nota
                                        if "ItensServico" in sample_servico and "ItensServico" not in servico:
                                            logging.info(f"Nota {idx+1} contém ItensServico que não estavam na primeira nota")
                                            
                                            # Salva esta nota para análise
                                            with open(f"estrutura_nota_{idx+1}.json", "w", encoding="utf-8") as json_file:
                                                json.dump(sample_inf_nfse, json_file, indent=2, ensure_ascii=False)
                                            logging.info(f"Estrutura da nota {idx+1} salva em 'estrutura_nota_{idx+1}.json'")
                
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
