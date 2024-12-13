import os
import json
import time
import subprocess
import shutil
import logging
from logging.handlers import RotatingFileHandler

def load_config():
    with open("config.json", "r") as f:
        return json.load(f)
    
def setup_logging(config):
    log_folder = config["log_folder"]
    if not os.path.exists(log_folder):
        os.makedirs(log_folder)

    log_file = os.path.join(log_folder, "print_service.log")
    max_log_size = 5 * 1024 * 1024  # 5 MB
    backup_count = 5  # Keep last 5 log files

    handler = RotatingFileHandler(log_file, maxBytes=max_log_size, backupCount=backup_count)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    handler.setFormatter(formatter)

    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    logger.addHandler(handler)

config = load_config()
setup_logging(config)
semaphore_extension = config["semaphore_extension"]
remove_printed_file = config["remove_printed_file"]
remove_printed_folder = config["remove_printed_folder"]
documents_to_print    = config["documents_to_print"]
        
class PDFPrintHandler:
    def __init__(self, config):
        self.config = config

    def get_printer_name_from_semaphore(self, folder_path):
        semaphore_path = os.path.join(folder_path, f"printer{semaphore_extension}")

        if os.path.exists(semaphore_path):
            with open(semaphore_path, "r") as f:
                printer_name = f.read().strip()
                return printer_name

        return None

    def process_folder(self, folder_path, printer_name):

        files_to_print = sorted(os.listdir(folder_path))
        
        printed_documents = 0 
        
        for filename in files_to_print:
            pdf_path = os.path.join(folder_path, filename)

            if filename.endswith(".pdf") and printed_documents <= documents_to_print:
                semaphore_file = os.path.join(folder_path, f"{os.path.splitext(filename)[0]}{semaphore_extension}")
                if os.path.exists(semaphore_file):
                    if self.print_pdf(pdf_path, printer_name):
                        
                        logging.info(f"{filename} impresso na impressora {printer_name}")
                        
                        try:
                            os.remove(semaphore_file)
                            logging.info(f"Arquivo de semáforo {semaphore_file} removido.")
                        except OSError as e:
                            logging.error(f"Erro ao remover o arquivo de semáforo {semaphore_file}: {e}")

                        try:
                            if remove_printed_file:
                                logging.info(f"Arquivo {pdf_path} removido.")
                                os.remove(pdf_path)
                            logging.info(f"Arquivo {pdf_path} removido.")
                            printed_documents += 1
                        except OSError as e:
                            logging.error(f"Erro ao remover o arquivo  {semaphore_file}: {e}")

                    
        semaphore_file = os.path.join(folder_path, f"printer{semaphore_extension}")

        del_files = [f for f in os.listdir(folder_path) if f == 'deletar.del']

        if remove_printed_folder and len(del_files) == 1:
            
            logging.info(f"Pasta {folder_path} apta a ser removida após impressão.")

            for filename in os.listdir(folder_path):
                file_path = os.path.join(folder_path, filename)
                
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    while True:
                        try:
                            os.unlink(file_path)
                            break
                        except PermissionError:
                            logging.warning(f"Arquivo {file_path} está em uso. ")
                            break
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            
            while True:
                try:
                    shutil.rmtree(folder_path)
                    logging.info(f"Pasta {folder_path} removida após impressão.")
                    break
                except PermissionError:
                    logging.warning(f"Pasta {folder_path} está em uso.")
                    break

    def print_pdf(self, pdf_path, printer_name):
        sumatra_path = self.config["sumatra_path"]
        imprimir_comando = self.config["print_command"]
        imprimir_timeout = self.config["print_timeout"]

        command = f'"{sumatra_path}" {imprimir_comando} "{printer_name}" "{pdf_path}"'

        try:
            logging.info(f"{pdf_path} sendo impresso em {printer_name}")
            result = subprocess.run(command, shell=True, timeout=imprimir_timeout)
            result.check_returncode()
            check_command = f'wmic printjob where "name like \'%{printer_name}%\'" get jobstatus'
            while True:
                result = subprocess.run(check_command, shell=True, capture_output=True, text=True)
                if "Printing" not in result.stdout:
                    break
                logging.info(f"Aguardando a conclusão da impressão de {pdf_path} na impressora {printer_name}...")
                time.sleep(5)
            return True
        
        except subprocess.TimeoutExpired:
            logging.error(f"Impressão do arquivo {pdf_path} na impressora {printer_name} abortada após {imprimir_timeout} segundos.")
            return False
        except subprocess.CalledProcessError as e:
            logging.error(f"Erro ao imprimir o arquivo {pdf_path} na impressora {printer_name}: {e}")
            return False

    def check_and_process_folders(self, monitor_folder):
        for folder_name in os.listdir(monitor_folder):
            folder_path = os.path.join(monitor_folder, folder_name)
            
            try:
                if os.path.isdir(folder_path) and folder_name.isdigit():
                    printer_name = self.get_printer_name_from_semaphore(folder_path)
                    if printer_name:
                        logging.info(f"Processando a pasta {folder_path} para impressora {printer_name}")
                        self.process_folder(folder_path, printer_name)
                    else:
                        current_time = time.time()
                        if not hasattr(self, 'last_log_time') or current_time - self.last_log_time >= 600:
                            logging.warning(f"Nenhum arquivo de semáforo válido encontrado em {folder_path}.")
                            self.last_log_time = current_time
            except Exception as e:
                logging.error(f"Erro ao processar a pasta {folder_path}: {e}")

def main():
    config = load_config()
    monitor_folder = config["monitor_folder"]
    print("serviço de impressão - versão 1.2")

    handler = PDFPrintHandler(config)

    logging.info(f"Monitorando a pasta {monitor_folder}...")
    try:
        while True:
            def get_all_printers():
                try:
                    result = subprocess.run('wmic printer get name', shell=True, capture_output=True, text=True)
                    printers = result.stdout.strip().split('\n')[1:]
                    with open("printers.txt", "w") as f:
                        for printer in printers:
                            f.write(printer.strip() + "\n")
                except subprocess.CalledProcessError as e:
                    logging.error(f"Erro ao obter a lista de impressoras: {e}")

            get_all_printers()
            
            def get_printing_jobs():
                try:
                    result = subprocess.run('wmic printjob get Document,Name,JobStatus', shell=True, capture_output=True, text=True)
                    jobs = result.stdout.strip().split('\n')[1:]
                    with open("printing_files.txt", "w") as f:
                        for job in jobs:
                            f.write(job.strip() + "\n")
                except subprocess.CalledProcessError as e:
                    logging.error(f"Erro ao obter a lista de documentos em impressão: {e}")

            get_printing_jobs()
            
            handler.check_and_process_folders(monitor_folder)

            time.sleep(5)
            
    except KeyboardInterrupt:
        logging.info("Monitoramento encerrado.")

if __name__ == "__main__":
    main()
