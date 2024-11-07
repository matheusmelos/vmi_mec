from models.pdf_manager import PDF
from models.dxf_manager import DXF
from models.dwg_manager import DWG
import traceback
import openpyxl
import rarfile
import zipfile
import shutil
import os

os.environ["PATH"] += os.pathsep + "/usr/bin"

# Processa o arquivo .zip de entrada e os 'transforma' em um Arquivo.zip e uma planilha
class ZipFolderManager:
    
    def __init__(self, zip_folder):
        
        self.pdfs = []
        self.dxfs = []
        self.dwgs = []
        
        self.folder = zip_folder
        
        self.extraction_folder = 'Extractedfiles'
        self.organization_folder = 'OrganizedFiles'
        self.organization_folder_zip = 'OrganizedFilesZip'
        self.all_pdfs = 'OrganizedFiles/PDFs'
        self.all_dwgs = 'OrganizedFiles/DWGs'
        self.all_dxfs = 'OrganizedFiles/DXFs'
    
        self.extract_folder()
        self.organize_folders()
        self.sheets = ZipFolderManager.create_sheet(self)
        self.processed_zip = ZipFolderManager.zip_file_process(self)
            
# Recebe a pasta .zip e realiza a primeira descompactação              
    def extract_folder(self): 
        try:
            # Remove a pasta de extração se ela já existir
            if os.path.exists(self.extraction_folder):
                shutil.rmtree(self.extraction_folder)
            
            # Cria a pasta de extração
            os.makedirs(self.extraction_folder)

            # Verifica o tipo de arquivo e realiza a extração
            if self.folder.lower().endswith('.zip'):
                with zipfile.ZipFile(self.folder, 'r') as archive:
                    archive.extractall(self.extraction_folder)
                    
            elif self.folder.lower().endswith('.rar'):
                with rarfile.RarFile(self.folder, 'r') as archive:
                    archive.extractall(self.extraction_folder)
            with open("relatorio.txt", "a") as f:
                f.write(f"Extraiu a pasta princinpal {self.folder}\n")
            # Chama a função para extrair subarquivos
            self.extract_subfiles(self.extraction_folder)
            
        except Exception as e:
            # Registra o erro no relatorio.txt
            with open("relatorio.txt", "a") as f:
                f.write(f"O arquivo {self.folder} não é válido ou está corrompido.\n")
                f.write(f"Erro ao processar o arquivo {self.folder}:\n")
                f.write(f"{str(e)}\n")
                f.write("Detalhes do rastreamento:\n")
                f.write(traceback.format_exc())
                f.write("\n\n")

# Armazena todas as pastas em uma pilha de processamento e ao encontrar um arquivo PDF DXF DWG os armazena criando classes       
    def extract_subfiles(self, root_folder):
        with open("relatorio.txt", "a") as f:
                f.write(f"Entrou na subfolder {self.folder}\n")
        pdf_processed = set()
        dxf_processed = set()
        dwg_processed = set()
        processed_files = set()
        
        stack = [root_folder]

        while stack:
            current_folder = stack.pop()

            for root, dirs, files in os.walk(current_folder):
                for file in files:
                    file_path = os.path.join(root, file)

                    if file_path in processed_files:
                        continue 

                    processed_files.add(file_path)

                    # Tratamento de erro para arquivos ZIP
                    if file.lower().endswith('.zip'):
                        sub_destino = os.path.join(root, os.path.splitext(file)[0])
                        os.makedirs(sub_destino, exist_ok=True)
                        try:
                            with zipfile.ZipFile(file_path, 'r') as sub_archive:
                                sub_archive.extractall(sub_destino)
                            os.remove(file_path)
                            stack.append(sub_destino)
                        except zipfile.BadZipFile as e:
                            with open("relatorio.txt", "a") as f:
                                f.write(f"O arquivo {file_path} não é um ZIP válido ou está corrompido.\n")
                                f.write(f"Erro ao processar o arquivo {file_path}:\n")
                                f.write(f"{str(e)}\n")
                                f.write("Detalhes do rastreamento:\n")
                                f.write(traceback.format_exc())
                                f.write("\n\n")

                    # Tratamento de erro para arquivos RAR
                    elif file.lower().endswith('.rar'):
                        sub_destino = os.path.join(root, os.path.splitext(file)[0])
                        os.makedirs(sub_destino, exist_ok=True)
                        try:
                            with rarfile.RarFile(file_path, 'r') as sub_archive:
                                sub_archive.extractall(sub_destino)
                            os.remove(file_path)
                            stack.append(sub_destino)
                        except rarfile.Error as e:
                            with open("relatorio.txt", "a") as f:
                                f.write(f"O arquivo {file_path} não é um RAR válido ou está corrompido.\n")
                                f.write(f"Erro ao processar o arquivo {file_path}:\n")
                                f.write(f"{str(e)}\n")
                                f.write("Detalhes do rastreamento:\n")
                                f.write(traceback.format_exc())
                                f.write("\n\n")

                    # Tratamento de erro para arquivos PDF
                    elif file.upper().endswith(".PDF"):
                        try:
                            if file_path not in pdf_processed:
                                pdf = PDF(file_path, os.path.basename(file))
                                self.pdfs.append(pdf)
                                pdf_processed.add(file_path)
                                with open("relatorio.txt", "a") as f:
                                    f.write(f"Extraiu {file_path}\n")
                        except Exception as e:
                            with open("relatorio.txt", "a") as f:
                                f.write(f"Erro ao processar o arquivo PDF {file_path}:\n")
                                f.write(f"{str(e)}\n")
                                f.write("Detalhes do rastreamento:\n")
                                f.write(traceback.format_exc())
                                f.write("\n\n")

                    # Tratamento de erro para arquivos DXF
                    elif file.upper().endswith(".DXF"):
                        try:
                            if file_path not in dxf_processed:
                                dxf = DXF(file_path, os.path.basename(file))
                                self.dxfs.append(dxf)
                                dxf_processed.add(file_path)
                                with open("relatorio.txt", "a") as f:
                                    f.write(f"Extraiu {file_path}\n")
                        except Exception as e:
                            with open("relatorio.txt", "a") as f:
                                f.write(f"Erro ao processar o arquivo DXF {file_path}:\n")
                                f.write(f"{str(e)}\n")
                                f.write("Detalhes do rastreamento:\n")
                                f.write(traceback.format_exc())
                                f.write("\n\n")

                    # Tratamento de erro para arquivos DWG
                    elif file.upper().endswith(".DWG"):
                        try:
                            if file_path not in dwg_processed:
                                dwg = DWG(file_path, os.path.basename(file))
                                self.dwgs.append(dwg)
                                dwg_processed.add(file_path)
                                with open("relatorio.txt", "a") as f:
                                    f.write(f"Extraiu {file_path}\n")
                        except Exception as e:
                            with open("relatorio.txt", "a") as f:
                                f.write(f"Erro ao processar o arquivo DWG {file_path}:\n")
                                f.write(f"{str(e)}\n")
                                f.write("Detalhes do rastreamento:\n")
                                f.write(traceback.format_exc())
                                f.write("\n\n")
       
# Cria pastas de acordo o material de cada projeto e os move de uma forma que todos os arquivos do mesmo tipo fiquem agrupados de acordo o seu material
    def organize_folders(self):
        with open("relatorio.txt", "a") as f:
                f.write(f"Entrou na organizer {self.folder}\n")
        def create_and_move_file(file_path, file_name, file_src):
            try:
                # Cria a pasta se não existir
                if not os.path.exists(file_path):
                    os.makedirs(file_path)
                    
                    # Dá permissão total (rwx) para todos os grupos (777)
                    os.chmod(file_path, 0o777)
                    
                destino_path = os.path.join(file_path, file_name)
                
                # Remove o arquivo de destino se ele já existir
                if os.path.exists(destino_path):
                    os.remove(destino_path)
                
                # Move o arquivo para o destino
                shutil.move(file_src, destino_path)

                # Dá permissão total (rwx) para todos os grupos (777) no arquivo movido
                os.chmod(destino_path, 0o777)
                
            except OSError as e:
                print(f"Erro ao criar ou mover o arquivo {file_name}: {e}")
            except Exception as e:
                print(f"Ocorreu um erro inesperado ao mover {file_name}: {e}")

        try:
           
            # Cria as pastas caso não existam
            for directory in [self.all_pdfs, self.all_dxfs, self.all_dwgs]:
                if os.path.exists(directory):
                    shutil.rmtree(directory)
                
                os.makedirs(directory)
                    
                # Dá permissão total (rwx) para todos os grupos (777)
                os.chmod(directory, 0o777)

            # Organiza os PDFs
            for pdf in self.pdfs:
                pdf_path = os.path.join(self.all_pdfs, pdf.material)
                create_and_move_file(pdf_path, pdf.pdf_name, pdf.pdf_file)
                
                # Associa o material dos PDFs aos DXFs e DWGs correspondentes
                for dxf in self.dxfs:
                    if pdf.name in dxf.dxf_name:
                        dxf.material = pdf.material
                
                for dwg in self.dwgs:
                    if pdf.name in dwg.dwg_name:
                        dwg.material = pdf.material

            # Organiza os DXFs
            for dxf in self.dxfs:
                dxf_path = os.path.join(self.all_dxfs, dxf.material)
                create_and_move_file(dxf_path, dxf.dxf_name, dxf.dxf_file)  

            # Organiza os DWGs
            for dwg in self.dwgs:
                dwg_path = os.path.join(self.all_dwgs, dwg.material)
                create_and_move_file(dwg_path, dwg.dwg_name, dwg.dwg_file)

        except OSError as e:
            print(f"Erro ao criar pastas ou mover arquivos: {e}")
        except Exception as e:
            print(f"Ocorreu um erro inesperado: {e}")

# Une PDF e DXF correspondentes, e salva em uma planilha os dados  
    def create_sheet(self):
        with open("relatorio.txt", "a") as f:
                f.write(f"Entrou na planilha {self.folder}\n")
        sheet_name = 'Arquivo.xlsx'
        exist_file = os.path.exists(sheet_name)
        
        if exist_file:
            wb = openpyxl.load_workbook(sheet_name)
            ws = wb.active
        else:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Dados"
            ws.append(["NÚMERO DA PEÇA", "REVISÃO", "TÍTULO", "MATERIAL", "ESPESSURA (mm)", 
                    "DOBRAS (TOTAL)", "TEMPO DE CORTE (s)"])

        pdf_files_used = set()

        rows_to_add = []  
        for dxf in self.dxfs:
            for pdf in self.pdfs:
                if pdf.pdf_file in pdf_files_used:
                    continue
                if pdf.name in dxf.dxf_name:
                    rows_to_add.append([pdf.name, pdf.revisao, pdf.title, pdf.material, pdf.espessura, pdf.dobras, dxf.cut_time])
                    pdf_files_used.add(pdf.pdf_file)
                    break

        for pdf in self.pdfs:
            if pdf.pdf_file not in pdf_files_used:
                rows_to_add.append([pdf.name, pdf.revisao, pdf.title, pdf.material, pdf.espessura, pdf.dobras, None])
                pdf_files_used.add(pdf.pdf_file)

        for row in rows_to_add:
            ws.append(row)

        wb.save(sheet_name)

        return sheet_name
 
# Realiza a compactação das pastas criadas                   
    def zip_file_process(self):
        with open("relatorio.txt", "a") as f:
                f.write(f"Criou o processed {self.folder}\n")
        all_zip = "ProcessedFiles.zip"
        relatorio = "relatorio.txt"
        pdfs_zip = "PDFs.zip"
        dxfs_zip = "DXFs.zip"
        dwgs_zip = "DWGs.zip"
    
        def zip_folder (base_file, zip_file):
      
            with zipfile.ZipFile(zip_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
                        for root, dirs, files in os.walk(base_file):
                            for file in files:
                                caminho = os.path.join(root, file)
                                zipf.write(caminho, os.path.relpath(caminho, base_file))
                
            return zip_file
                                        
        zip_pdfs = zip_folder(self.all_pdfs, pdfs_zip)
        zip_dxfs = zip_folder(self.all_dxfs, dxfs_zip)
        zip_dwgs = zip_folder(self.all_dwgs, dwgs_zip)
        
        if os.path.exists(self.organization_folder_zip):
            shutil.rmtree(self.organization_folder_zip)
        
        os.makedirs(self.organization_folder_zip) 
        
        shutil.move(zip_pdfs, self.organization_folder_zip)
        shutil.move(zip_dxfs, self.organization_folder_zip)
        shutil.move(zip_dwgs, self.organization_folder_zip)
        shutil.move(self.sheets, self.organization_folder_zip)
       
        if os.path.exists(relatorio):
            shutil.move(relatorio, self.organization_folder_zip)
            
        zip_file_processed = zip_folder(self.organization_folder_zip, all_zip)
        
        destino_path = os.path.join("uploads", zip_file_processed)
        
        if os.path.exists(destino_path):
            os.remove(destino_path)
        
        shutil.move(zip_file_processed, "uploads")

        return zip_file_processed

# Apaga pastas criadas durante o processo    
    def clean_all(self):
        with open("relatorio.txt", "a") as f:
                f.write(f"Limpou! {self.folder}\n")
        if os.path.exists(self.extraction_folder):
            shutil.rmtree(self.extraction_folder)
        if os.path.exists(self.organization_folder):
            shutil.rmtree(self.organization_folder)
        if os.path.exists(self.organization_folder_zip):
            shutil.rmtree(self.organization_folder_zip)
        if os.path.exists(self.folder):
            os.remove(self.folder)    
        