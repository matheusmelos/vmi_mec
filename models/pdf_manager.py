import fitz

class PDF: 
    def __init__(self, pdf_file, pdf_name):
        
        self.pdf_file = pdf_file
        self.pdf_name = pdf_name
        
        self.name = PDF.filter_name(pdf_name)
        self.size = PDF.type_page(self)
        self.revisao = PDF.search_revision_number(self)
        self.title = PDF.search_title(self)
        self.material, self.espessura = PDF.search_material_espessura(self)
        self.dobras = PDF.count_folds(self)
    
    def filter_name(pdf_name):
        return pdf_name[:-4]
     
    def type_page(self):
        
        doc = fitz.open(self.pdf_file)
        page = doc.load_page(0)
        text = page.get_text('text')
        large = page.rect.width
        
        if 'S   E   C   U   R   I   T   Y' in text:   
         
            if(large > 2000):
                size = 'A1_SECURITY'
            
            elif(1300 < large < 2000): 
                rect = fitz.Rect(1160, 995, 1660, 1030)
                texto = page.get_text("text", clip=rect)
                if "Material" in texto:
                    size = "A2_MATERIAL"
                elif "MATERIAL" in texto:
                    size = "A2.2"
                else:
                    size = 'A2_SECURITY'
            elif(870 <large < 1200):
                rect = fitz.Rect(680, 645, 1150, 680)   
                texto = page.get_text("text", clip=rect)
                rectfrase = fitz.Rect(840,680, 1120, 705) 
                frase = page.get_text("text", clip=rectfrase) 
                
                if "MATERIAL" in texto:  
                    size = 'A3.2_SECURITY'
                elif "Material" in texto:
                    if "Este Documento é de Propriedade da VMIS" in frase:
                         size = "A3_DESCRICAO"
                    else:
                            size = "A3_MATERIAL"
                else:
                    size = 'A3_SECURITY'
            
            elif(large < 840): 
                rect = fitz.Rect(65, 640, 500, 680)   
                texto = page.get_text("text", clip=rect)
                rectfrase = fitz.Rect(243, 690, 420, 708)
                frase = page.get_text("text", clip=rectfrase) 
            
               
                if "Material" in texto:
                    if "Este Documento é de Propriedade da VMIS" in frase:
                        
                        size = "A4_MATERIAL"
                    else:
                        size = 'A4'
                else:
                    size = "A4_RETRATO"
            
            else:
                size = 'A4_SECURITY'
        
        elif 'SISTEMAS DE SEGURANÇA' in text:
            
            if(large > 2000):
                size = 'A1_SISTEMAS'
            
            elif(1300 < large < 2000): 
                size = 'A2_SISTEMAS'
            
            elif(870 < large < 1200):
               size = 'A3_SISTEMAS'
            
            elif(845 < large < 1000):
                size = "A4_SISTEMAS"
            
            elif(large < 850):
                size = "A4.1_SISTEMAS"
        else:
            if(870 < large < 1200):
                rect = fitz.Rect(700, 645, 1050, 680)   
                texto = page.get_text("text", clip=rect)
               
                if "Material" in texto: 
                        size = 'A3_MATERIAL'
                else:
                        size = 'A3'
            else: 
                size = 'A4'
                
                
        doc.close()
        return size
    
    def search_revision_number(self):
        
        doc = fitz.open(self.pdf_file)
        page = doc.load_page(0)
        rect = None
        
        if self.size == 'A1_SECURITY':
            rect = fitz.Rect(2193, 1607, 2230, 1621)  # Revisão
        elif self.size == 'A2_SECURITY':
            rect = fitz.Rect(1500, 1123, 1540, 1140)  # Revisão
        elif self.size == 'A3_SECURITY':
            rect = fitz.Rect(1005, 774, 1048, 792)  # Revisão
        elif self.size == 'A3.2_SECURITY':
            rect = fitz.Rect(1008.4999389648438, 804, 1045, 820)  # Revisão
        elif self.size == 'A4_SECURITY':
            rect = fitz.Rect(760, 560, 795, 570)  # Revisão
        elif self.size == 'A4_SISTEMAS':
            rect = fitz.Rect(812, 548, 840, 555)  # Revisão
        elif self.size == 'A4.1_SISTEMAS':
            rect = fitz.Rect(778.1859130859375, 515, 805, 525)  # Revisão
        elif self.size == 'A3_MATERIAL':
            rect = fitz.Rect(1008.5, 775, 1045, 790) # Revisão
        elif self.size == 'A3':
            rect = fitz.Rect(1005, 774, 1048,  792)  # Revisão
        elif self.size == 'A4':
            rect = fitz.Rect(412, 775, 445, 789.1289672851562)
        elif self.size == "A4_RETRATO":
            rect =fitz.Rect(413, 775, 450, 790)
        elif self.size == "A4_MATERIAL":
            rect = fitz.Rect(413, 805, 450, 822)
        elif self.size == "A2_MATERIAL":
            rect = fitz.Rect(1500, 1123, 1540, 1140) #REVISAO
        elif self.size == "A2.2":
            rect = fitz.Rect(1500, 1155, 1540, 1170) 
        elif self.size =="A3_DESCRICAO":
            rect = fitz.Rect(1008.5, 805, 1045,820)
        text = page.get_text("text", clip=rect)
        revision_number = str(text.replace("\n", "").strip())
        
        doc.close()
        return revision_number

    def search_title(self):
        doc = fitz.open(self.pdf_file)
        page = doc.load_page(0)
        rect = None
        if self.size == 'A1_SECURITY':
            rect = fitz.Rect(2027.2437744140625, 1523, 2179, 1540)  # Denominação
        elif self.size == 'A2_SECURITY':
            rect = fitz.Rect(1334, 1038, 1530, 1055)  # Denominação
        elif self.size == 'A3_SECURITY':
            rect = fitz.Rect(841, 690, 1045, 708)  # Denominação
        elif self.size == 'A3.2_SECURITY':
            rect = fitz.Rect(825.3499755859375, 718, 1130, 735)  # Denominação
        elif self.size == 'A4_SECURITY':
            rect = fitz.Rect(471, 540, 615, 550)  # Denominação
        elif self.size == 'A4_SISTEMAS':
            rect = fitz.Rect(541, 585, 720, 595)  # Denominação
        elif self.size == 'A4.1_SISTEMAS':
            rect = fitz.Rect(528.177490234375, 550, 700, 560)  # Denominação
        elif self.size == "A3_MATERIAL":
           rect = fitz.Rect(840, 690, 1045, 705)
        elif self.size == "A3":
            rect = fitz.Rect(841, 690, 1045, 708)
        elif self.size == "A4":
            rect = fitz.Rect(243, 690, 420, 708)
        elif self.size == "A4_RETRATO":
            rect =fitz.Rect(247, 690, 440, 708)
        elif self.size == "A4_MATERIAL":
            rect = fitz.Rect(230, 720, 430, 735)
        elif self.size == "A2_MATERIAL":
            rect = fitz.Rect(1334, 1038, 1530, 1055) #DENOMINACAO
        elif self.size == "A2.2":
            rect = fitz.Rect(1320, 1067, 1530, 1083) 
        elif self.size =="A3_DESCRICAO":
            rect = fitz.Rect(825, 720, 1100, 735)
                      
                        
        text = page.get_text("text", clip=rect)
        title = str(text.replace("\n", " ").strip())
        
        doc.close()
        return title

    def search_material_espessura(self):
        doc = fitz.open(self.pdf_file)
        page = doc.load_page(0)
        
        rect_material = None
        rect_espessura = None
        espessura = None
        
        if self.size == 'A1_SECURITY':
            rect_material = fitz.Rect(2026.67724609375, 1639, 2160, 1652)  # Material
            rect_espessura = fitz.Rect(2198, 1638, 2228, 1655)  # Espessura
        elif self.size == 'A2_SECURITY':
            rect_material = fitz.Rect(1335, 1154, 1470, 1170)  # Material
            rect_espessura = fitz.Rect(1500, 1155, 1540, 1170)  # Espessura
        elif self.size == 'A3_SECURITY':
            rect_material = fitz.Rect(840, 805, 1000, 820)  # Material
            rect_espessura = fitz.Rect(1005, 805, 1048, 820)  # Espessura
        elif self.size == 'A3.2_SECURITY':
            rect_material = fitz.Rect(949, 645, 1063, 660)  # Material
        elif self.size == 'A4_SECURITY':
            rect_material = fitz.Rect(620, 540, 720, 550)  # Material
            rect_espessura = fitz.Rect(725, 540, 760, 550)  # Espessura
        elif self.size == 'A4_SISTEMAS':
            rect_material = fitz.Rect(542.8748168945312, 564, 613, 574)  # Material
        elif self.size == 'A4.1_SISTEMAS':
            rect_material = fitz.Rect(528, 530, 600, 540)  # Material
        elif self.size == "A3_MATERIAL":
            rect_material = fitz.Rect(790, 645, 880, 660)  # Material
            rect_espessura = fitz.Rect(723, 645, 790, 660)  # Espessura
        elif self.size == 'A3':
            rect_material = fitz.Rect(840, 805, 1000, 820)  # Material
            rect_espessura = fitz.Rect(1005, 805, 1048, 820)  # Espessura
        elif self.size == 'A4':    
            rect_material = fitz.Rect(195, 645, 290, 665)  # Material
            rect_espessura = fitz.Rect(130.0518798828125, 645.3804321289062, 191.1731414794922, 664)  # Espessura
        elif self.size == "A4_RETRATO":
            rect_material = fitz.Rect(246.5, 805, 400, 820) 
            rect_espessura = fitz.Rect(413.3974914550781, 805, 450, 820)
        elif self.size == "A4_MATERIAL":
            rect_material = fitz.Rect(370, 647, 460, 663)
        elif self.size == "A2_MATERIAL":
            rect_material = fitz.Rect(1280, 995, 1380, 1010) 
            rect_espessura = fitz.Rect(1230, 995,1270, 1010) 
        elif self.size == "A2.2":
            rect_material = fitz.Rect(1430, 995, 1580, 1010)
        elif self.size =="A3_DESCRICAO":
            rect_material = fitz.Rect(930, 645, 1045, 660)
                        
        text_material = page.get_text("text", clip=rect_material)
        material = str((text_material.strip()).replace("/", "-")).replace("\n", "")

        if rect_espessura:
            text_espessura = page.get_text("text", clip=rect_espessura)
            espessura = str(text_espessura.replace("mm", "").replace(",", ".").strip())
        else:
            if "mm" in material:
                espessura = ((material[-6:].replace("mm","")).replace(",",".")).replace("#", "")
            elif "MM" in material:
                espessura = ((material[-6:].replace("MM","")).replace(",",".")).replace("#", "")
            else:
                espessura = "Not found"
        doc.close()
        return material, espessura 
            
    def count_folds(self):
        
        doc = fitz.open(self.pdf_file)
        
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            
            text = page.get_text("text") 
            dobras_total = 0
            dobras_up = text.count("UP")
            dobras_down = text.count("DOWN")
            dobras_baixo = text.count("PARA BAIXO")
            dobras_cima = text.count("PARA CIMA")
               
            dobras_total += (dobras_baixo + dobras_cima + dobras_up + dobras_down)
            
        doc.close()
        
        return dobras_total

