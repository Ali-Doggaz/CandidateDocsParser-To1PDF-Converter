from PIL import Image
import os
from fpdf import FPDF
import win32com.client

ROOT = os.path.abspath(os.curdir)
WDFORMATPDF = 17
PPTTOPDF = 32


def convert_image_to_pdf(File_Path, name):
    try:
        image1 = Image.open(File_Path)
        im1 = image1.convert('RGB')
        im1.save('PDF_Converted_Files' + os.sep + f'{name}.pdf')
    except:
        print("Could not open document " + File_Path)


def convert_doc_to_pdf(name):
    """
    
    :param name: (string) Name of the file to convert. (i.e: test.doc, Resume_Mohamed.docx) 
                
    Converts docx, doc, dotm, docm, odt, and rtf files to readable pdf. The resulting PDF will have the same name 
    as the original file. Moreover, it will be stored in the "PDF_Converted_Files' folder.
    """
    
    try:

        in_file = ROOT + os.sep + "Files" + os.sep + name
        word = win32com.client.Dispatch('Word.Application')
        word.Visible = 0  # ADD IN CASE DOCUMENT OPENS EVERY TIME
        doc = word.Documents.Open(in_file)
        doc.SaveAs(ROOT + os.sep + "PDF_Converted_Files" + os.sep + f"{name}.pdf", WDFORMATPDF)
        doc.Close()
        word.Quit()
        
    except:
        print("Could not open document " + 'Files' + os.sep + name)

def convert_txt_to_pdf(name):
    """
    
    :param name: (string) Name of the file to convert. (i.e: test.txt, Resume_Mohamed.txt) 

    Converts .txt files to readable pdf. The resulting PDF will have the same name 
    as the original file. Moreover, it will be stored in the "PDF_Converted_Files' folder.
    """
    
    try:
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=15)
        f = open("Files" + os.sep + name, "r")
        for x in f:
            pdf.cell(200, 10, txt=x, ln=1, align='C')
        pdf.output("PDF_Converted_Files" + os.sep + f"{name}.pdf")
    except:
        print("Could not open document " + 'Files' + os.sep + name)

def convert_powerpoint_to_pdf(name):
    """

    :param name: (string) Name of the file to convert. (i.e: Motivation.ppt, Resume_Mohamed.pptx) 

    Converts .ppt and .pptx files to readable pdf. The resulting PDF will have the same name 
    as the original file. Moreover, it will be stored in the "PDF_Converted_Files' folder.
    """
    
    try:
        in_file = ROOT + os.sep + "Files" + os.sep + name
        powerpoint = win32com.client.Dispatch("Powerpoint.Application")
        deck = powerpoint.Presentations.Open(in_file, WithWindow=False)
        deck.SaveAs(ROOT + os.sep + "PDF_Converted_Files" + os.sep + f"{name}.pdf",
                    PPTTOPDF)  # formatType = 32 for ppt to pdf
        deck.Close()
        powerpoint.Quit()
    except:
        print("Could not open document " + 'Files' + os.sep + name)
        
def Convert_Files_To_PDFs():
    """
    Converts all files in the folder "Files" to PDF.
    Compatible file types: jpg,jpeg,png,docx,doc,dotm,docm,odt,rtf,txt,pdf  

    Remark: Incompatible file types will be ignored, and an error message with the file's name 
    will be displayed in the console.
    """
    
    for name in os.listdir("Files"):

        if (name.endswith('.jpeg') or name.endswith('.jpg') or name.endswith('.png')):
            convert_image_to_pdf('Files' + os.sep + name, name)

        elif name.endswith('pdf'):
            in_file = ROOT + os.sep + "Files" + os.sep + name
            os.rename(in_file, "PDF_Converted_Files/" + os.sep + f"{name}.pdf")

        elif name.endswith(".docx") or name.endswith(".dotm") or name.endswith(".docm") or name.endswith('.doc') or \
                name.endswith(".odt") or name.endswith(".rtf"):
            convert_doc_to_pdf(name)

        elif name.endswith(".txt"):
            convert_txt_to_pdf(name)

        elif name.endswith(".ppt") or name.endswith(".pptx"):
            convert_powerpoint_to_pdf(name)

        else:
            print("[ERROR] File format incompatible: " + 'Files' + os.sep + name)


