from PIL import Image
import os
from fpdf import FPDF
import win32com.client



def convert_image(File_Path,name):
    try:
        image1 = Image.open(File_Path)
        im1 = image1.convert('RGB')
        im1.save('PDF_Converted_Files' + os.sep + f'{name}.pdf')
    except:
        print("Could not open document " + File_Path)


def Convert_Files_To_PDFs():
    root = os.path.abspath(os.curdir)
    print(root)
    wdFormatPDF = 17
    ppttoPDF = 32
    global count

    for name in os.listdir("Files"):
        '''
        Converts all files in the folder "Files" to PDF.
        Compatible file types: jpg,jpeg,png,docx,doc,dotm,docm,odt,rtf,txt,pdf  
        '''

        if (name.endswith('.jpeg') or name.endswith('.jpg') or name.endswith('.png')):
            convert_image('Files' + os.sep + name, name)

        elif name.endswith('pdf'):
            os.rename(in_file, "PDF_Converted_Files/" + os.sep + f"{name}.pdf")

        elif name.endswith(".docx") or name.endswith(".dotm") or name.endswith(".docm") or name.endswith('.doc') or \
                name.endswith(".odt") or name.endswith(".rtf"):
            try:
                in_file = root + os.sep + "Files" + os.sep + name
                word = win32com.client.Dispatch('Word.Application')
                word.Visible = 0 #ADD IN CASE DOCUMENT OPENS EVERY TIME
                doc = word.Documents.Open(in_file)
                doc.SaveAs(root + os.sep + "PDF_Converted_Files" + os.sep + f"{name}.pdf", wdFormatPDF)
                doc.Close()
                word.Quit()
            except:
                print("Could not open document " + 'Files' + os.sep + name)

        elif name.endswith(".txt"):
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

        elif name.endswith(".ppt") or name.endswith(".pptx"):
            try:
                in_file = root + os.sep + "Files" + os.sep + name
                powerpoint = win32com.client.Dispatch("Powerpoint.Application")
                deck = powerpoint.Presentations.Open(in_file, WithWindow=False)
                deck.SaveAs(root + os.sep + "PDF_Converted_Files" + os.sep + f"{name}.pdf", ppttoPDF)  # formatType = 32 for ppt to pdf
                deck.Close()
                powerpoint.Quit()
            except:
                print("Could not open document " + 'Files' + os.sep + name)

        else:
            print("[ERROR] File format incompatible: " + 'Files' + os.sep + name)


