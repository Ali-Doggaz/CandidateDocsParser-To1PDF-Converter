from PIL import Image
import os
from docx2pdf import convert
from fpdf import FPDF
import win32com.client
count = 0

def convert_image(File_Path):
    global count
    count += 1
    image1 = Image.open(File_Path)
    im1 = image1.convert('RGB')
    im1.save('PDF_Converted_Files'+os.sep+f'{count}.pdf')



wdFormatPDF=17
for name in os.listdir("Files"):
    if(name.endswith('.jpeg') or name.endswith('.jpg') or name.endswith('.png')):
        convert_image('Files'+os.sep+name)

    elif name.endswith('pdf'):
        os.chdir("PDF_Converted_Files")

    #elif name.endswith('.docx'):
        #count += 1
        #convert('Files'+os.sep+name, "PDF_Converted_Files"+os.sep+f"{count}.pdf")

    elif name.endswith(".docx") or name.endswith(".dotm") or name.endswith(".docm"):
        try:

            word = win32com.client.Dispatch('Word.Application')
            word.Visible = False
            doc = word.Documents.Open('Files' + os.sep + name)
            count += 1
            doc.SaveAs("PDF_Converted_Files" + os.sep + f"{count}.pdf", FileFormat=wdFormatPDF)
            doc.Close()
            word.Quit()
            word.Visible = True
        except:
            print("Could not open document " + 'Files' + os.sep + name)

    elif name.endswith('.doc') or name.endswith(".odt") or name.endswith(".rtf"):
        try:
            word = win32com.client.Dispatch('Word.Application')
            word.Visible = False
            doc = word.Documents.Open('Files' + os.sep + name)
            count += 1
            doc.SaveAs("PDF_Converted_Files" + os.sep + f"{count}.pdf", FileFormat=wdFormatPDF)
            doc.Close()
            word.Quit()
            word.Visible = True
        except:
            print("Could not open document " + 'Files' + os.sep + name)


    elif name.endswith(".txt"):
        try:
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=15)
            f = open("Files"+os.sep+name, "r")
            for x in f:
                pdf.cell(200, 10, txt=x, ln=1, align='C')
            # save the pdf with name .pdf
            count += 1
            pdf.output("PDF_Converted_Files"+os.sep+f"{count}.pdf")
        except:
            print("Could not open document " + 'Files' + os.sep + name)




