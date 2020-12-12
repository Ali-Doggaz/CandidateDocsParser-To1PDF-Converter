import os
from resume_parser import resumeparse
from fpdf import FPDF
from PyPDF2 import PdfFileMerger, PdfFileReader


def find_cv():
    '''
    :param: N/A
    :return: The name of the file that is the most likely to be the candidate's resume/cv, as well as the candidate email
    adress, his phone number, and his name.
    '''

    #Special case where candidate's folder is empty or contains only 1 file

    if len(os.listdir('PDF_Converted_Files')) == 0:
        return
    if len(os.listdir('PDF_Converted_Files')) == 1:
        return os.listdir('PDF_Converted_Files')[0]


    # If a file contains the words CV/Resume/etc..., returns directly that file name
    for name in os.listdir('PDF_Converted_Files'):
        if name.lower().startswith("cv"):
            data = resumeparse.read_file('PDF_Converted_Files' + os.sep + name)
            return name, data['email'], data['phone'], data['name']

        for keyword in ['cv.', 'resume', 'résumé', 'curriculum vitae', 'curriculumvitae']:
            if keyword in name.lower():
                data = resumeparse.read_file('PDF_Converted_Files' + os.sep + name)
                return name, data['email'], data['phone'], data['name']

    # Attribute a score to each file. The score will allow us to estimate the probability of
    # that file being the candidate's resume.

    # "maxi" will store the highest score reached yet.
    maxi = 0

    # loop over all files and attribute a "score" to each one.
    # If the file's score is >= to maxi, maxi = score. In this case,
    # we will also store the email, phone number, and fullname present in the file.
    for name in os.listdir('PDF_Converted_Files'):
        # Parse the file, looking for relevant info (email, skills, education, etc...)
        score = 0
        data = resumeparse.read_file('PDF_Converted_Files' + os.sep + name)

        # Increase score if we find relevant info in the file
        if data['skills']:
            score += 5  # If the file contains the candidate's skills, there are very high chances that this
            # file is the candidate's resume. So we increase its score by 5.
        if data['email']:
            score += 1
        if data['phone']:
            score += 1
        if data['degree']:
            score += 3
        if data['university']:
            score += 2
        if data['total_exp']:
            score += 2
        if score >= maxi:
            cv_name, email, phone, FullName = name, data['email'], data['phone'], data['name']
            maxi = score

    # Return the file with the highest score (highest chances of being the candidate's resume)
    return cv_name, email, phone, FullName

def find_motivation_letter(cv_name):
    # If a file contains the word "motivation", returns directly that file name
    for name in os.listdir('PDF_Converted_Files'):
        if name == cv_name:  # The motivation letter can't be the same file as the resume.
            continue
        if "motivation" in name.lower():
            return name

    # Else, we return the second file submitted by the candidate, which probably is his motivation letter.

    if len(os.lisdir('PDF_Converted_Files')) < 2:
        return
    return os.listdir('PDF_Converted_Files')[1] if os.listdir('PDF_Converted_Files')[1] != cv_name else os.listdir('PDF_Converted_Files')[0]

def generate_number():
    '''
    :return:
    int ID: unique ID for the name of the new PDF
    '''
    if not len(os.listdir("Final")):
        ID = 100000
    else:
        # Gets ID of last file in the folder 'Final"
        str = os.listdir("Final")[-1]
        ID = int(str[len(str)-10:len(str)-4]) + 1

    return ID

def merge_pdfs(cv_name,motivation_letter_name, EMAIL, PHONE_NUMBER, FULLNAME):
    if not len(os.listdir("PDF_Converted_Files")):
        return
    file_name = "Candidat" + str(generate_number()) + ".pdf"
    output_file = "PDF_Converted_Files" + os.sep + file_name
    root = os.path.abspath(os.curdir)

    #Generate a PDF containing the first page, with the candidate name, phone number, and email
    pdf = FPDF()
    pdf.add_page()
    pdf.set_text_color(112, 100, 0)
    pdf.set_font('arial', 'B', 30)
    text = 'Nom: ' + FULLNAME + '\nAdresse Email: ' + EMAIL + '\nNumero de Telephone: ' + PHONE_NUMBER
    pdf.set_y(112)
    pdf.multi_cell(h=10, align='C', w=0, txt=text, border=0)
    pdf.output(output_file, 'F')

    #Merge the first page, the cv , and the motivation letter
    merger = PdfFileMerger()
    merger.append(PdfFileReader(open(root + os.sep + output_file, 'rb')))
    merger.append(PdfFileReader(open(root + os.sep + "PDF_Converted_Files" + os.sep + cv_name, 'rb')))
    if motivation_letter_name and len(motivation_letter_name) > 4:
        merger.append(PdfFileReader(open(root + os.sep + "PDF_Converted_Files" + os.sep + motivation_letter_name, 'rb')))

    #Merge the rest of the files
    for file in os.listdir("PDF_Converted_Files"):
        if file == cv_name or file == motivation_letter_name or file == file_name:
            continue
        if file.endswith(".pdf"):
            filepath = root + os.sep + "PDF_Converted_Files" + os.sep + file
            merger.append(PdfFileReader(open(filepath, 'rb')))

    merger.write(root + os.sep + "Final" + os.sep + file_name)

    # Returns the new PDF's name
    return file_name

def Create_Unique_PDF():
    '''
    Creates a unique PDF document for the candidate, by merging the files present in the 'Files' folder.
    It will also save the PDF name in a dictionnary, associating each PDF with the candidate's email and phone number.
    The dictionary is then stored in the Database.csv file.
    '''

    cv_name, EMAIL, PHONE_NUMBER, FULLNAME = find_cv()
    motivation_letter_name = find_motivation_letter(cv_name)
    file_name = merge_pdfs(cv_name, motivation_letter_name, EMAIL, PHONE_NUMBER, FULLNAME)

    #Add the new pdf to the database
    with open('Database.csv', mode='a+') as file:
        sep = ','
        file.write(file_name+sep+FULLNAME+sep+EMAIL+sep+PHONE_NUMBER)
        file.write("\n")

def clean():
    '''
    Remove all temp files in the 'PDF_Converted_Files' folder
    '''
    for file in os.listdir("PDF_Converted_Files"):
        os.remove("PDF_Converted_Files" + os.sep + file)

# TEST FUNCTIONS (TO DELETE)
# merge_pdfs("CV_AliDoggaz.pdf", "", 'test@gmail.com', '26337393', "ALI DOGGAZ")

