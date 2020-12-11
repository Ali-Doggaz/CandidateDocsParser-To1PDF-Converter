#from pyresparser import ResumeParser

from resume_parser import resumeparse



if __name__ == '__main__':
    #data = ResumeParser(r'CV_AliDoggaz.pdf').get_extracted_data()
    data = resumeparse.read_file('CV_AliDoggaz.pdf')
    print(data)
