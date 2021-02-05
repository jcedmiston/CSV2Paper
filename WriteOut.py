import csv
from os.path import join, abspath
from mailmerge import MailMerge
from docx2pdf import convert

def write_out(map, responsesFilePath, template, folder):
    with open(responsesFilePath, encoding='utf8', newline='') as auditionsFile:
        auditions = csv.DictReader(auditionsFile)

        next(auditions)
        for audition in auditions:
            document = MailMerge(template)
            for field in document.get_merge_fields():
                document.merge( field = audition[map[field]] )

            filename = str(audition["Child's Name"]).strip()+".docx"
            filepath = abspath(join(folder, filename))

            convert(document, filepath)
            #document.write(filepath)
            document.close()