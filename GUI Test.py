import csv
from os.path import join, abspath
from mailmerge import MailMerge

def writeOut(map, responsesFilePath, template, folder):
    with open(responsesFilePath, encoding='utf8', newline='') as auditionsFile:
        auditions = csv.DictReader(auditionsFile)

        next(auditions)
        for audition in auditions:
            document = MailMerge(template)
            for field in document.get_merge_fields():
                document.merge( field = audition[map[field]] )

            filename = str(audition[1]).strip()+"-"+str(audition[7]).strip()+".docx"
            filepath = abspath(join(folder, filename))

            document.write(filepath)
            document.close()

"""
            document.merge(
                auditon_number=audition[1],
                childs_name=audition[7],
                date_of_birth=datetime.strptime(audition[8][:15], "%a %b %d %Y").strftime("%m/%d/%Y"),
                age=audition[9],
                ballet_school=audition[10],
                training_years=audition[11],
                hrs_per_week=audition[12],
                training_level=audition[13],
                other_training=audition[14],
                danced_before=audition[16],
                previous_role_1=audition[17],
                previous_role_2=audition[18],
                previous_role_3=audition[19],
                other_leads=audition[15],
                height=audition[20],
                weight=audition[21],
                bust=audition[22],
                waist=audition[23],
                hips=audition[24],
                girth=audition[25],
                inseam=audition[26],
                guardian_name=audition[2],
                phone_number=audition[4],
                parents_email=audition[3],
                street_address=audition[6],
                conflict_dates=audition[27])
"""