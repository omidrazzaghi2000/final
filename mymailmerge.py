
from mailmerge import MailMerge

def setWordData(data):

    with MailMerge('age_band1.docx') as document:
        #print(document.get_merge_fields())
        
        document.merge(
            arzyab_name     = str(data['arzyab_name']),
            exam_date       = str(data['exam_date']),
            first_name      = str(data['first_name']),
            last_name       = str(data['last_name']),
            birth_data      = str(data['birth_date']),
            year            = str(data['year']),
            month           = str(data['month']),
            telephone       = str(data['telephone']),
            height          = str(data['height']),
            weight          = str(data['weight']),
            hand            = str(data['hand']),
            foot            = str(data['foot']),
            rg1             = str(data['raw_grade_1']),
            rg2             = str(data['raw_grade_2']),
            rg3             = str(data['raw_grade_3']),
            rg4             = str(data['raw_grade_4']),
            rg5             = str(data['raw_grade_5']),
            rg6             = str(data['raw_grade_6']),
            rg7             = str(data['raw_grade_7']),
            rg8             = str(data['raw_grade_8']),
            rg9             = str(data['raw_grade_9']),
            rg10            = str(data['raw_grade_10']),
            g1              = str(data['grade_1']),
            g2              = str(data['grade_2']),
            g3              = str(data['grade_3']),
            g4              = str(data['grade_4']),
            g5              = str(data['grade_5']),
            g6              = str(data['grade_6']),
            g7              = str(data['grade_7']),
            g8              = str(data['grade_8']),
            g9              = str(data['grade_9']),
            g10             = str(data['grade_10']),
            md1             = str(data['MD1']),
            MD2             = str(data['MD2']),
            MD3             = str(data['MD3']),
            ac1             = str(data['A&C1']),
            ac2             = str(data['A&C2']),
            Bal1            = str(data['Bal1']),
            Bal2            = str(data['Bal2']),
            Bal3            = str(data['Bal3']),
        summation_of_total  = str(data['summation_of_total']),
        percentage_total    = str(data['percentage_total']),
        standard_grade_total= str(data['standard_grade_total']),


            )
        document.write('merge.docx')

