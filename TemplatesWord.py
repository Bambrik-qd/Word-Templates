#from __future__ import print_function
from mailmerge import MailMerge
from datetime import datetime

def _TemplatesWord(file_data_structure):

    #template = "b5476133-371e-4779-ae99-6f6ea2d7e7df (1).docx"
    key = 'template'
    #template = get('template', "")
    if key in file_data_structure:
        template = file_data_structure[key]
        #print(user)
    else:
        print("Не указано имя фафла шаблона ")

    document = MailMerge(template)
    #print(document.get_merge_fields()) #- вывод всех полей-параметров
    #merge_fields = document.get_merge_fields()

    DocumentNumber = file_data_structure.get('DocumentNumber', "")
    DateDoc = file_data_structure.get('DateDoc', "")
    RequisiteRqCompanyName = file_data_structure.get('RequisiteRqCompanyName', "")
    RequisiteUfCrm = file_data_structure.get('RequisiteUfCrm', "")
    RequisiteRqDirector = file_data_structure.get('RequisiteRqDirector', "")
    CompanyEmailWork = file_data_structure.get('CompanyEmailWork', "")
    RequisiteRqEdrpou = file_data_structure.get('RequisiteRqEdrpou', "")
    RequisiteRegisteredAddressPostalCode = file_data_structure.get('RequisiteRegisteredAddressPostalCode', "")
    RequisiteRegisteredAddressCity2 = file_data_structure.get('RequisiteRegisteredAddressCity2', "")
    RequisiteRegisteredAddressAddress1 = file_data_structure.get('RequisiteRegisteredAddressAddress1', "")
    RequisiteRegisteredAddressAddress2 = file_data_structure.get('RequisiteRegisteredAddressAddress2', "")
    RequisitePrimaryAddressPostalCode = file_data_structure.get('RequisitePrimaryAddressPostalCode', "")
    RequisitePrimaryAddressCity = file_data_structure.get('RequisitePrimaryAddressCity', "")
    RequisitePrimaryAddressAddress1 = file_data_structure.get('RequisitePrimaryAddressAddress1', "")
    RequisitePrimaryAddressAddress2 = file_data_structure.get('RequisitePrimaryAddressAddress2', "")

    #DateDoc = datetime.now()
    month = DateDoc.strftime("%m")
    month_list = ['січня', 'лютого', 'березня', 'квітня', 'травня', 'червня ', 'липня', 'серпня', 'вересня', 'жовтня', 'листопада', 'грудня']
    #LAST_NAME = "Директор"
    #NAME = "Имя"
    #SECOND_NAME = "Отчество"
    #RequisiteRqDirector = LAST_NAME + " " + NAME[0] + ". " + SECOND_NAME[0] + "." #из оду уже с параметрами выводить в нужном формате
    #CompanyEmailWork


    document.merge(
        DocumentNumber = DocumentNumber,
        DocumentCreateTime = DateDoc.strftime('"%d" ') + month_list[int(month) - 1] + DateDoc.strftime(' %Y'),
        RequisiteRqCompanyName = RequisiteRqCompanyName,
        RequisiteUfCrm = RequisiteUfCrm,   
        RequisiteRqDirector = RequisiteRqDirector,
        CompanyEmailWork = CompanyEmailWork, 
        RequisiteRqEdrpou = RequisiteRqEdrpou,
        RequisiteRegisteredAddressPostalCode = RequisiteRegisteredAddressPostalCode,
        RequisiteRegisteredAddressCity2 = RequisiteRegisteredAddressCity2,
        RequisiteRegisteredAddressAddress1 = RequisiteRegisteredAddressAddress1,
        RequisiteRegisteredAddressAddress2 = RequisiteRegisteredAddressAddress2,
        RequisitePrimaryAddressPostalCode = RequisitePrimaryAddressPostalCode,
        RequisitePrimaryAddressCity = RequisitePrimaryAddressCity,
        RequisitePrimaryAddressAddress1 = RequisitePrimaryAddressAddress1,
        RequisitePrimaryAddressAddress2 = RequisitePrimaryAddressAddress2
         )

    document.write('test-output.docx')
    
#template = "b5476133-371e-4779-ae99-6f6ea2d7e7df (1).docx"

template_values = {'template':"b5476133-371e-4779-ae99-6f6ea2d7e7df (1).docx",
                  'DocumentNumber':'Номер документа',
                  'DateDoc': datetime(2023, 3, 23),
                  'RequisiteRqCompanyName':"Имя клиента",
                  'RequisiteUfCrm':"номер записи",
                  'RequisiteRqDirector': "Фамилия И.О.",
                  'CompanyEmailWork':"Почта@ukr.net",
                  'RequisiteRqEdrpou': "0123456789",
                  'RequisiteRegisteredAddressPostalCode':"01234",
                  'RequisiteRegisteredAddressCity2':"Город",
                  'RequisiteRegisteredAddressAddress1':"Адрес",
                  'RequisiteRegisteredAddressAddress2':"",           
                  'RequisitePrimaryAddressPostalCode':"01234",
                  'RequisitePrimaryAddressCity':"Город",                  
                  'RequisitePrimaryAddressAddress1':"Адрес",
                  'RequisitePrimaryAddressAddress2':""
                   }
_TemplatesWord(template_values)