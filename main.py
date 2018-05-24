import re
from bs4 import BeautifulSoup
from excel import Excel
from openpyxl.styles import PatternFill

TEST_STEP = (
'Maximum_RMS_Power',
'Testpattern_A',
'Testpattern_B',
'Testpattern_C',
'Testpattern_D',
'Testpattern_E',
'Testpattern_F',
'Testpattern_G',
'Testpattern_H',
'Carrier_Frequency_Error',
'Minimum_RMS_Power',
'ACLR_Channel__minus_2__RMS_',
'ACLR_Channel__minus_1__RMS_',
'ACLR_Channel__plus_1__RMS_',
'ACLR_Channel__plus_2__RMS_',
'Emission_Mask',
'Occupied_Bandwidth',
'Error_Vector_Magnitude__RMS_',
'Phase_Discontinuity', 
)

excel_location = {
'Maximum_RMS_Power':(12, 6),
'Minimum_RMS_Power':(13, 6),
'Transmit On/Off time mask ':(14, 6),
'Emission_Mask0':(15, 6),
'Emission_Mask1':(16, 6),
'Emission_Mask2':(17, 6),
'Emission_Mask3':(18, 6),
'ACLR1':(19, 6),
'ACLR2':(20, 6),
'Occupied_Bandwidth':(21, 6),
'Error_Vector_Magnitude__RMS_':(22, 6),
'Carrier_Frequency_Error':(24, 6),
'ILPC':(26, 6),
'PD':(27, 7)

}


BAND = {
'Operating Band I':'WCDMAB1',
'Operating Band III':'WCDMAB3',
'Operating Band VIII':'WCDMAB8'
}



excel = Excel('./a.xlsx')
result_list = []
for file in ['./wcdma_b1.XML', './wcdma_b3.XML', './wcdma_b8.XML']:
    file_content = open(file, 'r')
    soup = BeautifulSoup(file_content, 'xml')
    blocks = soup.find_all('Block')
    band = BAND[blocks[0].line.description.text.split(',')[0]]
    print(band)
    ILPCs = []
    PDs = []
    ACLRs = []
    result_dic = {}
    result_dic['ILPC'] = []
    result_dic['PD'] = []
    result_dic['ACLR1'] = []
    result_dic['ACLR2'] = []

    for block in blocks:
        lines = block.find_all('line')
        test_steps = block.find_all('test_step')
        for test_step in test_steps:

            description = test_step.description.text
            if description in TEST_STEP:

                if re.match('Testpattern',description):
                    ILPCs.append(test_step.Status.text)
                    break
                elif re.match('Phase_Discontinuity',description):
                    print('ok')
                    PDs.append(test_step.Status.text)
                    break
                elif re.match('ACLR', description):
                    ACLRs.append(test_step.Measurement_Value.text)
                    if description == 'ACLR_Channel__plus_2__RMS_':
                        ACLR1 = min(ACLRs[1], ACLRs[2])
                        ACLR2 = min(ACLRs[0], ACLRs[3])
                        result_dic['ACLR1'].append(ACLR1)
                        result_dic['ACLR2'].append(ACLR2)
                        ACLRs = []
                    continue
                elif re.match('Emission_Mask', description):
                    for n in range(4):
                        try:
                            result_dic[description+str(n)].append(test_step.Status.text)
                        except KeyError:
                            result_dic[description+str(n)] = []
                            result_dic[description+str(n)].append(test_step.Status.text)

                else:
                    try:
                        result_dic[description].append(test_step.Measurement_Value.text)
                    except KeyError:
                        result_dic[description] = []
                        result_dic[description].append(test_step.Measurement_Value.text)
        if description == 'Testpattern_H':
            if 'Failed' in ILPCs:
                result_dic['ILPC'].append('Failed')
            else:
                result_dic['ILPC'].append('Pass')

            Testpatterns = []
    print(PDs)
    if 'Failed' in PDs:
        result_dic['PD'].append('Failed')   
    else:
        result_dic['PD'].append('Pass')
    print(result_dic)


    for key in result_dic:
        n = 0
        #print(key)
        for value in result_dic[key]:
            result_list.append([band, excel_location[key][0], excel_location[key][1]+n, value])
            n = n + 1
print('正在写入excel')
NOK = 1
while NOK:
    try:
        excel.writeto(result_list)
        NOK = 0
    except PermissionError:
        input('请关闭excel文件后,Enter重试')
input('写入完成，按回车键退出')
