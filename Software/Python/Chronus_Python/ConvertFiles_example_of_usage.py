from Chronus_Python.ConvertFiles import *

# loaded_files = LoadFiles(r'D:\UnB\Projetos-Software\Chronus\Software\Blank_Test\teste91500_10042017\91500_TESTE',
#                          '.exp')

folder_address1 = r'D:\UnB\Projetos-Software\Chronus\Software\Blank_Test\DATA2\20131001.b'
folder_address2 = r'D:\UnB\Geocronologia\Cálculo U-Pb\Certificação do método\91500_BB9_PLESOVICE'

loaded_files = LoadFiles(
    folder_address2,
    '.exp',
    ',',
    'Thermo Finnigan Neptune_206MIC')


# a = loaded_files[0]
# print(a.GetCyclesDateTime())

# print()
# print('print_RawData')
# a.print_RawData()
#
# print()
# print('PrintSuportedEquipment')
# a.PrintSuportedEquipment()
#
# print()
# print('Print_Constants_Neptune')
# a.Print_Constants_Neptune()
#
# print()
# print('GetCyclesDateTime')
# print(a.GetCyclesDateTime())
#
# print()
# print('plot_save_All')
# a.plot_save_All(
#     SameFigure=True,
#     plot202=True,
#     plot204=True,
#     plot206=True,
#     plot207=True,
#     plot208=True,
#     plot232=True,
#     plot238=True,
#     showplot=True,
#     saveplots=False)

for file in loaded_files:
    file.split_blank_sample(print_separeted_files=False, new_split_points=[1, 11, 20, 60])
    # print('file.ID =', file.ID)
    # print(file.folderpath)

    # __Column_split_points = 4
    # __Line_new_num_cycles = 0
    # RawData = [[],[],[],[]]
    #
    # if __Column_split_points - 1 > len(RawData[__Line_new_num_cycles]):
    #     for n in range(__Column_split_points - len(RawData[__Line_new_num_cycles]) + 1):
    #         RawData[__Line_new_num_cycles].append('')
    #
    # print(RawData[__Line_new_num_cycles])
