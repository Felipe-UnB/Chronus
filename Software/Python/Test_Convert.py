import ConvertFiles

a = ConvertFiles.ListFiles(r'D:\UnB\Projetos-Software\Chronus\Software\Blank_Test\teste91500_10042017\91500_TESTE',
                           '.exp')
b = ConvertFiles.Convert_Files_to_Lists(a)
c = ConvertFiles.SampleData(b[0][0], b[0][1], 'Thermo Finnigan Neptune', False)

# print(c.Name)
#
# c.print_RawData()

c.plot_save_All(
    SameFigure=True,
    plot202=True,
    plot204=True,
    plot206=True,
    plot207=True,
    plot208=True,
    plot232=True,
    plot238=True)

# print('c plot')
# c.plot_optimized_parameters()
# for n in c.GetCyclesDateTime():
#     print(str(n))
# c.CUMSUM()

# a=detect_cusum.detect_cusum(c.Signal_207,10000,4000,ending=True)
# print(a)
# print(type(a))
# print(type(a[0]))
# print(a[1])
# print(a[2])
# print(a[3])
