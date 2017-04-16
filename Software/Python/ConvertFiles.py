from __future__ import division, print_function
import detect_cusum


class configuration:
    '''
    Some constants necessary to handle data in the files from some mass spectrometers
    '''

    def __init__(self):
        self.MassSpectometersSuported = [
            'Thermo Finnigan Neptune'
        ]

        self.__Constants_Neptune = {
            'Line_IsotopeHeaders': 15,
            'Line_Cups': 101,
            'Line_AnalysisDate': 8,
            'Line_FirstCycle': 15,
            'Line_LastCycle': 100,
            'Column_CycleTime': 1,
            'Column_206': 2,
            'Column_208': 3,
            'Column_232': 4,
            'Column_238': 5,
            'Column_202': 6,
            'Column_204': 7,
            'Column_207': 8,
            'NumberCycles': 85,
            'Columns_num': 13
        }

    def PrintSuportedEquipment(self):
        for equip in self.MassSpectometersSuported:
            print(equip)

    def Print_Constants_Neptune(self):
        '''Prints all constants in the dictionary'''
        for key in self.__Constants_Neptune:
            print(str(key) + ' : ' + str(self.__Constants_Neptune[key]))

    def GetConfigurations(self, equipment):
        if equipment == 'Thermo Finnigan Neptune':
            return self.__Constants_Neptune


config = configuration()


class SampleData():
    '''
    This object represents a single data file. Now, only the 206 intensities were translated to a list
    '''

    import datetime

    def __init__(
            self,
            Name,
            File_as_Array,
            Equipment,
            unit_202='cps',
            unit_204='cps',
            unit_206='mV',
            unit_207='cps',
            unit_208='cps',
            unit_232='mV',
            unit_238='mV'
    ):

        import datetime

        self.__Constants = config.GetConfigurations(Equipment)
        self.Name = Name
        self.__RawData = File_as_Array

        AnalysisDate = File_as_Array[self.__Constants['Line_AnalysisDate']][0]
        self.AnalysisDate = datetime.datetime.strptime(AnalysisDate.split(' ')[2], "%d/%m/%Y")
        FirstCycleTime = File_as_Array[self.__Constants['Line_FirstCycle']][self.__Constants['Column_CycleTime']]
        self.FirstCycleTime = datetime.datetime.strptime(FirstCycleTime, '%H:%M:%S:%f')

        # self.__DateTimeAnalysis = datetime.datetime.combine(AnalysisDate,FirstCycleTime.time())
        # print(self.__DateTimeAnalysis)

        for line in self.__RawData:
            if len(line) != self.__Constants['Columns_num']:
                print('Class SampleData')
                print(self.Name)
                print('Lines with different number of columns.')
                exit()

        Line_FirstIntensity = self.__Constants['Line_IsotopeHeaders']
        Line_LastIntensity = Line_FirstIntensity + self.__Constants['NumberCycles']

        self.__CyclesDateTime = []
        self.__Signal_206 = []
        self.__Signal_208 = []
        self.__Signal_232 = []
        self.__Signal_238 = []
        self.__Signal_202 = []
        self.__Signal_204 = []
        self.__Signal_207 = []

        __Column_CycleTime = self.__Constants['Column_CycleTime']
        __Column_206 = self.__Constants['Column_206']
        __Column_208 = self.__Constants['Column_208']
        __Column_232 = self.__Constants['Column_232']
        __Column_238 = self.__Constants['Column_238']
        __Column_202 = self.__Constants['Column_202']
        __Column_204 = self.__Constants['Column_204']
        __Column_207 = self.__Constants['Column_207']

        for line in range(Line_FirstIntensity, Line_LastIntensity):

            CyclesTime = datetime.datetime.strptime(File_as_Array[line][__Column_CycleTime], '%H:%M:%S:%f')
            self.__CyclesDateTime.append(datetime.datetime.combine(self.AnalysisDate.date(), CyclesTime.time()))
            # print(self.__CyclesDateTime)
            try:
                self.__Signal_206.append(float(File_as_Array[line][__Column_206]))
                self.__Signal_208.append(float(File_as_Array[line][__Column_208]))
                self.__Signal_232.append(float(File_as_Array[line][__Column_232]))
                self.__Signal_238.append(float(File_as_Array[line][__Column_238]))
                self.__Signal_202.append(float(File_as_Array[line][__Column_202]))
                self.__Signal_204.append(float(File_as_Array[line][__Column_204]))
                self.__Signal_207.append(float(File_as_Array[line][__Column_207]))

            except:
                print('Class SampleData')
                print(self.Name)
                print('Error while converting str to float.')
                print('File_as_Array[line][__Column_206] = ', 'File_as_Array[', line, '][', __Column_206, '] = ',
                      File_as_Array[line][__Column_206])
                print('File_as_Array[line][__Column_208] = ', 'File_as_Array[', line, '][', __Column_208, '] = ',
                      File_as_Array[line][__Column_208])
                print('File_as_Array[line][__Column_232] = ', 'File_as_Array[', line, '][', __Column_232, '] = ',
                      File_as_Array[line][__Column_232])
                print('File_as_Array[line][__Column_238] = ', 'File_as_Array[', line, '][', __Column_238, '] = ',
                      File_as_Array[line][__Column_238])
                print('File_as_Array[line][__Column_202] = ', 'File_as_Array[', line, '][', __Column_202, '] = ',
                      File_as_Array[line][__Column_202])
                print('File_as_Array[line][__Column_204] = ', 'File_as_Array[', line, '][', __Column_204, '] = ',
                      File_as_Array[line][__Column_204])
                print('File_as_Array[line][__Column_207] = ', 'File_as_Array[', line, '][', __Column_207, '] = ',
                      File_as_Array[line][__Column_207])
                exit()

    def Get206(self):
        return self.__Signal_206

    def Get208(self):
        return self.__Signal_208

    def Get232(self):
        return self.__Signal_232

    def Get238(self):
        return self.__Signal_238

    def Get202(self):
        return self.__Signal_202

    def Get204(self):
        return self.__Signal_204

    def Get207(self):
        return self.__Signal_207

    def GetCyclesDateTime(self):
        return self.__CyclesDateTime

    def plot_All(self,
                 plot202=True,
                 plot204=True,
                 plot206=True,
                 plot207=True,
                 plot208=True,
                 plot232=True,
                 plot238=True):
        import matplotlib.pyplot as plt

        f, (ax206, ax208, ax232, ax238, ax202, ax204, ax207) = plt.subplots(7, figsize=(10, 10))
        Xs = self.__CyclesDateTime

        if plot206 == True:
            ax206.plot(Xs, self.__Signal_206, color='r')
            ax206.set_title('206 signal', fontsize=10)

        if plot208 == True:
            ax208.plot(Xs, self.__Signal_208, color='b')
            ax208.set_title('208 signal', fontsize=10)

        if plot232 == True:
            ax232.plot(Xs, self.__Signal_232, color='g')
            ax232.set_title('232 signal', fontsize=10)

        if plot238 == True:
            ax238.plot(Xs, self.__Signal_238, color='y')
            ax238.set_title('238 signal', fontsize=10)

        if plot202 == True:
            ax202.plot(Xs, self.__Signal_202, color='m')
            ax202.set_title('202 signal', fontsize=10)

        if plot204 == True:
            ax204.plot(Xs, self.__Signal_204, color='b')
            ax204.set_title('204 signal', fontsize=10)

        if plot207 == True:
            ax207.plot(Xs, self.__Signal_207, color='c')
            ax207.set_title('207 signal', fontsize=10)

        # Fine-tune figure; make subplots close to each other and hide x ticks for
        # all but bottom plot.
        f.subplots_adjust(hspace=2)
        plt.setp([a.get_xticklabels() for a in f.axes[:-1]], visible=False)
        plt.show()

    def print_RawData(self):
        ROW_FORMAT = '| {0!s:10s} | {1!s:10s}'

        for line in range(0, len(self.__RawData)):
            # print(ROW_FORMAT.format(self.__RawData[line]))
            print('Line ', line, ' = ', self.__RawData[line])

    def CUMSUM(self):

        x = self.__Signal_207
        threshold = 0.25 * max(x)
        print(threshold)
        drift = threshold/8
        print(drift)
        ta, tai, taf, amp = detect_cusum.detect_cusum(x, threshold, drift, True, True, 1)
        # x = self.__Signal_206
        # ta, tai, taf, amp = detect_cusum.detect_cusum(x, 0.01, 0.0005, True, True, 1)
        # print('ta = ', ta)
        # print('tai = ', tai)
        # print('taf = ', taf)
        # print('amp = ', amp)


def ListFiles(FolderPath, Extension):
    '''
    Returns a list of files with the desired extension in the indicated folder
    '''

    import os

    # print(FolderPath)

    if type(FolderPath) is not str:
        print('Folder is not a string.')
        exit()
    else:
        FolderPath = os.path.normcase(
            FolderPath)  # Normalize the case of a pathname. On Unix and Mac OS X, this returns the path unchanged;
        # on case-insensitive filesystems, it converts the path to lowercase. On Windows, it also converts forward slashes to backward
        # slashes.

    if type(Extension) is not str:
        print('Extension must be a string.')
        exit()

    # print(FolderPath)


    if os.path.exists(FolderPath) == False:
        print('The folder does not exist.')
        exit()

    FilesList = []

    for root, dirs, files in os.walk(FolderPath, topdown=False):
        # print(root)
        for name in files:

            if os.path.splitext(os.path.join(root, name))[1] == Extension:
                FilesList.append([name, os.path.join(root, name)])

                # for name in dirs:
                #     print(os.path.join(root, name))

    if len(FilesList) == 0:
        print('No files with the desired extension (' + Extension + ') were found.')

    # print(len(FilesList))
    return FilesList


def LoadFiles(FilesList):
    '''
    Transforms each file in an array and append it to LoadedFiles as [SampleName,DataArray]
    '''

    import csv

    if len(FilesList) == 0:
        print("FileList is empty!")
        exit()

    LoadedFiles = []

    for file in FilesList:
        a = csv.reader(open(file[1], 'r'), delimiter='\t')  # delimiter='\t' is the tab delimiter

        temp = []
        for line in a:  # I could not find an easy way to import an tab delimited with text and values file
            # into an array in python. But this must exist, I just could not find it!
            temp.append(line)
            # print(line)

        LoadedFiles.append([file[0], temp])

    return LoadedFiles


a = ListFiles(r'D:\UnB\Projetos-Software\Chronus\Software\Blank_Test\teste91500_10042017\91500_TESTE', '.exp')
b = LoadFiles(a)
c = SampleData(b[0][0], b[0][1], 'Thermo Finnigan Neptune')
# print(c.Name)

import pprint

# c.print_RawData()

# c.plot_All(plot202=True,
#            plot204=True,
#            plot206=True,
#            plot207=True,
#            plot208=True,
#            plot232=True,
#            plot238=True)
#
# # for n in c.GetCyclesDateTime():
# #     print(str(n))
c.CUMSUM()
# import matplotlib.pyplot as plt
# plt.plot(c.Get207())
# plt.show()

# for n in c.Get207():
#     print(n)
# print(c.Get202())
# print(c.Get204())
# print(c.Get206())

# print(c.Get208())
# print(c.Get232())
# print(c.Get238())

# a = configuration()
# # print(a.__Constants['Column_232Th'])
# a.Print__Constants_Neptune()
# a.PrintSuportedEquipment()

# x = np.random.randn(300) / 5
#
# x[100:200] += np.arange(0, 4, 4 / 100)
# ta, tai, taf, amp = detect_cusum(x, 2, .02, True, True)
