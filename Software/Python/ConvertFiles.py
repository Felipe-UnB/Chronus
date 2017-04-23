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


class SampleData(configuration):
    '''
    This object represents a single data file.

    parameters:
            Name - Name of the sample
            File_as_Array - Array representing the raw data file
            Equipment - Equipment used to analyze the sample (it must be described by the configuration class)
            blank_together (default = False) - Is blank recorded with each sample?
            breakingpoints (default is an empty tuple) - If the blank was recorded with the sample then the breaking
                points to separete both analyses must be provided
            unit_202 (default = 'cps') - Unit of the 202 mass intensities
            unit_204 (default = ='cps') - Unit of the 204 mass intensities
            unit_206 (default = ='mV') - Unit of the 206 mass intensities
            unit_207 (default = ='cps') - Unit of the 207 mass intensities
            unit_208 (default = ='cps') - Unit of the 208 mass intensities
            unit_232 (default = ='mV') - Unit of the 232 mass intensities
            unit_238 (default = ='mV) - Unit of the 238 mass intensities

    '''

    def __init__(
            self,
            Name,
            File_as_Array,
            Equipment,
            blank_together=False,
            breaking_points=(),
            unit_202='cps',
            unit_204='cps',
            unit_206='mV',
            unit_207='cps',
            unit_208='cps',
            unit_232='mV',
            unit_238='mV'
    ):
        super(SampleData, self).__init__()

        import datetime

        self.__Constants = self.GetConfigurations(Equipment)
        self.Name = Name
        self.__RawData = File_as_Array
        self.blank_together = blank_together
        self.breaking_points = breaking_points
        self.unit_202 = unit_202
        self.unit_204 = unit_204
        self.unit_206 = unit_206
        self.unit_207 = unit_207
        self.unit_208 = unit_208
        self.unit_232 = unit_232
        self.unit_238 = unit_238

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
        self.Signal_206 = []
        self.Signal_208 = []
        self.Signal_232 = []
        self.Signal_238 = []
        self.Signal_202 = []
        self.Signal_204 = []
        self.Signal_207 = []

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
                self.Signal_206.append(float(File_as_Array[line][__Column_206]))
                self.Signal_208.append(float(File_as_Array[line][__Column_208]))
                self.Signal_232.append(float(File_as_Array[line][__Column_232]))
                self.Signal_238.append(float(File_as_Array[line][__Column_238]))
                self.Signal_202.append(float(File_as_Array[line][__Column_202]))
                self.Signal_204.append(float(File_as_Array[line][__Column_204]))
                self.Signal_207.append(float(File_as_Array[line][__Column_207]))

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

        if self.blank_together == True and len(self.breaking_points) == 0:  # If sample and blank were recorded togeter
            # and the user didn't indicated the breaking points, it is necessary to determine them

            drift_206 = 0
            drift_207 = 0
            drift_232 = 0
            drift_238 = 0

            ShowParameters = True

            if self.unit_206 == 'mV':
                drift_206 = 0.01
            elif self.unit_206 == 'cps':
                drift_206 = 6250000
            else:
                print('206 unit (', self.unit_206, ' unrecognized')
                exit()

            if self.unit_207 == 'mV':
                drift_207 = 0.01
            elif self.unit_207 == 'cps':
                drift_207 = 6250000
            else:
                print('206 unit (', self.unit_207, ' unrecognized')
                exit()

            if self.unit_232 == 'mV':
                drift_232 = 0.01
            elif self.unit_232 == 'cps':
                drift_232 = 6250000
            else:
                print('206 unit (', self.unit_232, ' unrecognized')
                exit()

            if self.unit_238 == 'mV':
                drift_238 = 0.01
            elif self.unit_238 == 'cps':
                drift_238 = 6250000
            else:
                print('206 unit (', self.unit_238, ' unrecognized')
                exit()

            print('206')
            print(max(self.Signal_206))
            print(100 * 0.005 / max(self.Signal_206))
            print(100 * 0.0005 / max(self.Signal_206))

            print()
            print('207')
            print(max(self.Signal_207))
            print(100 * 20000 / max(self.Signal_207))
            print(100 * 5000 / max(self.Signal_207))

            print()
            print('232')
            print(max(self.Signal_232))
            print(100 * 0.001 / max(self.Signal_232))
            print(100 * 0.00005 / max(self.Signal_232))

            print()
            print('238')
            print(max(self.Signal_238))
            print(100 * 0.05 / max(self.Signal_238))
            print(100 * 0.005 / max(self.Signal_238))

            parameters_206 = [self.Signal_206, 0.005, 0.0005]
            parameters_207 = [self.Signal_207, 20000, 5000]
            parameters_232 = [self.Signal_232, max(self.Signal_232) / 2, 0.01 * max(self.Signal_232) / 2]
            parameters_238 = [self.Signal_238, 0.05, 0.005]

            self.__CUMSUM_Parameters_206 = detect_cusum.detect_cusum(
                parameters_206[0],
                parameters_206[1],
                parameters_206[2],
                True,
                ShowParameters
            )
            self.__CUMSUM_Parameters_207 = detect_cusum.detect_cusum(
                parameters_207[0],
                parameters_207[1],
                parameters_207[2],
                True,
                ShowParameters
            )
            self.__CUMSUM_Parameters_232 = detect_cusum.detect_cusum(
                parameters_232[0],
                parameters_232[1],
                parameters_232[2],
                True,
                ShowParameters
            )
            self.__CUMSUM_Parameters_238 = detect_cusum.detect_cusum(
                parameters_238[0],
                parameters_238[1],
                parameters_238[2],
                True,
                ShowParameters
            )

        if self.blank_together == True and len(self.breaking_points) != 0:
            print('Prepare blank together with breaking points')



    def GetCyclesDateTime(self):

        return self.__CyclesDateTime

    def plot_save_All(self,
                      SameFigure=True,
                      plot202=True,
                      plot204=True,
                      plot206=True,
                      plot207=True,
                      plot208=True,
                      plot232=True,
                      plot238=True,
                      showplot=True,
                      saveplots=True):
        '''
        
        :param SameFigure: Should all mass intensities be plotted in the same figure?
        :param plot202: Should the 202 mass intensities be plotted?
        :param plot204: Should the 204 mass intensities be plotted?
        :param plot206: Should the 206 mass intensities be plotted?
        :param plot207: Should the 207 mass intensities be plotted?
        :param plot208: Should the 208 mass intensities be plotted?
        :param plot232: Should the 232 mass intensities be plotted?
        :param plot238: Should the 238 mass intensities be plotted?
        :return: A plot as required
        
        '''
        import matplotlib.pyplot as plt

        Should_Plot = [
            ['202 intensities', self.Signal_202, plot202],
            ['204 intensities', self.Signal_204, plot204],
            ['206 intensities', self.Signal_206, plot206],
            ['207 intensities', self.Signal_207, plot207],
            ['208 intensities', self.Signal_208, plot208],
            ['232 intensities', self.Signal_232, plot232],
            ['238 intensities', self.Signal_238, plot238]
        ]

        Plots = []

        Plots_final = Should_Plot[:]

        NumPlots = 0
        for plot in Should_Plot:
            if plot[2] == True:
                NumPlots += 1
                Plots.append(plot)
            else:
                Plots_final.remove(plot)

        Xs = self.__CyclesDateTime

        if SameFigure == False:
            for intensity in Plots_final:
                fig, ax = plt.subplots()
                ax.plot(Xs, intensity[1], color='r')
                ax.set_title(intensity[0], fontsize=10)
                fig.autofmt_xdate()  # I dont't know if this have another effect, but the one I wanted was it automatically rotates the x axis labels

                if showplot == True:
                    fig.show()

                if saveplots == True:
                    print(self.Name + intensity[0] + '.pdf')
                    plt.savefig(self.Name + intensity[0] + '.pdf')

        elif SameFigure == True:
            num_subplots = len(Plots_final)
            num_rows = int(num_subplots / 2) + 1
            num_columns = 2
            position = 1
            fig = plt.figure(1, figsize=(13, 6))
            fig.suptitle(str(self.Name + ' mass intensities'), fontsize=15)

            for intensity in Plots_final:
                # add every single subplot to the figure with a for loop
                ax = fig.add_subplot(num_rows, num_columns, position)
                position += 1
                ax.plot(Xs, intensity[1])
                ax.set_title(intensity[0], fontsize=10)

            fig.subplots_adjust(left=0.1, right=0.9, bottom=0.1, top=0.9,
                                wspace=0.2, hspace=0.5)

            fig.autofmt_xdate()

            if showplot == True:
                plt.show()

            if saveplots == True:
                print(self.Name + '.pdf')
                plt.savefig(self.Name + '.pdf')


    def print_RawData(self):
        ROW_FORMAT = '| {0!s:10s} | {1!s:10s}'

        for line in range(0, len(self.__RawData)):
            # print(ROW_FORMAT.format(self.__RawData[line]))
            print('Line ', line, ' = ', self.__RawData[line])


def Set_breakingpoints(parameters_206, parameters_207, parameters_232, parameters_238, ShowParameters=False):
    '''    
    :param parameters_206: [list of the 206 intensities, threshold, drift]
    :param parameters_207: [list of the 207 intensities, threshold, drift]
    :param parameters_232: [list of the 232 intensities, threshold, drift]
    :param parameters_238: [list of the 238 intensities, threshold, drift]
    :return: [CUMSUM_parameters_206, CUMSUM_parameters_207, CUMSUM_parameters_232, CUMSUM_parameters_238]
    '''

    CUMSUM_206 = detect_cusum.detect_cusum(
        parameters_206[0],
        parameters_206[1],
        parameters_206[2],
        True,
        ShowParameters)

    # CUMSUM_207 = detect_cusum.detect_cusum(
    #     parameters_207[0],
    #     parameters_207[1],
    #     parameters_207[2],
    #     True,
    #     ShowParameters)
    #
    # CUMSUM_232 = detect_cusum.detect_cusum(
    #     parameters_232[0],
    #     parameters_232[1],
    #     parameters_232[2],
    #     True,
    #     ShowParameters)
    #
    # CUMSUM_238 = detect_cusum.detect_cusum(
    #     parameters_238[0],
    #     parameters_238[1],
    #     parameters_238[2],
    #     True,
    #     ShowParameters)


def WriteTXT(List, FileAddress):
    # print(FileAddress)
    SaveTo = open(FileAddress, 'w')

    for item in List:
        SaveTo.write("%s\n" % item)


def ListFiles(FolderPath, Extension):
    '''
    Returns a list of lists in the following format
        [name, os.path.join(root, name)]
        
    item[0] - files with the desired extension in the indicated folder
    item[0][0] - file name
    item[0][1] - file path
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


def Convert_Files_to_Lists(FilesList):
    '''
    Transforms each file in a list and append it to LoadedFiles as [SampleName,DataArray]
    Each item of the list is another list and represents a row of the original file.

    :param FilesList: List with the paths of the files to be opened
    :return: LoadedFiles, a list of the files as [SampleName,DataArray]
    '''
    import csv

    if len(FilesList) == 0:
        print("FileList is empty!")
        exit()

    LoadedFile = []

    for file in FilesList:
        a = csv.reader(open(file[1], 'r'), delimiter='\t')  # delimiter='\t' is the tab delimiter

        temp = []
        for line in a:  # I could not find an easy way to import an tab delimited with text and values file
            # into an array in python. But this must exist, I just could not find it!
            temp.append(line)
            # print(line)

        LoadedFile.append([file[0], temp])

    return LoadedFile


def LoadFiles(FolderPath, extension):
    list_of_files = ListFiles(FolderPath, extension)

    list_of_files_converted_to_List = Convert_Files_to_Lists(list_of_files)

    loaded_files = []

    for file in list_of_files_converted_to_List:
        loaded_files.append(SampleData(file[0], file[1], 'Thermo Finnigan Neptune', False))

    return loaded_files


loaded_files = LoadFiles(r'D:\UnB\Projetos-Software\Chronus\Software\Blank_Test\teste91500_10042017\91500_TESTE',
                         '.exp')

loaded_files[0].plot_save_All(
    SameFigure=False,
    plot202=True,
    plot204=True,
    plot206=True,
    plot207=True,
    plot208=True,
    plot232=True,
    plot238=True,
    showplot=True,
    saveplots=True)


# for file in loaded_files:
#     file.plot_All(
#         SameFigure=False,
#         plot202=True,
#         plot204=True,
#         plot206=True,
#         plot207=True,
#         plot208=True,
#         plot232=True,
#         plot238=True)

# c = SampleData(b[0][0], b[0][1], 'Thermo Finnigan Neptune',True)
# c = SampleData(b[1][0], b[1][1], 'Thermo Finnigan Neptune',True)
# c = SampleData(b[2][0], b[2][1], 'Thermo Finnigan Neptune',True)


# print(c.Name)
#
# c.print_RawData()

# c.plot_All(
#     SameFigure=True,
#     plot202=True,
#     plot204=True,
#     plot206=True,
#     plot207=True,
#     plot208=True,
#     plot232=True,
#     plot238=True)
