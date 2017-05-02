from __future__ import division, print_function

from Chronus_Python import detect_cusum


class configuration:
    '''
    Some constants necessary to handle data in the files from some mass spectrometers
    '''

    def __init__(self):
        self.MassSpectometersSuported = [
            'Thermo Finnigan Neptune', 'Spec2'
        ]

        self.__Constants_Neptune = {
            'Standard_NumberCycles': 85,
            'Line_IsotopeHeaders': 14,
            'Line_Cups': 101,
            'Line_AnalysisDate': 8,
            'Line_FirstCycle': 15,
            'Line_LastCycle': 100,
            'Line_new_num_cycles': 1,
            # This line sets the line where the new number of cyles (after separating blanks and samples) should be recorded
            'Column_AnalysisDate': 0,
            'Column_CycleTime': 1,
            'Column_206': 2,
            'Column_208': 3,
            'Column_232': 4,
            'Column_238': 5,
            'Column_202': 6,
            'Column_204': 7,
            'Column_207': 8,
            'Columns_num': 13,
            'Column_new_num_cycles_header': 2,
            'Column_new_num_cycles': 3,
            'Column_split_points': 4
            # This line sets the columns where the new number of cyles (after separating blanks and samples) should be recorded
        }

        self.__Constants_Spec2 = {
            'Standard_NumberCycles': 217,
            'Line_IsotopeHeaders': 4,
            'Line_AnalysisDate': 3,
            'Line_FirstCycle': 5,
            'Line_LastCycle': 221,
            'Line_new_num_cycles': 1,
            # This line sets the line where the new number of cyles (after separating blanks and samples) should be recorded
            'Column_AnalysisDate': 0,
            'Column_CycleTime': 1,
            'Column_206': 4,
            'Column_208': 6,
            'Column_232': 7,
            'Column_238': 9,
            'Column_207': 5,
            'Columns_num': 9,  # Number of columns in each file
            'Column_new_num_cycles_header': 2,
            'Column_new_num_cycles': 3,
            'Column_split_points': 4
            # This line sets the columns where the new number of cyles (after separating blanks and samples) should be recorded
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
            folderpath - Folder where you would like to save the plots that you might generate
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
    __LastID = 1

    def __init__(
            self,
            Name,
            folderpath,
            File_as_List,
            Equipment,
            FileExtension,
            blank_together=False,
            split_points=[],
            unit_202='cps',
            unit_204='cps',
            unit_206='mV',
            unit_207='cps',
            unit_208='cps',
            unit_232='mV',
            unit_238='mV'
    ):

        super(SampleData, self).__init__()

        self.ID = SampleData.__LastID
        SampleData.__LastID += 1

        import datetime

        self.__Constants = self.GetConfigurations(Equipment)
        self.Name = Name
        self.folderpath = folderpath
        self.FileExtension = FileExtension
        self.__RawData = File_as_List
        self.blank_together = blank_together
        self.split_points = split_points
        self.unit_202 = unit_202
        self.unit_204 = unit_204
        self.unit_206 = unit_206
        self.unit_207 = unit_207
        self.unit_208 = unit_208
        self.unit_232 = unit_232
        self.unit_238 = unit_238

        Line_new_num_cycles = self.__Constants['Line_new_num_cycles']
        Column_split_points = self.__Constants['Column_split_points']
        if len(self.__RawData[Line_new_num_cycles][Column_split_points]) > 0:
            self.split_points = self.__RawData[Line_new_num_cycles][Column_split_points]

        # print('self.__Constants[Line_new_num_cycles]', self.__Constants['Line_new_num_cycles'])
        # print('self.__Constants[Column_new_num_cycles]', self.__Constants['Column_new_num_cycles'])

        # print(self.__RawData)
        numcycles = self.__RawData[self.__Constants['Line_new_num_cycles']][self.__Constants['Column_new_num_cycles']]
        # print('numcycles',numcycles)
        if numcycles != '':
            self.__Constants['Standard_NumberCycles'] = numcycles

        AnalysisDate = File_as_List[self.__Constants['Line_AnalysisDate']][self.__Constants['Column_AnalysisDate']]

        self.AnalysisDate = datetime.datetime.strptime(AnalysisDate.split(' ')[2], "%d/%m/%Y")

        FirstCycleTime = File_as_List[self.__Constants['Line_FirstCycle']][self.__Constants['Column_CycleTime']]
        self.FirstCycleTime = datetime.datetime.strptime(FirstCycleTime, '%H:%M:%S:%f')

        # self.__DateTimeAnalysis = datetime.datetime.combine(AnalysisDate,FirstCycleTime.time())
        # print(self.__DateTimeAnalysis)

        for line in self.__RawData:

            if len(line) != self.__Constants['Columns_num']:
                print('Class SampleData')
                print(self.Name)
                print('Lines with different number of columns.')
                exit()

        Line_FirstIntensity = self.__Constants['Line_FirstCycle']
        Line_LastIntensity = int(Line_FirstIntensity) + int(self.__Constants['Standard_NumberCycles'])

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

            try:
                CyclesTime = datetime.datetime.strptime(File_as_List[line][__Column_CycleTime], '%H:%M:%S:%f')
            except:
                print('Class SampleData')
                print(self.Name)
                print('Error while interpreting file date.')
                print('line', line)
                print('__Column_CycleTime', __Column_CycleTime)
                print('File_as_Array[line][__Column_CycleTime]')
                pass

            self.__CyclesDateTime.append(datetime.datetime.combine(self.AnalysisDate.date(), CyclesTime.time()))

            try:
                self.Signal_206.append(float(File_as_List[line][__Column_206]))
                self.Signal_208.append(float(File_as_List[line][__Column_208]))
                self.Signal_232.append(float(File_as_List[line][__Column_232]))
                self.Signal_238.append(float(File_as_List[line][__Column_238]))
                self.Signal_202.append(float(File_as_List[line][__Column_202]))
                self.Signal_204.append(float(File_as_List[line][__Column_204]))
                self.Signal_207.append(float(File_as_List[line][__Column_207]))

            except:
                import os
                print('Class SampleData')
                print(self.Name)
                print(os.path.join(self.folderpath, self.Name))
                print('Error while converting str to float.')
                print('File_as_Array[line][__Column_206] = ', 'File_as_Array[', line, '][', __Column_206, '] = ',
                      File_as_List[line][__Column_206])
                print('File_as_Array[line][__Column_208] = ', 'File_as_Array[', line, '][', __Column_208, '] = ',
                      File_as_List[line][__Column_208])
                print('File_as_Array[line][__Column_232] = ', 'File_as_Array[', line, '][', __Column_232, '] = ',
                      File_as_List[line][__Column_232])
                print('File_as_Array[line][__Column_238] = ', 'File_as_Array[', line, '][', __Column_238, '] = ',
                      File_as_List[line][__Column_238])
                print('File_as_Array[line][__Column_202] = ', 'File_as_Array[', line, '][', __Column_202, '] = ',
                      File_as_List[line][__Column_202])
                print('File_as_Array[line][__Column_204] = ', 'File_as_Array[', line, '][', __Column_204, '] = ',
                      File_as_List[line][__Column_204])
                print('File_as_Array[line][__Column_207] = ', 'File_as_Array[', line, '][', __Column_207, '] = ',
                      File_as_List[line][__Column_207])
                exit()

                # if self.blank_together == True and len(self.split_points) == 0:  # If sample and blank were recorded together
                #     # and the user didn't indicated the breaking points, it is necessary to determine them
                #
                #     drift_206 = 0
                #     drift_207 = 0
                #     drift_232 = 0
                #     drift_238 = 0
                #
                #     ShowParameters = True
                #
                #     if self.unit_206 == 'mV':
                #         drift_206 = 0.01
                #     elif self.unit_206 == 'cps':
                #         drift_206 = 6250000
                #     else:
                #         print('206 unit (', self.unit_206, ' unrecognized')
                #         exit()
                #
                #     if self.unit_207 == 'mV':
                #         drift_207 = 0.01
                #     elif self.unit_207 == 'cps':
                #         drift_207 = 6250000
                #     else:
                #         print('206 unit (', self.unit_207, ' unrecognized')
                #         exit()
                #
                #     if self.unit_232 == 'mV':
                #         drift_232 = 0.01
                #     elif self.unit_232 == 'cps':
                #         drift_232 = 6250000
                #     else:
                #         print('206 unit (', self.unit_232, ' unrecognized')
                #         exit()
                #
                #     if self.unit_238 == 'mV':
                #         drift_238 = 0.01
                #     elif self.unit_238 == 'cps':
                #         drift_238 = 6250000
                #     else:
                #         print('206 unit (', self.unit_238, ' unrecognized')
                #         exit()
                #
                #     parameters_206 = [self.Signal_206, 0.005, 0.0005]
                #     parameters_207 = [self.Signal_207, 20000, 5000]
                #     parameters_232 = [self.Signal_232, max(self.Signal_232) / 2, 0.01 * max(self.Signal_232) / 2]
                #     parameters_238 = [self.Signal_238, 0.05, 0.005]
                #
                #     self.__CUMSUM_Parameters_206 = detect_cusum.detect_cusum(
                #         parameters_206[0],
                #         parameters_206[1],
                #         parameters_206[2],
                #         True,
                #         ShowParameters
                #     )
                #     self.__CUMSUM_Parameters_207 = detect_cusum.detect_cusum(
                #         parameters_207[0],
                #         parameters_207[1],
                #         parameters_207[2],
                #         True,
                #         ShowParameters
                #     )
                #     self.__CUMSUM_Parameters_232 = detect_cusum.detect_cusum(
                #         parameters_232[0],
                #         parameters_232[1],
                #         parameters_232[2],
                #         True,
                #         ShowParameters
                #     )
                #     self.__CUMSUM_Parameters_238 = detect_cusum.detect_cusum(
                #         parameters_238[0],
                #         parameters_238[1],
                #         parameters_238[2],
                #         True,
                #         ShowParameters
                #     )
                #
                # if self.blank_together == True and len(self.split_points) != 0:
                #     print('Prepare blank together with breaking points')

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
                fig, ax1 = plt.subplots()
                ax1.plot(Xs, intensity[1], color='r', marker='8')
                ax1.set_xlabel('Time', color='b', fontsize=12)
                # ax1.set_title(intensity[0], fontsize=10)
                fig.autofmt_xdate()  # I dont't know if this have another effect, but the one I wanted was it automatically rotates the x axis labels
                plt.ylabel('Intensity', color='b', fontsize=12)

                ax2 = ax1.twiny()
                Xs2 = range(1, len(Xs) + 1)
                ax2.plot(Xs2, intensity[1], color='r', marker='8')
                ax2.set_xlabel('Cycles', color='b', fontsize=12)
                ax2.tick_params('y', colors='r')

                if saveplots == True:
                    import os

                    try:
                        plt.savefig(os.path.join(self.folderpath, str(self.Name + intensity[0] + '.pdf')))
                        print(self.Name, 'plots saved to',
                              str(os.path.join(self.folderpath, str(self.Name + intensity[0] + '.pdf'))))
                    except:
                        print('Could not save', self.Name, 'plots to',
                              str(os.path.join(self.folderpath, str(self.Name + intensity[0] + '.pdf'))))

                if showplot == True:
                    plt.show()  # This must be done after saving the figure because it creates a new figure
                    # https://stackoverflow.com/questions/9012487/matplotlib-pyplot-savefig-outputs-blank-image

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
                ax.plot(Xs, intensity[1], marker='8')
                ax.set_title(intensity[0], fontsize=10)

            fig.subplots_adjust(left=0.1, right=0.9, bottom=0.1, top=0.9,
                                wspace=0.2, hspace=0.5)

            fig.autofmt_xdate()

            if saveplots == True:

                import os
                try:
                    plt.savefig(os.path.join(self.folderpath, str(self.Name + '.pdf')))
                    print(self.Name, 'plots saved to', str(os.path.join(self.folderpath, str(self.Name + '.pdf'))))
                except:
                    print('Could not save', self.Name, 'plots to',
                          str(os.path.join(self.folderpath, str(self.Name + '.pdf'))))

            if showplot == True:
                plt.show()  # This must be done after saving the figure because it creates a new figure
                # https://stackoverflow.com/questions/9012487/matplotlib-pyplot-savefig-outputs-blank-image

    def print_RawData(self):
        ROW_FORMAT = '| {0!s:10s} | {1!s:10s}'

        for line in range(0, len(self.__RawData)):
            # print(ROW_FORMAT.format(self.__RawData[line]))
            print('Line ', line, ' = ', self.__RawData[line])

    def update_split_points(self, new_split_points):

        self.split_points = new_split_points

        print(self.Name, 'breaking points updated to', new_split_points)

    def split_blank_sample(self, print_separeted_files=False, new_split_points=[]):
        '''
        Using the given breaking points, this procedure will split the raw data file in two (blank and sample).
        These files will be saved to a subfolder inside the self.folderpath
        :return: 
        '''

        from copy import deepcopy
        import os
        split_points = self.split_points

        Line_new_num_cycles = self.__Constants['Line_new_num_cycles']
        Column_new_num_cycles = self.__Constants['Column_new_num_cycles']
        Column_new_num_cycles_header = self.__Constants['Column_new_num_cycles_header']
        Column_split_points = self.__Constants['Column_split_points']
        Line_IsotopeHeaders = self.__Constants['Line_IsotopeHeaders']
        Line_FirstCycle = self.__Constants['Line_FirstCycle']

        # print(' self.split_points', self.split_points)
        # print(self.__RawData[Line_new_num_cycles][Column_new_num_cycles])
        # print(type(self.__RawData[Line_new_num_cycles][Column_new_num_cycles]))

        # print(self.__RawData[Line_new_num_cycles][Column_new_num_cycles])
        # print(self.__RawData[Line_new_num_cycles][Column_new_num_cycles] != '')
        # print(self.folderpath)
        if self.__RawData[Line_new_num_cycles][Column_new_num_cycles] != '':
            print(self.Name, 'was already split using the split points',
                  self.__RawData[Line_new_num_cycles][Column_split_points])
            print('You must delete the files created previously by splitting the original files.')
            print('Check the folder', self.folderpath)
            exit()

        # print('new_split_points', new_split_points)
        # print('self.split_points', self.split_points)
        # print(self.split_points != new_split_points)
        if self.split_points != new_split_points:
            self.update_split_points(new_split_points)
            split_points = list(self.split_points)

        if len(split_points) != 4:
            print(self.Name)
            print(
                'You must provide 4 breaking points corresponding to the "start of the blank", "end of the blank", "start of the sample" and "end of the sample".')
            exit()

        Blank = []
        Sample = []
        BothFile_Header = []

        for line in range(Line_IsotopeHeaders + 1):
            BothFile_Header.append(self.__RawData[line])

        for line in range(Line_FirstCycle + split_points[0] - 1, Line_FirstCycle + split_points[1]):
            Blank.append(self.__RawData[line])

        for line in range(Line_FirstCycle + split_points[2] - 1, Line_FirstCycle + split_points[3]):
            Sample.append(self.__RawData[line])

        blank_with_header = deepcopy(BothFile_Header)  # Deepcopy was necessary because my list is multidimensional and
        # simply slicing do not solve my problema. Changes made in the higher dimensions, i.e. the lists within the
        # original list, will be shared between the two.
        # https://stackoverflow.com/questions/8744113/python-list-by-value-not-by-reference

        sample_with_header = deepcopy(BothFile_Header)

        blank_with_header.extend(Blank)
        sample_with_header.extend(Sample)

        blank_with_header[Line_new_num_cycles][Column_new_num_cycles] = split_points[1] - split_points[
            0] + 1  # new number of cycles
        blank_with_header[Line_new_num_cycles][Column_new_num_cycles_header] = 'Number of cycles'
        blank_with_header[Line_new_num_cycles][Column_split_points] = self.split_points

        sample_with_header[Line_new_num_cycles][Column_new_num_cycles] = split_points[3] - split_points[
            2] + 1  # new number of cycles
        sample_with_header[Line_new_num_cycles][Column_new_num_cycles_header] = 'Number of cycles'
        sample_with_header[Line_new_num_cycles][Column_split_points] = self.split_points

        if print_separeted_files == True:

            print()
            print('BothFile_Header')
            for line in BothFile_Header:
                print(line)

            print()
            print('blank_with_header')
            for line in blank_with_header:
                print(line)

            print()
            print('sample_with_header')
            for line in sample_with_header:
                print(line)

        # print(os.path.join(self.folderpath,'files_split'))
        # print(os._exists(os.path.join(self.folderpath,'files_split')))

        if os._exists(os.path.join(self.folderpath, 'files_split')) == False:

            try:
                os.makedirs(os.path.join(self.folderpath, 'files_split'))
            except:
                # print('exist')
                pass

        # print(os.path.join(self.folderpath,'files_split', 'blank_' + self.Name))
        WriteTXT(blank_with_header,
                 os.path.join(self.folderpath, 'files_split', 'blank_' + str(self.ID) + self.FileExtension))
        # print(sample_with_header, os.path.join(self.folderpath,'files_split', 'sample_' + self.Name))
        WriteTXT(sample_with_header, os.path.join(self.folderpath, 'files_split', self.Name))


def WriteTXT(List, FileAddress):
    '''
    
    :param List: List to be to converted to a tab delimited file 
    :param FileAddress: Full path (including the name of the file with its extension) of the file to be created
    :return: nothing
    '''

    import csv
    # print(FileAddress)
    file = open(FileAddress, 'w', newline='')

    writer = csv.writer(file, delimiter='\t')
    writer.writerows(List)


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
    '''
    
    :param FolderPath: Folder where all the raw data files are stored
    :param extension: Extension of the desired files (.exp for exported neptune files)
    :return: A list with all files described by the SampleData class
    '''

    list_of_files = ListFiles(FolderPath, extension)

    list_of_files_converted_to_List = Convert_Files_to_Lists(list_of_files)

    loaded_files = []

    for file in list_of_files_converted_to_List:
        loaded_files.append(SampleData(file[0], FolderPath, file[1], 'Thermo Finnigan Neptune', extension, False))

    return loaded_files

# a = [[1,2,3],[1,2,3],[1,2,3]]
# folderpt = r'D:\UnB\Projetos-Software\Chronus\Software\Blank_Test\teste91500_10042017\91500_TESTE\testestest.txt'
# WriteTXT(a,folderpt)
