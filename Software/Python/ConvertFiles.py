from __future__ import division, print_function
import pprint


class configuration:
    '''
    Some constants necessary to handle data in the files from some mass spectrometers
    '''

    def __init__(self):
        # 206F and 206MIC describes the differents methods employed at Brasília when the 206 mass is analyzed using
        # a faraday cup or an ion counter

        self.__Constants_Neptune_206F = {
            'Standard_NumberCycles': 60,
            'Line_IsotopeHeaders': 14,
            'Line_Cups': 76,
            'Line_AnalysisDate': 8,
            'Line_FirstCycle': 15,
            'Line_LastCycle': 74,
            'Line_new_num_cycles': 1,
            # This line sets the line where the new number of cyles (after separating blanks and samples) should be recorded
            'Column_AnalysisDate': 0,
            'Column_CycleTime': 1,
            'Column_206': 2,  # ¨False¨ if the isotope is not analyzed
            'Column_208': 3,  # ¨False¨ if the isotope is not analyzed
            'Column_232': 4,  # ¨False¨ if the isotope is not analyzed
            'Column_238': 5,  # ¨False¨ if the isotope is not analyzed
            'Column_202': 6,  # ¨False¨ if the isotope is not analyzed
            'Column_204': 7,  # ¨False¨ if the isotope is not analyzed
            'Column_207': 8,  # ¨False¨ if the isotope is not analyzed
            'Columns_num': 13,
            'Column_new_num_cycles_header': 2,
            'Column_new_num_cycles': 3,
            'Column_split_points': 4,
            'FirstCycle_datetime_fmt': '%H:%M:%S:%f',
            # This line sets the columns where the new number of cyles (after separating blanks and samples) should be recorded
            'AnalysisDate_datetime_fmt': '%d/%m/%Y'  # Format of the date in the cell where the analysis date
            # was recorded. Some equipment may presente the date and the time (corresponding to the time of the first cycle?)
        }

        self.__Constants_Neptune_206MIC = {
            'Standard_NumberCycles': 60,
            'Line_IsotopeHeaders': 14,
            'Line_Cups': 75,
            'Line_AnalysisDate': 8,
            'Line_FirstCycle': 15,
            'Line_LastCycle': 74,
            'Line_new_num_cycles': 1,
            # This line sets the line where the new number of cyles (after separating blanks and samples) should be recorded
            'Column_AnalysisDate': 0,
            'Column_CycleTime': 1,
            'Column_206': 7,  # ¨False¨ if the isotope is not analyzed
            'Column_208': 2,  # ¨False¨ if the isotope is not analyzed
            'Column_232': 3,  # ¨False¨ if the isotope is not analyzed
            'Column_238': 4,  # ¨False¨ if the isotope is not analyzed
            'Column_202': 5,  # ¨False¨ if the isotope is not analyzed
            'Column_204': 6,  # ¨False¨ if the isotope is not analyzed
            'Column_207': 8,  # ¨False¨ if the isotope is not analyzed
            'Columns_num': 13,
            'Column_new_num_cycles_header': 2,
            'Column_new_num_cycles': 3,
            'Column_split_points': 4,
            'FirstCycle_datetime_fmt': '%H:%M:%S:%f',
            # This line sets the columns where the new number of cyles (after separating blanks and samples) should be recorded
            'AnalysisDate_datetime_fmt': '%d/%m/%Y'  # Format of the date in the cell where the analysis date
            # was recorded. Some equipment may presente the date and the time (corresponding to the time of the first cycle?)
        }

        self.__Constants_Spec2 = {
            'Standard_NumberCycles': 217,
            'Line_IsotopeHeaders': 3,
            'Line_AnalysisDate': 2,
            'Line_FirstCycle': 4,
            'Line_LastCycle': 220,
            'Line_new_num_cycles': 1,
            # This line sets the line where the new number of cyles (after separating blanks and samples) should be recorded
            'Column_AnalysisDate': 0,
            'Column_CycleTime': 0,
            'Column_206': 3,  # ¨False¨ if the isotope is not analyzed
            'Column_208': 5,  # ¨False¨ if the isotope is not analyzed
            'Column_232': 6,  # ¨False¨ if the isotope is not analyzed
            'Column_235': 7,  # ¨False¨ if the isotope is not analyzed
            'Column_238': 8,  # ¨False¨ if the isotope is not analyzed
            'Column_202': False,  # ¨False¨ if the isotope is not analyzed
            'Column_204': False,  # ¨False¨ if the isotope is not analyzed
            'Column_207': 4,  #¨False¨ if the isotope is not analyzed
            'Columns_num': 9,  # Number of columns in each file
            'Column_new_num_cycles_header': 2,
            'Column_new_num_cycles': 3,
            'Column_split_points': 4,
        # This line sets the columns where the new number of cyles (after separating blanks and samples) should be recorded
            'FirstCycle_datetime_fmt': '%S.%f',
            # https://docs.python.org/3.6/library/datetime.html#strftime-strptime-behavior
            'AnalysisDate_datetime_fmt': '%b %d %Y  %I:%M %p'  # Format of the date in the cell where the analysis date
            #was recorded. Some equipment may presente the date and the time (corresponding to the time of the first cycle?)
        }

        self.MassSpectometersSuported = {
            'Thermo Finnigan Neptune_206F': self.__Constants_Neptune_206F,
            'Thermo Finnigan Neptune_206MIC': self.__Constants_Neptune_206MIC,
            'Spec2': self.__Constants_Spec2
        }

    def PrintSuportedEquipment(self):
        for equip in self.MassSpectometersSuported:
            print(equip)

    def Print_Constants(self, Equipment_name):
        '''Prints all constants in the dictionary'''
        if Equipment_name in self.MassSpectometersSuported:
            pprint.pprint(self.MassSpectometersSuported[Equipment_name])

    def GetConfigurations(self, Equipment_name):

        if Equipment_name in self.MassSpectometersSuported:
            return self.MassSpectometersSuported[Equipment_name]
        # if Equipment_name == 'Thermo Finnigan Neptune_206F':
        #     return self.__Constants_Neptune_206F
        # if Equipment_name == 'Thermo Finnigan Neptune_206MIC':
        #     return self.__Constants_Neptune_206MIC
        # if Equipment_name == 'Spec2':
        #     return self.__Constants_Spec2
        else:
            print(Equipment_name, ' is not currently supported.')
            exit()


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

        import os

        super(SampleData, self).__init__()

        self.ID = SampleData.__LastID
        SampleData.__LastID += 1  #Used to count the number of samples

        self.__Constants = self.GetConfigurations(Equipment)
        self.Name = Name
        self.equipment = Equipment
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
        self.standard_output_extension = 'dynamic.exp'

        # I am thinking if the lines below were a good idea. They make the code more clean, but if I need to change
        # come self.property, the variable
        __Line_new_num_cycles = self.__Constants['Line_new_num_cycles']
        __Line_FirstCycle = self.__Constants['Line_FirstCycle']
        __Line_LastCycle = self.__Constants['Line_LastCycle']
        __Line_AnalysisDate = self.__Constants['Line_AnalysisDate']

        __Column_split_points = self.__Constants['Column_split_points']
        __Column_new_num_cycles = self.__Constants['Column_new_num_cycles']
        __Column_CycleTime = self.__Constants['Column_CycleTime']
        __Column_206 = self.__Constants['Column_206']
        __Column_208 = self.__Constants['Column_208']
        __Column_232 = self.__Constants['Column_232']
        __Column_238 = self.__Constants['Column_238']
        __Column_202 = self.__Constants['Column_202']
        __Column_204 = self.__Constants['Column_204']
        __Column_207 = self.__Constants['Column_207']
        __Column_AnalysisDate = self.__Constants['Column_AnalysisDate']

        __FirstCycle_datetime_fmt = self.__Constants[
            'FirstCycle_datetime_fmt']  # Format of the date information in the cell
        # that records the date of analysis
        __AnalysisDate_datetime_fmt = self.__Constants[
            'AnalysisDate_datetime_fmt']  # Format of the date and time information in the cell
        #that records the date of analysis

        try:
            AnalysisDate = File_as_List[__Line_AnalysisDate][__Column_AnalysisDate]
            FirstCycleTime = File_as_List[__Line_FirstCycle][__Column_CycleTime]

        except:
            print()
            print('__Line_AnalysisDate =', __Line_AnalysisDate)
            print('__Line_FirstCycle =', __Line_FirstCycle)
            print('__Column_AnalysisDate =', __Column_AnalysisDate)
            print('__Column_CycleTime =', __Column_CycleTime)
            print('self.folderpath =', os.path.join(self.folderpath, self.Name))
            print()
            print('Printing raw data...')
            self.print_RawData()
            exit()

        # Below the date of analysis and the first time cycle will be combined in a single datetime objetc.
        Analysis_Date_FirstTime = CombineDateTime(
            AnalysisDate,
            FirstCycleTime,
            __AnalysisDate_datetime_fmt,
            __FirstCycle_datetime_fmt,
            self.equipment, self.Name)

        try:
            self.AnalysisDate = Analysis_Date_FirstTime.date()
            self.FirstCycleTime = Analysis_Date_FirstTime.time()

        except:
            print()
            print('__Line_AnalysisDate =', __Line_AnalysisDate)
            print('__Line_FirstCycle =', __Line_FirstCycle)
            print('__Column_AnalysisDate =', __Column_AnalysisDate)
            print('__Column_CycleTime =', __Column_CycleTime)
            print('AnalysisDate = ', File_as_List[__Line_AnalysisDate][__Column_AnalysisDate])
            print('FirstCycleTime = ', File_as_List[__Line_FirstCycle][__Column_CycleTime])
            print('self.folderpath =', os.path.join(self.folderpath, self.Name))
            print()
            print('Printing raw data...')
            self.print_RawData()
            exit()

        File_as_List[__Line_AnalysisDate][__Column_AnalysisDate] = 'Date: ' + Analysis_Date_FirstTime.strftime(
            '%d/%m/%Y')
        # print(File_as_List[__Line_AnalysisDate][__Column_AnalysisDate])

        for line in self.__RawData:

            if len(line) - 1 < self.__Constants['Columns_num']:
                for n in range(self.__Constants['Columns_num'] - len(line)):
                    line.append('')

            if len(line) != self.__Constants['Columns_num']:
                print('Class SampleData')
                print(self.Name)
                print('Lines with different number of columns.')
                print('len(line) = ', len(line))
                print('self.__Constants[', 'Columns_num', '] = ', self.__Constants['Columns_num'])
                exit()

        numcycles = self.__RawData[__Line_new_num_cycles][__Column_new_num_cycles]

        if numcycles != '':
            self.__Constants['Standard_NumberCycles'] = numcycles

        if len(self.__RawData[__Line_new_num_cycles][__Column_split_points]) > 0:
            self.split_points = self.__RawData[__Line_new_num_cycles][__Column_split_points]

        __Line_LastCycle = int(__Line_FirstCycle) + int(self.__Constants['Standard_NumberCycles'])

        self.__CyclesDateTime = []
        self.Signal_206 = []
        self.Signal_208 = []
        self.Signal_232 = []
        self.Signal_238 = []
        self.Signal_202 = []
        self.Signal_204 = []
        self.Signal_207 = []

        for line in range(__Line_FirstCycle, __Line_LastCycle):

            CyclesDateTime = CombineDateTime(
                AnalysisDate,
                File_as_List[line][__Column_CycleTime],
                __AnalysisDate_datetime_fmt,
                __FirstCycle_datetime_fmt,
                self.equipment, self.Name)

            if CyclesDateTime == 'Error':
                print('Class SampleData')
                print(self.Name)
                print('Error while interpreting file date.')
                print('line', line)
                print('__Column_CycleTime', __Column_CycleTime)
                print('File_as_Array[line][__Column_CycleTime]')
                exit()

            self.__CyclesDateTime.append(CyclesDateTime)
            File_as_List[line][__Column_CycleTime] = CyclesDateTime.strftime('%H:%M:%S:%f')[
                                                     :-3]  # Updates the cycle time to the
            # standard format used by Chronus (e.g. 13:54:24:852). A little help from http://stackoverflow.com/a/18406412/2449724

            try:
                if __Column_206 != False:
                    self.Signal_206.append(float(File_as_List[line][__Column_206]))

                if __Column_208 != False:
                    self.Signal_208.append(float(File_as_List[line][__Column_208]))

                if __Column_232 != False:
                    self.Signal_232.append(float(File_as_List[line][__Column_232]))

                if __Column_238 != False:
                    self.Signal_238.append(float(File_as_List[line][__Column_238]))

                if __Column_202 != False:
                    self.Signal_202.append(float(File_as_List[line][__Column_202]))

                if __Column_204 != False:
                    self.Signal_204.append(float(File_as_List[line][__Column_204]))

                if __Column_207 != False:
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

                if __Column_202 != False:
                    print('File_as_Array[line][__Column_202] = ', 'File_as_Array[', line, '][', __Column_202, '] = ',
                          File_as_List[line][__Column_202])
                else:
                    print('__Column_202 = ', __Column_202)

                if __Column_204 != False:
                    print('File_as_Array[line][__Column_204] = ', 'File_as_Array[', line, '][', __Column_204, '] = ',
                          File_as_List[line][__Column_204])
                else:
                    print(' __Column_204 =', __Column_204)

                print('File_as_Array[line][__Column_207] = ', 'File_as_Array[', line, '][', __Column_207, '] = ',
                      File_as_List[line][__Column_207])
                exit()

    def Print_CyclesDateTime(self):
        datetime = []
        [datetime.append(str(x)) for x in self.__CyclesDateTime]
        return datetime

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
        
        :param print_separeted_files: shoulde the program print the files split?
        :param new_split_points: list with four itens, the first and the second describe the start and the end of the 
        blank, and the the third and the forth describe the beginning and the end of the sample
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
        Line_LastCycle = self.__Constants['Line_LastCycle']

        # The conditinal clause below check if there is a file in the files_split folder with the name of the file that
        # would be created (A little confuse this explanation, is not it? GOSH!)
        if os.path.exists(
                os.path.join(
                    self.folderpath,
                    'files_split',
                    self.Name
                )
        ):

            print(self.Name, 'was already split using the split points',
                  self.__RawData[Line_new_num_cycles][Column_split_points])
            print('You must delete the files created previously by splitting the original files.')
            print('Check the folder', os.path.join(self.folderpath,'files_split'))
            exit()

        if self.split_points != new_split_points:
            self.update_split_points(new_split_points)
            split_points = list(self.split_points)

        if len(split_points) != 4:
            print(self.Name)
            print(
                'You must provide 4 breaking points corresponding to the "start of the blank", "end of the blank", "start of the sample" and "end of the sample".')
            exit()

        # print('new_split_points[3]',new_split_points[3])
        # print('Line_LastCycle',Line_LastCycle)
        if new_split_points[3] - 1 > Line_LastCycle:
            print(self.Name)
            print(
                'The indicated breaking point related to the "end of the sample" is bigger than the number of cycles.')
            exit()

        Blank = []
        Sample = []
        BothFile_Header = []

        for line in range(Line_IsotopeHeaders + 1):
            # print(self.__RawData[line])
            BothFile_Header.append(self.__RawData[line])
        # print()
        for line in range(Line_FirstCycle + split_points[0] - 1, Line_FirstCycle + split_points[1]):
            # print(self.__RawData[line])
            Blank.append(self.__RawData[line])
        # print()
        for line in range(Line_FirstCycle + split_points[2] - 1, Line_FirstCycle + split_points[3]):
            # print(self.__RawData[line])
            Sample.append(self.__RawData[line])
        # print()

        blank_with_header = deepcopy(BothFile_Header)  # Deepcopy was necessary because my list is multidimensional and
        # simply slicing do not solve my problem. Changes made in the higher dimensions, i.e. the lists within the
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

        if os._exists(os.path.join(self.folderpath, 'files_split')) == False:

            try:
                os.makedirs(os.path.join(self.folderpath, 'files_split'))
            except:
                # print('exist')
                pass

        WriteTXT(blank_with_header,
                 os.path.join(self.folderpath, 'files_split',
                              'blank_' + str(self.ID) + '.' + self.standard_output_extension),
                 self.equipment)

        WriteTXT(sample_with_header,
                 os.path.join(self.folderpath, 'files_split', self.Name),
                 self.equipment)


def WriteTXT(List, FileAddress,equipment):
    '''
    
    :param List: List to be to converted to a tab delimited file 
    :param FileAddress: Full path (including the name of the file with its extension) of the file to be created
    :return: nothing
    '''

    import csv
    # print(FileAddress)
    file = open(FileAddress, 'w', newline='')

    if equipment == 'Thermo Finnigan Neptune_206F':
        writer = csv.writer(file, delimiter='\t')
    if equipment == 'Thermo Finnigan Neptune_206MIC':
        writer = csv.writer(file, delimiter='\t')
    elif equipment == 'Spec2':
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


def Convert_Files_to_Lists(FilesList,Equipment):
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

    # print(Equipment)

    for file in FilesList:

        if Equipment == 'Thermo Finnigan Neptune_206F':
            file_read = csv.reader(open(file[1], 'r'), delimiter='\t')  # delimiter='\t' is the tab delimiter
        if Equipment == 'Thermo Finnigan Neptune_206MIC':
            file_read = csv.reader(open(file[1], 'r'), delimiter='\t')  # delimiter='\t' is the tab delimiter
        elif Equipment == 'Spec2':
            file_read = csv.reader(open(file[1], 'r'))
        else:
            print('Failed to read csv file. Equipment not in Chronus equipment list.')

        temp = []

        for line in file_read:  # I could not find an easy way to import an tab delimited with text and values file
            # into an array in python. But this must exist, I just could not find it!
            # print(line)

            temp.append(line)

        LoadedFile.append([file[0], temp])

    return LoadedFile


def LoadFiles(FolderPath, extension, delimiter='\t', equipment='Thermo Finnigan Neptune_206F'):
    '''
    
    :param FolderPath: Folder where all the raw data files are stored
    :param extension: Extension of the desired files (.exp for exported neptune files)
    :param delimter: Delimiter which will be used to split text into columns ('\t' for tab, ',' for comma)
    :return: A list with all files described by the SampleData class
    '''

    list_of_files = ListFiles(FolderPath, extension)

    list_of_files_converted_to_List = Convert_Files_to_Lists(list_of_files,equipment)

    loaded_files = []

    for file in list_of_files_converted_to_List:
        loaded_files.append(SampleData(file[0], FolderPath, file[1], equipment, extension, False))

    return loaded_files


def CombineDateTime(date_str, time_str, date_fmt, time_fmt, equipment, sample_name):
    '''
    Mixes the date of analysis and the start of a cycle into a datetime object.
    
    :param date_str: a string with a data and/or time recorded
    :param fmt: format of datetime_str according to https://docs.python.org/3.6/library/datetime.html#strftime-strptime-behavior
    :param equipment: depending on the equipment, the datetime_str must be sliced to be handled by the strpformat
    :return: datetime object
    '''

    #Updated based on the question at http://stackoverflow.com/q/44014424/2449724

    import datetime
    # print()
    if equipment == 'Spec2':
        # print(equipment)
        remove_start = date_str.find(':') + 2
        remove_end = date_str.find('using') - 1
        date_str = date_str[remove_start:remove_end]
    elif equipment == 'Thermo Finnigan Neptune_206F' or equipment == 'Thermo Finnigan Neptune_206MIC':
        # print(equipment)
        remove_start = date_str.find(':') + 2
        # print(date_str)
        date_str = date_str[remove_start:len(date_str)]
        # print('date_str', date_str)

    try:
        date_obj = datetime.datetime.strptime(date_str, date_fmt)
        date_obj_time = date_obj.time()

        time_obj = datetime.datetime.strptime(time_str, time_fmt)
        time_obj += datetime.timedelta(hours=date_obj_time.hour, minutes=date_obj_time.minute,
                                       seconds=date_obj_time.second,
                                       microseconds=date_obj_time.microsecond)

        # print()
        # TEMP = datetime.datetime.combine(date_obj.date(), time_obj.time())
        # print(TEMP.strftime('%H:%M:%S:%f'))

        return datetime.datetime.combine(date_obj.date(), time_obj.time())

    except Exception as e:
        print('Failed to combine date and time of', sample_name)
        print('time_str =', time_str)
        print('time_fmt =', time_fmt)
        print('date_str =', date_str)
        print('date_fmt =', date_fmt)
        print()
        print('Printing excepetion...')
        print(e)
        return 'Error'

