import numpy as np
import matplotlib.pyplot as plt

primordial_lead_206_204_TAT73 = 9.307  # Tatsumoto et al 1973
primordial_lead_207_204_TAT73 = 10.294  # Tatsumoto et al 1973

primordial_lead_206_204_SK75_second_stage = 11.152  # Stacey and Kramers, 1975
primordial_lead_207_204_SK75_second_stage = 12.998  # Stacey and Kramers, 1975

primordial_uranium_lead_238_204_DS74 = 9.58  # Doe and Stacey 1974
primordial_uranium_lead_235_204_DS74 = primordial_uranium_lead_238_204_DS74 / 137.88

primordial_uranium_lead_238_204_SK75_first_stage = 7.19  # Stacey and Kramers, 1975
primordial_uranium_lead_235_204_SK75_first_stage = primordial_uranium_lead_238_204_SK75_first_stage / 137.88

primordial_uranium_lead_238_204_SK75_second_stage = 9.74  # Stacey and Kramers, 1975
primordial_uranium_lead_235_204_SK75_second_stage = primordial_uranium_lead_238_204_SK75_second_stage / 137.88

decay_constant_238_JA71 = 1.55125 * 10 ** -4  # Jaffey et al 1971
decay_constant_238_DS74 = 1.55 * 10 ** -4
decay_constant_235_J71 = 9.8485 * 10 ** -4  # Jaffey et al 1971
decay_constant_235_DS74 = 9.85 * 10 ** -4

earth_age_SK75_first_stage = 4570  # Stacey and Kramers, 1975
earth_age_SK75_second_stage = 3700  # Stacey and Kramers, 1975
earth_age_DS74 = 4430  # Doe and Stacey, 1974


def OpenIsoplot():
    import win32com.client

    try:
        xl = win32com.client.Dispatch("Excel.Application")
        xl.Visible = False

        xl.Workbooks.open(
            'D:\\UnB\\Programas\\TEMP\\Isoplot4.15.xlam')

        # https://stackoverflow.com/questions/345920/need-skeleton-code-to-call-excel-vba-from-pythonwin

        # print(xl.Application.Run('AgePb76',0.5))
        # print(xl.Application.Run('AgePb76',0.6))
        # print(xl.Application.Run('AgePb76',0.7))
        # print(xl.Application.Run('AgePb76',0.8))
        return xl

    except:
        print('Unexpected error while trying to open Isoplot.')
        return


def single_stage_lead_DK74(primordial_lead, primordial_uranium_lead, decay_constant, single_stage_start,
                           single_stage_end):
    '''
    :param primordial_lead: Initial 207/204 or 206/204 of the system
    :param primordial_uranium_lead: Initial 238/204 or 235/204 of the system
    :param decay_constant: 238 or 235 decay constant
    :param single_stage_start: usually taken to be the age of the earth
    :param single_stage_end: Geologict time at which single-stage evolution stopped
    :return: correspondent 207/204 or 206/204 ratio
    '''

    # print()
    # print('primordial_lead',primordial_lead)
    # print('primordial_uranium_lead',primordial_uranium_lead)
    # print('decay_constant',decay_constant)
    # print('single_stage_start',single_stage_start)
    # print('single_stage_end',single_stage_end)
    # print('np.exp(decay_constant * single_stage_start)',np.exp(decay_constant * single_stage_start))
    # print('np.exp(decay_constant * single_stage_end)',np.exp(decay_constant * single_stage_end))
    # print(primordial_lead + \
    #        primordial_uranium_lead * \
    #        (np.exp(decay_constant * single_stage_start) - \
    #         np.exp(decay_constant * single_stage_end)))

    return primordial_lead + \
           primordial_uranium_lead * \
           (np.exp(decay_constant * single_stage_start) -
            np.exp(decay_constant * single_stage_end))


def SingleStagePbR_Isoplot(Age, WhichRatio):
    '''
    
    :param Age: For each Age the ratios must be calculate?
    :param WhichRatio: 0 for 206Pb/204Pb and 1 for 207Pb/204Pb
    :return: the ratio for the specified age as calculated by Isoplot
    '''

    if WhichRatio < 0 or WhichRatio > 2:
        print('WhichRatio must be 0, 1 or 2 as Isoplot requires')

    Isoplot = OpenIsoplot()

    return Isoplot.Application.Run('SingleStagePbR', float(Age), WhichRatio)


def appendix_C_DS_1974(MaxAge):
    # ratios_64_single_stage and ratios_74_single_stage use the same parameters used to create the
    # Appendix C of Doe and Stacey (1974)

    Xs = np.arange(0, MaxAge + 1, 10)

    ratios_64_single_stage_DS74 = [single_stage_lead_DK74(
        primordial_lead_206_204_TAT73,
        primordial_uranium_lead_238_204_DS74,
        decay_constant_238_DS74,
        earth_age_DS74,
        age
    )
        for age in Xs]

    ratios_74_single_stage_DS74 = [single_stage_lead_DK74(
        primordial_lead_207_204_TAT73,
        primordial_uranium_lead_235_204_DS74,
        decay_constant_235_DS74,
        earth_age_DS74,
        age
    )
        for age in Xs]

    for a, b, c in zip(ratios_64_single_stage_DS74, ratios_74_single_stage_DS74, Xs):
        print('ratios_64_single_stage =', round(a, 4), 'ratios_74_single_stage =', round(b, 4), 'for age', c, 'Ma')

    plt.plot(ratios_64_single_stage_DS74, ratios_74_single_stage_DS74, label='ratios 64 and 74 from single stage')
    plt.legend()
    plt.show()

    return (ratios_64_single_stage_DS74, ratios_74_single_stage_DS74)


def SingleStagePbR(MaxAge):
    # ratios_64_single_stage_SK75 and ratios_74_single_stage_SK75 use the same parameters used by Isoplot 4.15.
    # This should return the exact same values as the SingleStagePbR function of Isoplot for the 206Pb/204Pb and
    # 207Pb/204Pb

    Xs = np.arange(0, MaxAge + 1, 10)

    ratios_64_single_stage_SK75 = [single_stage_lead_DK74(
        primordial_lead_206_204_SK75_second_stage,
        primordial_uranium_lead_238_204_SK75_second_stage,
        decay_constant_238_JA71,
        earth_age_SK75_second_stage,
        age) for age in Xs]

    ratios_74_single_stage_SK75 = [single_stage_lead_DK74(
        primordial_lead_207_204_SK75_second_stage,
        primordial_uranium_lead_235_204_SK75_second_stage,
        decay_constant_235_J71,
        earth_age_SK75_second_stage,
        age) for age in Xs]

    # The lines below are use to compare the results of this function to those calculated by Isoplot in Excel
    # for a, b, c in zip(ratios_64_single_stage_SK75, ratios_74_single_stage_SK75, Xs):
    #     try:
    #         ratio_64_isoplot = SingleStagePbR_Isoplot(c, 0)
    #         ratio_74_isoplot = SingleStagePbR_Isoplot(c, 1)
    #
    #         print()
    #         print('ratios_64_single_stage =', round(a, 4))
    #         print('            64_Isoplot =', round(ratio_64_isoplot, 4))
    #         print('ratios_74_single_stage =', round(b, 4))
    #         print('            74_Isoplot =', round(ratio_74_isoplot, 4))
    #         print('for age', c, 'Ma')
    #
    #     except Exception as e:
    #         print(e)
    #         exit()

    plt.plot(ratios_64_single_stage_SK75, ratios_74_single_stage_SK75, label='ratios 64 and 74 from single stage')
    plt.legend()
    plt.show()

    return (ratios_64_single_stage_SK75, ratios_74_single_stage_SK75)


SingleStagePbR(3700)

# print(SingleStagePbR_Isoplot(3700,1))
