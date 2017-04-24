from Chronus_Python.ConvertFiles import *


def CUMSUM_206(loaded_files):
    for a in loaded_files:

        a.plot_save_All(
            SameFigure=False,
            plot202=False,
            plot204=False,
            plot206=True,
            plot207=False,
            plot208=False,
            plot232=False,
            plot238=False,
            showplot=True,
            saveplots=False)

        threshold = 0.3 * max(a.Signal_206)
        drift = 0.01 * max(a.Signal_206)
        drift = 0.1 * threshold

        CUMSUM = detect_cusum.detect_cusum(a.Signal_206, threshold, drift, True, False)

        print()
        print(a.Name)
        print('max Signal_206 ', round(max(a.Signal_206), 4))
        print('max Signal_206 ', round(max(a.Signal_206) * 6250000, 4))
        # print('threshold ', round(threshold,5), '(', round(100 * threshold / max(a.Signal_206), 2), 'of max', ')')
        # print('drift ', round(drift,5), '(', round(100 * drift / max(a.Signal_206), 2), 'of max', ')')
        # print('change started',CUMSUM[1])
        # print('change ended',CUMSUM[2])

        c = []
        for a in CUMSUM[1]:
            c.append(a)
        for a in CUMSUM[2]:
            c.append(a)

        c = set(c)
        c = list(c)
        c.sort()
        print('breaking_points ', c)

        rng = []
        for x in range(len(c) - 1):
            rng.append(c[x + 1] - c[x])

        print('rng', rng)


loaded_files = LoadFiles(r'D:\UnB\Projetos-Software\Chronus\Software\Blank_Test\teste91500_10042017\91500_TESTE',
                         '.exp')
a = loaded_files[0]

parameters_207 = [a.Signal_207, 50000, 10000]
parameters_232 = [a.Signal_232, max(a.Signal_232) / 2, 0.01 * max(a.Signal_232) / 2]
parameters_238 = [a.Signal_238, 0.05, 0.005]

# CUMSUM_206(loaded_files)
# a.plot_save_All(
#     SameFigure=False,
#     plot202=False,
#     plot204=False,
#     plot206=True,
#     plot207=False,
#     plot208=False,
#     plot232=False,
#     plot238=False,
#     showplot=True,
#     saveplots=True)
# a.print_RawData()

for file in loaded_files:
    file.split_blank_sample(print_separeted_files=True, new_split_points=[1, 15, 22, 60])
