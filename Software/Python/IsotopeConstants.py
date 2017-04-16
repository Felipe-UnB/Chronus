import math

class IsotopeConstants(object):

    def __init__(self):
        #Constants for Nd
        self._SpikeConcentration_150Nd=0.01605 #micromols/g
        self._AtomicMass_Nd=144.24 #mol/g
        self._NaturalAbundance_150Nd=0.056251 #, Russ et al, 1971; Wasserburg et al. 1981
        self._SpikeRatio_146_150Nd=0.00446707 #As presented in UnB's spreadsheet
        self._SpikeRatio_146_144Nd=0.8845800
        self._SpikeRatio_146_143Nd=1.79014169
        self._NaturalRatio_146_150Nd=3.05352835 #Wasserburg et al. 1981
        self._NaturalRatio_146_144Nd = 0.721869989

        #Constants for Sm
        self._SpikeConcentration_149Sm = 0.01334888  # micromols/g
        self._AtomicMass_Sm = 150.35  # mol/g
        self._NaturalAbundance_149Sm = 13.8504  # %
        self._SpikeRatio_147_149Sm = 0.0033163
        self._SpikeRatio_147_152Sm = 0.9397
        self._SpikeRatio_149_152Sm = 249.5
        self._NaturalRatio_147_149Sm = 1.088396 #Wasserburg et al. 1981
        self._NaturalRatio_147_152Sm = 0.5651235 #From the lab spreadsheet
        self._NaturalRatio_149_152Sm = 0.519226

        #Constants for U-Pb
        self._DecayConstant_235U_207Pb = 9.8485 * math.pow(10,-10) # /yr, Jaffey et al (1971)
        self._DecayConstant_238U_206Pb = 1.55125 * math.pow(10, -10)  # /yr Jaffey et al (1971)
        self.NaturalRatio_238U_235U = 137.88 # Shields (1960)

        #List of masses
        self._NdMasses_list=[143,146,144,150]
        self._NdRatios_list = ["143/146", "144/146", "150/146"]

    # Nd --------------------------
    def SpikeConcentration_150Nd(self):
        return self._SpikeConcentration_150Nd

    def AtomicMass_Nd(self):
        return self._AtomicMass_Nd

    def NaturalAbundance_150Nd(self):
        return self._NaturalAbundance_150Nd

    def SpikeRatio_146_150Nd(self):
        return self._SpikeRatio_146_150Nd

    def SpikeRatio_146_144Nd(self):
        return self._SpikeRatio_146_144Nd

    def SpikeRatio_146_143Nd(self):
        return self._SpikeRatio_146_143Nd

    def NaturalRatio_146_150Nd(self):
        return self._NaturalRatio_146_150Nd

    def NaturalRatio_146_144Nd(self):
        return self._NaturalRatio_146_144Nd

    #Sm --------------------------
    def SpikeConcentration_149Sm(self):
        return self._SpikeConcentration_149Sm

    def AtomicMass_Sm(self):
        return self._AtomicMass_Sm

    def NaturalAbundance_149Sm(self):
        return self._NaturalAbundance_149Sm

    def SpikeRatio_147_149Sm(self):
        return self._SpikeRatio_147_149Sm

    def SpikeRatio_147_152Sm(self):
        return self._SpikeRatio_147_152Sm

    def SpikeRatio_149_152Sm(self):
        return self._SpikeRatio_149_152Sm

    def NaturalRatio_147_149Sm(self):
        return self._NaturalRatio_147_149Sm

    def NaturalRatio_147_152Sm(self):
        return self._NaturalRatio_147_152Sm

    def NaturalRatio_149_152Sm(self):
        return self._NaturalRatio_149_152Sm

    #Constants for U-Pb
    def DecayConstant_235U_207Pb(self):
        return self._DecayConstant_235U_207Pb

    def DecayConstant_238U_206Pb(self):
        return self._DecayConstant_238U_206Pb

    def NaturalRatio_238U_235Pb(self):
        return self.NaturalRatio_238U_235U

    #Lists
    def NdMasses_list(self):
        return self._NdMasses_list

    def NdRatios_list(self):
        return self._NdRatios_list

    def __repr__(self):
        from pprint import pformat
        return "<" + type(self).__name__ + "> " + pformat(vars(self), indent=4, width=1)

Cnsts = IsotopeConstants()
# print(Cnsts.DecayConstant_235U_207Pb()-Cnsts.DecayConstant_238U_206Pb())