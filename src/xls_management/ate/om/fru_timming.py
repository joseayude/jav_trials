import pandas as pd


class FRUTiming:
#Option Explicit
#

    def __init(
        self,
        #feature:str,   #Public Feature As String
        #reifegrad:str, #Public Reifegrad As String
        #umsetzer:str,  #Public Umsetzer As String
        #i_stufe:str,   #Public IStufe As String
        columns:pd.DataFrame,
        row:int,
        fru_timming_index: dict[str,"FRUTiming"],
    ):
        #self.feature = feature
        #self.reifegrad = reifegrad
        #self.umsetzer = umsetzer
        #self.i_stufe = i_stufe
#       'Neues FRUTiming anlegen
#       Set FRUTiming = New FRUTiming
#       'Feature einlesen
#       FRUTiming.Feature = rngFRUTimingAttribute(1).Offset(lngZeile, 0).Value
        self.feature = str(columns['FeatureName'][row])
#       'Reifegrad einlesen
#       FRUTiming.Reifegrad = rngFRUTimingAttribute(2).Offset(lngZeile, 0).Value
        self.reifegrad = str(columns['Reifegrad'][row])
#       'Umsetzer einlesen
#       FRUTiming.Umsetzer = rngFRUTimingAttribute(3).Offset(lngZeile, 0).Value
        self.umsetzer = str(columns['Umsetzer'][row])
#       'I-Stufe einlesen
#       FRUTiming.IStufe = rngFRUTimingAttribute(4).Offset(lngZeile, 0).Value
        self.i_stufe = str(columns['FE_Meilenstein'][row])
#       'FRU-Key erzeugen
#       strFRUKey = FRUTiming.Feature & FRUTiming.Reifegrad & FRUTiming.Umsetzer
        fru_key = f'{self.feature}{self.reifegrad}{self.umsetzer}'
#       'Erfasstes FRU-Timing in globaler FRUTiming-Liste hinzufügen
#       FRUTimingList.Add Item:=FRUTiming, Key:=strFRUKey
        fru_timming_index[fru_key] = self
