import json
import pandas as pd
from collections import defaultdict


"""
輸入單位:
-------------------------
長度: cm
線密度: kg/m

輸出單位:
-------------------------
長度: m
面積: m^2
體積: m^3
重量: kg

"""


# 全域變數:
STEEL_DATA  = "./steel_data.json"
BEAM_DATA   = "./beam_data.json"
BEAM_COUNT  = "./beam_count.json"
OUTPUT_PATH = "./output.xlsx"


class Beam:
    def __init__(self, 長度, 截面, 主筋, 腰筋, 箍筋):
        self.長度 = 長度
        self.截面 = 截面
        self.主筋 = 主筋
        self.腰筋 = 腰筋
        self.箍筋 = 箍筋
    
    
    # 計算單一樑中的主筋總長度:
    def CountReinforcing(self, steels):
        usedSteel = defaultdict(float)
        reinforcing = self.主筋
        beamLen = self.長度
    
        # 中主筋:
        for i in '上下':
            for steel in reinforcing[i +'中']:
                steelNum = f"#{steel['號數']}"
                if steel['同樑長']:
                    usedSteel[steelNum] += steel['數目'] * (beamLen + 2 * steels[steelNum]['樑板牆搭接長'])
                else:
                    isSingleSide = False
                    maxSideLen = 0.
                    for sideReinforcing in reinforcing[i + '左右']:
                        isSingleSide = isSingleSide or sideReinforcing['單邊']
                        if sideReinforcing['長度'] > maxSideLen:
                            maxSideLen = sideReinforcing['長度']
    
                    if isSingleSide:
                        usedSteel[steelNum] += steel['數目'] * (beamLen - maxSideLen     + 2 * steels[steelNum]['樑板牆搭接長'])
                    else:
                        usedSteel[steelNum] += steel['數目'] * (beamLen - maxSideLen * 2 + 2 * steels[steelNum]['樑板牆搭接長'])
        
        # 左右主筋:
        for i in '上下':
            for steel in reinforcing[i + '左右']:
                steelNum = f"#{steel['號數']}"
                if steel['單邊']:
                    usedSteel[steelNum] += steel['數目'] * (steel['長度'] + 2 * steels[steelNum]['樑板牆搭接長'])
                else:
                    usedSteel[steelNum] += steel['數目'] * (steel['長度'] + 2 * steels[steelNum]['樑板牆搭接長']) * 2
        
        return usedSteel
    
    
    # 計算單一樑中的箍筋總長度:
    def CountStirrup(self):
        usedSteel = defaultdict(float)
        b, h = self.截面
        stirrupLen = (b + h) * 2
        beamLen = self.長度
    
        # 左右箍筋:
        maxSideSpace = 0.
        for sideStirrup in self.箍筋['左右']:
            steelNum = f"#{sideStirrup['號數']}"
            steelCount = sideStirrup['數目']
            usedSteel[steelNum] += (steelCount * 2) * stirrupLen
            sideSpace = steelCount * sideStirrup['間距']
            if sideSpace > maxSideSpace:
                maxSideSpace = sideSpace
        
        # 中箍筋:
        centerSpace = beamLen - 2 * maxSideSpace
        for centerStirrup in self.箍筋['中']:
            steelNum = f"#{centerStirrup['號數']}"
            usedSteel[steelNum] += (centerSpace / centerStirrup['間距']) * stirrupLen
    
        return usedSteel
    
    
    # 計算單一樑中的腰筋總長度:
    def CountWaist(self, steels):
        usedSteel = defaultdict(float)
        beamLen = self.長度
        for waist in self.腰筋:
            waistNum = f"#{waist['號數']}"
            extend = steels[waistNum]['樑板牆搭接長']
            usedSteel[waistNum] += waist['數目'] * (beamLen + extend * 2.)
    
        return usedSteel
    
    
    # 計算單一樑中的混凝土體積:
    def CountConcrete(self):
        return self.截面[0] * self.截面[1] * self.長度
    
    
    # 計算單一樑所要用的模板面積:
    def CountTemplate(self):
        return self.截面[1] * self.長度 * 2


# 按照所提供的鋼筋線密度，將鋼筋長度轉換至鋼筋重量:
def ConvertSteelLengthToWeight(steelLength, steelData):
    steelWeight = defaultdict(float)
    steelLength = {steel: length * 0.01 for steel, length in steelLength.items()}
    for steelNum, steelLen in steelLength.items():
        steelWeight[steelNum] = steelData[steelNum]['單位重量'] * steelLen
    
    return steelWeight


# 將兩鋼筋長度算量的table相加:
def MergeSteelLength(usedSteel0, usedSteel1):
    useSteel = defaultdict(float)
    for steel, length in usedSteel0.items():
        useSteel[steel] += length
    
    for steel, length in usedSteel1.items():
        useSteel[steel] += length
    
    return useSteel


#-------------------------------------------------------------------------------------------------------------

# 讀取.json檔案:
def Init(steelJson, beamJson, beamCountJson):
    with open(steelJson, 'r', encoding='utf-8') as f1:
        with open(beamJson, 'r', encoding='utf-8') as f2:
            with open(beamCountJson, 'r', encoding='utf-8') as f3:
                return json.load(f1), json.load(f2), json.load(f3)


# 對所有樑做數量計算並總和:
def CountBeam(beamData, steelData, beamCounts):
    usedSteel, usedConcrete, usedTemplate = defaultdict(float), 0, 0
    for beamCount in beamCounts.values():
        for beamNum, count in beamCount.items():
            for beamLen in count:
                beam = Beam(beamLen, **beamData[beamNum])

                # 統計鋼筋總長度:
                usedSteel = MergeSteelLength(usedSteel, beam.CountReinforcing(steelData))
                usedSteel = MergeSteelLength(usedSteel, beam.CountWaist(steelData))
                usedSteel = MergeSteelLength(usedSteel, beam.CountStirrup())
                
                # 統計總混凝土:
                usedConcrete += beam.CountConcrete()

                # 統計模板面積:
                usedTemplate += beam.CountTemplate()
    
    # 轉換單位:
    usedSteel = ConvertSteelLengthToWeight(usedSteel, steelData)
    usedConcrete *= 1e-6
    usedTemplate *= 1e-4
    
    return dict(usedSteel), usedConcrete, usedTemplate


# 輸出數量計算結果至excel:
def WriteOutput(usedSteel, usedConcrete, usedTemplate, path):
    count, units = {}, {}
    count['混凝土總體積'], units['混凝土總體積'] = usedConcrete, "m^3"
    count['模板總面積'], units['模板總面積'] = usedTemplate, "m^2"
    for steelNum, weight in usedSteel.items():
        count[f"{steelNum}鋼筋總重量"] = weight
        units [f"{steelNum}鋼筋總重量"] = "m"
    
    result = pd.DataFrame({"數值": count, "單位": units})
    result.index.name = "品類"
    result.sort_index(kind="mergesort", inplace=True)
    result.to_excel(path)
    return result


# 主程式:
def Main():
    steels, beams, beamCount = Init(STEEL_DATA, BEAM_DATA, BEAM_COUNT)
    usedSteel, usedConcrete, usedTemplate = CountBeam(beams, steels, beamCount)
    return WriteOutput(usedSteel, usedConcrete, usedTemplate, OUTPUT_PATH)


if __name__ == '__main__':
    result = Main()











