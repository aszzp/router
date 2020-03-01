from win32api import GetSystemMetrics
from IPython.display import Image, display, SVG
import os as os
import comtypes
from comtypes.client import CreateObject
from comtypes.gen import STKUtil
from comtypes.gen import STKObjects
import numpy as np
import pandas as pd

def create_orbit(name='MySatellite',meanmotion=10,eccentricity=0,inclination=90,argofperigee=0,ascnode=0,location=0):
    root.UnitPreferences.Item('TimeUnit').SetCurrentUnit('day') # 输出星历的间隔最大时长为1天，所以单位不妥会报错
    satellite = sc.Children.New(18, name) # eSatellite
    satellite_i = satellite.QueryInterface(STKObjects.IAgSatellite)
    satellite_i.SetPropagatorType(1)
    propagator = satellite_i.Propagator
    propagator_i = propagator.QueryInterface(STKObjects.IAgVePropagatorJ2Perturbation)
    # propagator_i.InitialState.Epoch="08 Jun 2016 15:14:26"
    # propagator_i.step = 1 # 输出星历点的时间间隔，其默认值是60秒

    # IAgSatellite satellite: Satellite object
    keplerian = propagator_i.InitialState.Representation.ConvertTo(STKUtil.eOrbitStateClassical) # eOrbitStateClassical, Use the Classical Element interface
    keplerian_i = keplerian.QueryInte
    rface(STKObjects.IAgOrbitStateClassical)
    # keplerian_i.CoordinateSystemType = STKObjects.eCoordinateSystemJ2000

    keplerian_i.SizeShapeType = STKObjects.eSizeShapeMeanMotion
    keplerian_i.LocationType = STKObjects.eLocationMeanAnomaly
    keplerian_i.Orientation.AscNodeType = STKObjects.eAscNodeRAAN

    root.UnitPreferences.Item('AngleUnit').SetCurrentUnit('revs')
    root.UnitPreferences.Item('TimeUnit').SetCurrentUnit('day')

    # 定义轨道的形状和大小
    keplerian_i.SizeShape.QueryInterface(STKObjects.IAgClassicalSizeShapeMeanMotion).MeanMotion = meanmotion # 一天绕地多少圈，最大值17
    keplerian_i.SizeShape.QueryInterface(STKObjects.IAgClassicalSizeShapeMeanMotion).Eccentricity = eccentricity # 偏心率，[0,1)，0为标准圆

    # 定义轨道在空间的倾角
    root.UnitPreferences.Item('AngleUnit').SetCurrentUnit('deg')
    root.UnitPreferences.Item('TimeUnit').SetCurrentUnit('sec')
    keplerian_i.Orientation.Inclination = inclination # 轨道倾角：与赤道平面的夹角
    keplerian_i.Orientation.ArgOfPerigee = argofperigee # 近地点幅角：升交点与近地点夹角
    keplerian_i.Orientation.AscNode.QueryInterface(STKObjects.IAgOrientationAscNodeRAAN).Value = ascnode # 升交点赤经：赤道平面春分点向右与升交点夹角

    # 定义卫星在轨道中的位置
    keplerian_i.Location.QueryInterface(STKObjects.IAgClassicalLocationMeanAnomaly).Value = location #平近点角：卫星从近地点开始按平均轨道角速度运动转过的角度

    propagator_i.InitialState.Representation.Assign(keplerian)
    # 启动卫星
    propagator_i.Propagate()
    return satellite


app=CreateObject("STK11.Application")
# app.Visible= True
# app.UserControl= True

root=app.Personality2
root.NewScenario("get_orbit")
sc=root.CurrentScenario
sc2=sc.QueryInterface(STKObjects.IAgScenario)
sc2.StartTime = '10 Jan 2020 04:00:00.000'
sc2.StopTime = '20 Jan 2020 04:00:00.000'

df = pd.DataFrame()
for i in range(10):
    ranum = np.random.rand()
    meanmotion = 7*ranum+7
    eccentricity = ranum
    inclination = 180*ranum
    argofperigee = 360*ranum
    ascnode = 360*ranum

    for t in range(6):
        location = t*60
#         print("{}-{}".format(i,t),meanmotion,eccentricity,inclination,argofperigee,ascnode,location)
        satellite = create_orbit("{}-{}".format(i,t),meanmotion,eccentricity,inclination,argofperigee,ascnode,location)

        root.UnitPreferences.Item('DateFormat').SetCurrentUnit('EpSec')
        root.UnitPreferences.Item('Distance').SetCurrentUnit('km')

        satPosDP = satellite.DataProviders.Item('Cartesian Position')
        satPosDP = satPosDP.QueryInterface(STKObjects.IAgDataProviderGroup)
        satPosDP = satPosDP.Group.Item('J2000')
        satPosDP = satPosDP.QueryInterface(STKObjects.IAgDataPrvTimeVar)
        satPosDP = satPosDP.Exec(sc2.StartTime,sc2.StopTime,60)
        # velResult=cartVelJ2000TimeVar.ExecElements(sc2.StartTime,sc2.StopTime,60,rptElements) # 60为生成数据的时间间隔，单位sec
        time = satPosDP.DataSets.Item(0).GetValues()
        x = satPosDP.DataSets.Item(1).GetValues()
        y = satPosDP.DataSets.Item(2).GetValues()
        z = satPosDP.DataSets.Item(3).GetValues()
        df["time"] = time
        df["{}-{}-x".format(i,t)] = x
        df["{}-{}-y".format(i,t)] = y
        df["{}-{}-z".format(i,t)] = z

df.to_csv('orbit.csv', sep=',', header=True, index=True)