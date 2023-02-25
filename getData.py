import xlwings as xw
import requests
from datetime import datetime, timedelta


timeDeltaDays = 7
lastDateAvailable = datetime.today() - timedelta(days=timeDeltaDays)
startDateRage = lastDateAvailable - timedelta(days=timeDeltaDays-1)
todayData = datetime.today()
lastFutureData = todayData + timedelta(days=timeDeltaDays-1)

excellCellDataReference = {
    'precipitation_sum': 'Data_PrecipitazioniTot',
    'temperature_2m_max': 'Data_Tmax',
    'temperature_2m_min': 'Data_Tmin',
    'temperature_2m_mean': 'Data_Tmed',
    'relativehumidity_2m_mean' : 'Data_Umidità',
    # 'incident_radiation_sum' : 'Data_RadiazioneSolareInc',
    'windspeed_10m_mean' : 'Data_VelocitàVento',
    'et0_fao_evapotranspiration' : 'Data_ET0_fao',
    'time': 'Data_Giorno',
}

excelCellPosition = {
    "precipitation_sum" : 6,
    "temperature_2m_max" : 8,
    "temperature_2m_mean" : 9,
    "temperature_2m_min" : 10,
    "et0_fao_evapotranspiration" : 13,
    "time": 11,
}

# latitude = 45.70
# longitude = 9.67

def getHistoricData(latitude, longitude):
    openMeteoApiURL = f"""https://archive-api.open-meteo.com/v1/archive?latitude={latitude}&longitude={longitude}&start_date={startDateRage.year}-{startDateRage.month:02d}-{startDateRage.day:02d}&end_date={lastDateAvailable.year}-{lastDateAvailable.month:02d}-{lastDateAvailable.day:02d}&models=best_match&hourly=relativehumidity_2m,windspeed_10m&daily=temperature_2m_max,temperature_2m_min,temperature_2m_mean,precipitation_sum,et0_fao_evapotranspiration,shortwave_radiation_sum&timezone=Europe%2FBerlin"""
    r = requests.get(openMeteoApiURL)
    dataDaily = r.json()['daily']
    dataHourly = r.json()['hourly']

    workbook = xw.Book.caller() 
    sheet = workbook.sheets['Dati meteo']

    dataDaily['windspeed_10m_mean'] = getMeanValues(getHourlyDataSets(dataHourly['windspeed_10m']))
    dataDaily['relativehumidity_2m_mean'] = getMeanValues(getHourlyDataSets(dataHourly['relativehumidity_2m']))

    for dataKey, excelCellRange in excellCellDataReference.items():
        values = dataDaily[dataKey]
        sheet.range(excelCellRange).value = values

def getForecastData(latitude, longitude):
    openMeteoApiURL = f"""https://api.open-meteo.com/v1/forecast?latitude={latitude}&longitude={longitude}&start_date={todayData.year}-{todayData.month:02d}-{todayData.day:02d}&end_date={lastFutureData.year}-{lastFutureData.month:02d}-{lastFutureData.day:02d}&models=best_match&hourly=temperature_2m,relativehumidity_2m,windspeed_10m,direct_radiation,diffuse_radiation&daily=temperature_2m_max,temperature_2m_min,precipitation_sum,et0_fao_evapotranspiration&timezone=Europe%2FBerlin"""
    r = requests.get(openMeteoApiURL)
    dataDaily = r.json()['daily']
    dataHourly = r.json()['hourly']

    workbook = xw.Book.caller() 
    sheet = workbook.sheets['Dati meteo']

    dataDaily['windspeed_10m_mean'] = getMeanValues(getHourlyDataSets(dataHourly['windspeed_10m']))
    dataDaily['relativehumidity_2m_mean'] = getMeanValues(getHourlyDataSets(dataHourly['relativehumidity_2m']))
    dataDaily['temperature_2m_mean'] = getMeanValues(getHourlyDataSets(dataHourly['temperature_2m']))

    for dataKey, excelCellRange in excellCellDataReference.items():
        values = dataDaily[dataKey]
        sheet.range(excelCellRange).value = values

    # incidentRadiationValues = []
    # # Sum diffuse and direct radiation values to get the incident radiation
    # for index, value in enumerate(dataHourly['diffuse_radiation']):
    #     incidentRadiationValues.append(value + dataHourly['direct_radiation'][index])
        
    # dataDaily['incident_radiation_sum'] = getMeanValues(getHourlyDataSets(incidentRadiationValues))

def getMeanTemp(temperaturesList):
    T_mean = []
    T_sum, counter = 0, 0
    dataPerDay = 24

    for temperature in temperaturesList:
        T_sum += temperature
        counter += 1
        if counter == 24:
            T_mean.append(T_sum/dataPerDay)
            T_sum, counter = 0, 0

    return T_mean

def getHourlyDataSets(rawdataList: list):
    dataSets = []
    dataPerDay = 24

    newSet = []
    counter = 0
    for value in rawdataList:
        newSet.append(value)
        counter += 1
        if counter == dataPerDay:
            dataSets.append(newSet)
            newSet = []    
            counter = 0

    return dataSets


def getMeanValues(valueList: list):
    means = []
    for set in valueList:
        valuesSum = 0
        for values in set:
            valuesSum += values
        means.append(valuesSum/len(set))
    return means
