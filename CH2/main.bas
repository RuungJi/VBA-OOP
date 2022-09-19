Attribute VB_Name = "main"
'@Folder("CH2.Weather")
Option Explicit

Sub WeatherStation()
    Dim weather As New weatherData
    Dim currentDS As New CurrentConditionsDisplay
    Dim statisticsDS As New StatisticsDisplay
    Dim forecaseDS As New ForecastDisplay
    currentDS.create weather
    statisticsDS.create weather
    forecaseDS.create weather

    weather.setMeasurements 20, 35, 31
    Debug.Print ""
    weather.setMeasurements 22, 40, 28
    Debug.Print ""
    weather.setMeasurements 18, 50, 28
    Debug.Print ""
End Sub