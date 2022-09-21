Attribute VB_Name = "main"
'@Folder("CH2.Weather")
Option Explicit

Sub WeatherStation()
    Dim weather As weatherData
    Dim currentDS As New CurrentConditionsDisplay
    Dim statisticsDS As New StatisticsDisplay
    Dim forecaseDS As New ForecastDisplay
    Set weather = New weatherData
    currentDS.create weather
    statisticsDS.create weather
    forecaseDS.create weather

    weather.setMeasurements 22, 35, 31
    Debug.Print ""
    weather.setMeasurements 26, 40, 28
    Debug.Print ""
    weather.setMeasurements 18, 50, 28
    Debug.Print ""
End Sub