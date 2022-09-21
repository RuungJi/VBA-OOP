VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StatisticsDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("CH2.Weather")
'@ModuleDescription("StatisticsDisplay Concrete Class")
Option Explicit
Implements IObserver
Implements IDisplayElement
Private Type TStatisticsDisplay
    maxTemp As Single
    minTemp As Single
    tempSum As Single
    numReadings As Integer
    weather As weatherData
End Type
Private this As TStatisticsDisplay

Private Sub Class_Initialize()
    this.maxTemp = 0
    this.minTemp = 100
End Sub

Public Function create(weather As ISubject)
    Set this.weather = weather
    this.weather.ISubject_registerObserver Me
End Function

Public Function IObserver_update()
    Dim temp As Single
    temp = this.weather.getTemperature
    this.tempSum = this.tempSum + temp
    this.numReadings = this.numReadings + 1

    If temp > this.maxTemp Then
        this.maxTemp = temp
    End If

    If temp < this.minTemp Then
        this.minTemp = temp
    End If

    IDisplayElement_display
End Function

Public Sub IDisplayElement_display()
    Debug.Print "평균/최고/최저 온도: " & this.tempSum / this.numReadings & "/" & this.maxTemp & "/" & this.minTemp
End Sub
