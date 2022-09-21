VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WeatherData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("CH2.Weather")
'@ModuleDescription("WeatherData Concrete Class")
Option Explicit
Implements ISubject
Private Type TWeatherData
    observers As Collection
    temperature As Single
    humidity As Single
    pressure As Single
End Type
Private this As TWeatherData

Private Sub Class_Initialize()
    Set this.observers = New Collection
End Sub

Public Function ISubject_registerObserver(o As IObserver)
    this.observers.Add o
End Function

Public Function ISubject_removeObserver(o As IObserver)
    this.observers.Remove o
End Function

Public Sub ISubject_notifyObservers()
    Dim i As Byte
    For i = 1 To this.observers.Count
        Call this.observers(i).update
    Next i
End Sub

Public Sub measurementsChanged()
    ISubject_notifyObservers
End Sub

Public Function setMeasurements(temperature As Single, humidity As Single, pressure As Single)
    this.temperature = temperature
    this.humidity = humidity
    this.pressure = pressure
    measurementsChanged
End Function

Public Function getTemperature()
    getTemperature = this.temperature
End Function

Public Function getHumidity()
    getHumidity = this.humidity
End Function

Public Function getPressure()
    getPressure = this.pressure
End Function