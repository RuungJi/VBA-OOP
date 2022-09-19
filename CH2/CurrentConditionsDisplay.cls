VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CurrentConditionsDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("CH2.Weather")
'@ModuleDescription("CurrentConditionsDisplay Concrete Class")
Option Explicit
Implements IObserver
Implements IDisplayElement
Private Type TCurrentConditionsDisplay
    temperature As Single
    humidity As Single
    weather As ISubject
End Type
Private this As TCurrentConditionsDisplay

Public Function create(weather As ISubject)
    Set this.weather = weather
    this.weather.registerObserver New CurrentConditionsDisplay
End Function

Public Function IObserver_update(temperature As Single, humidity As Single, pressure As Single)
    this.temperature = temperature
    this.humidity = humidity
    IDisplayElement_display
End Function

Public Sub IDisplayElement_display()
    Debug.Print "현재 상태: 온도 " & this.temperature & "도, 습도 " & this.humidity & "%"
End Sub