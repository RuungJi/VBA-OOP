VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ForecastDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("CH2.Weather")
'@ModuleDescription("ForecastDisplay Concrete Class")
Option Explicit
Implements IObserver
Implements IDisplayElement
Private Type TForecastDisplay
    currentPressure As Single
    lastPressure As Single
    weather As ISubject
End Type
Private this As TForecastDisplay

Private Sub Class_Initialize()
    this.currentPressure = 30
End Sub

Public Function create(weather As ISubject)
    Set this.weather = weather
    this.weather.registerObserver New ForecastDisplay
End Function

Public Function IObserver_update(temperature As Single, humidity As Single, pressure As Single)
    this.lastPressure = this.currentPressure
    this.currentPressure = pressure
    IDisplayElement_display
End Function

Public Sub IDisplayElement_display()
    If this.currentPressure > this.lastPressure Then
        Debug.Print "날씨가 좋아질 것으로 보입니다."
    ElseIf this.currentPressure = this.lastPressure Then
        Debug.Print "지금과 비슷할 것으로 보입니다."
    Else
        Debug.Print "쌀쌀하며 비가 올 것으로 보입니다"
    End If
End Sub