VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISubject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("CH2.Weather")
'@ModuleDescription("ISubject Interface")
Option Explicit

Public Function registerObserver(o As IObserver)
End Function

Public Function removeObserver(o As IObserver)
End Function

Public Sub notifyObservers()
End Sub