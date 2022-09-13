VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RubberDuck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("CH1.Duck")
'@ModuleDescription("고무오리 클래스입니다")
Option Explicit
Implements IDuck

Public Sub IDuck_performQuack()
    Debug.Print "삑삑!"
End Sub

Public Sub IDuck_swim()
    Debug.Print "고무오리는 물에 뜹니다"
End Sub

Public Sub IDuck_display()
    Debug.Print "고무오리입니다!"
End Sub

Public Sub IDuck_performFly()
    ' 고무오리는 날 수 없습니다
End Sub
