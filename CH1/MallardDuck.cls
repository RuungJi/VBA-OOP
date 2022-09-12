VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MallardDuck"
'@Folder("CH1.Duck")
'@ModuleDescription("물오리 클래스입니다")
Option Explicit
Implements IDuck

Public Sub IDuck_quack()
    Debug.Print "저는 물오리입니다"
End Sub

Public Sub IDuck_swim()
    Debug.Print "물오리는 헤엄칩니다"
End Sub

Public Sub IDuck_display()
    Debug.Print "물오리입니다!"
End Sub

Public Sub IDuck_fly()
    Debug.Print "물오리가 날아다닙니다"
End Sub
