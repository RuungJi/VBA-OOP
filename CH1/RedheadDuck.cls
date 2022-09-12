VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RedHeadDuck"
'@Folder("CH1.Duck")
'@ModuleDescription("붉은머리오리 클래스입니다")
Option Explicit

Public Sub quack()
    Debug.Print "저는 붉은머리오리입니다"
End Sub

Public Sub swim()
    Debug.Print "붉은머리오리는 헤엄칩니다"
End Sub

Public Sub display()
     Debug.Print "붉은머리오리입니다!"
End Sub

Public Sub fly()
    Debug.Print "붉은머리오리가 날아다닙니다"
End Sub
