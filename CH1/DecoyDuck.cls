VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DecoyDuck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("CH1.Duck")
'@ModuleDescription("고무오리 클래스입니다")
Option Explicit
Implements IDuck

Public Sub IDuck_performQuack()
    ' 가짜 오리는 아무 소리도 낼 수 없습니다
End Sub

Public Sub IDuck_swim()
    ' 가짜오리는 물에 뜨지 않습니다
End Sub

Public Sub IDuck_display()
    Debug.Print "가짜오리입니다!"
End Sub

Public Sub IDuck_performFly()
    ' 가짜오리는 날 수 없습니다
End Sub
