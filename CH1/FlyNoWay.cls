VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FlyNoWay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("CH1.Duck.Behavior")
'@ModuleDescription("Duck 행동 구현클래스입니다")
'@Interface
Option Explicit
Implements IFlyBehavior

Private Sub IFlyBehavior_Fly()
    ' 날 수 없음!
    Debug.Print "날 수 없어요!"
End Sub
