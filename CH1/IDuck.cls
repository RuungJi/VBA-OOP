VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDuck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("CH1.Duck")
'@ModuleDescription("Duck 인터페이스입니다")
Option Explicit

Public Sub performQuack()
    '오리가 소리냅니다
    Dim behavior As IQuackBehavior
    behavior.quack
End Sub

Public Sub swim()
    '오리가 헤엄칩니다
End Sub

Public Sub display()
    '오리를 보여줍니다
End Sub

Public Sub performFly()
    '오리가 날아다닙니다
    Dim behavior As IFlyBehavior
    behavior.fly
End Sub
