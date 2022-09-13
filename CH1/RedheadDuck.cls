VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RedHeadDuck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("CH1.Duck")
'@ModuleDescription("붉은머리오리 클래스입니다")
Option Explicit

Public Sub quack()
    Dim behavior As IQuackBehavior
    Set behavior = New quack
    behavior.quack
End Sub

Public Sub swim()
    Debug.Print "붉은머리오리는 헤엄칩니다"
End Sub

Public Sub display()
     Debug.Print "붉은머리오리입니다!"
End Sub

Public Sub fly()
    Dim behavior As IFlyBehavior
    Set behavior = New FlyWithWings
    behavior.fly
End Sub
