VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MallardDuck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("CH1.Duck")
'@ModuleDescription("물오리 클래스입니다")
Option Explicit
Implements IDuck

Public Sub IDuck_performQuack()
    Dim behavior As IQuackBehavior
    Set behavior = New quack
    behavior.quack
End Sub

Public Sub IDuck_swim()
    Debug.Print "물오리는 헤엄칩니다"
End Sub

Public Sub IDuck_display()
    Debug.Print "물오리입니다!"
End Sub

Public Sub IDuck_performFly()
    Dim behavior As IFlyBehavior
    Set behavior = New FlyWithWings
    behavior.fly
End Sub
