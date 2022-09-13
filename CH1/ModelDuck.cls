VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModelDuck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("CH1.Duck")
'@ModuleDescription("모형오리 클래스입니다")
Option Explicit
Implements IDuck

Dim QuackBehavior As IQuackBehavior
Dim FlyBehavior As IFlyBehavior

Private Sub Class_Initialize()
    ' 모형 오리는 바닥에서 삻을 시작합니다. 아직 날 수 없어요.
    Set FlyBehavior = New FlyNoWay
    Set QuackBehavior = New quack
End Sub

Public Sub IDuck_performQuack()
    QuackBehavior.quack
End Sub

Public Sub IDuck_swim()
End Sub

Public Sub IDuck_display()
End Sub

Public Sub IDuck_performFly()
    FlyBehavior.fly
End Sub

Public Function IDuck_setQuackBehavior(qb As IQuackBehavior)
    Set QuackBehavior = qb
End Function

Public Function IDuck_setFlyBehavior(fb As IFlyBehavior)
    Set FlyBehavior = fb
End Function
