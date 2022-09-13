VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Quack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("CH1.Duck.Behavior")
'@ModuleDescription("Duck 행동 구현클래스입니다")
'@Interface
Option Explicit
Implements IQuackBehavior

Private Sub IQuackBehavior_quack()
    ' 꽥꽥 소리
    Debug.Print "꽥꽥!"
End Sub
