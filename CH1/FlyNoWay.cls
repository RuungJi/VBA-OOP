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
'@ModuleDescription("Duck �ൿ ����Ŭ�����Դϴ�")
'@Interface
Option Explicit
Implements IFlyBehavior

Private Sub IFlyBehavior_Fly()
    ' �� �� ����!
    Debug.Print "�� �� �����!"
End Sub