VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MuteQuack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("CH1.Duck.Behavior")
'@ModuleDescription("Duck �ൿ ����Ŭ�����Դϴ�")
'@Interface
Option Explicit
Implements IQuackBehavior

Private Sub IQuackBehavior_quack()
    ' �Ҹ��� �ȳ���
    Debug.Print "<<..!..>>"
End Sub