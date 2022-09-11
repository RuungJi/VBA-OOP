Attribute VB_Name = "main"
'@Folder("CH1.Duck")
Option Explicit

Sub DuckSimulator
    Dim Mallard As IDuck
    Set Mallard = New MallardDuck
    Mallard.quack
    Mallard.swim
    Mallard.display

    Dim RedHead As New RedHeadDuck
    RedHead.quack
    RedHead.swim
    RedHead.display
End Sub