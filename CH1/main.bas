Attribute VB_Name = "main"
'@Folder("CH1.Duck")
Option Explicit

Sub DuckSimulator()
    Dim Mallard As IDuck
    Set Mallard = New MallardDuck
    Mallard.performQuack
    Mallard.swim
    Mallard.display
    Mallard.performFly

    Debug.Print ""

    Dim RedHead As New RedHeadDuck
    RedHead.quack
    RedHead.swim
    RedHead.display
    RedHead.fly

    Debug.Print ""
    
    Dim Rubber As IDuck
    Set Rubber = New RubberDuck
    Rubber.performQuack
    Rubber.swim
    Rubber.display
    Rubber.performFly
    
    Debug.Print ""

    Dim Decoy As IDuck
    Set Decoy = New DecoyDuck
    Decoy.performQuack
    Decoy.swim
    Decoy.display
    Decoy.performFly
End Sub